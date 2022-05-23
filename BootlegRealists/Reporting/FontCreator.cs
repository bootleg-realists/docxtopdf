using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;
using BootlegRealists.Reporting.Extension;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class creates a font (PDF) from a OpenXML element.
/// </summary>
public static class FontCreator
{
	/// <summary>
	/// Return iTextSharp.text.pdf.BaseFont base on text and fonts definitions of word file.
	/// The algorithm is as following,
	/// 1. Map text's code point to font type (Ascii/EastAsia/ComplexScript/HighAnsi)
	/// 2. Base on font type, go through Run:rPr > Paragraph:StyleId:rPr > Paragraph:StyleId:pPr > default style >
	/// docDefault:rPr > Run:rPr:hint to find out the font name
	/// 3. According to the font name to create iTextSharp.text.pdf.BaseFont and return
	/// </summary>
	/// <param name="docxDocument">The word document.</param>
	/// <param name="compositeElement">Run object where the text belongs to. It will be used for traversal.</param>
	/// <param name="bold">True if font is bold</param>
	/// <param name="italic">True if font is italic</param>
	/// <param name="text">The text to extract the code point from.</param>
	/// <returns>Return created iBaseFontEx if success, otherwise return null.</returns>
	public static BaseFontEx? GetBaseFont(WordprocessingDocument docxDocument, OpenXmlElement compositeElement, bool bold, bool italic, string text)
	{
		// 1. rFont
		// Run:rPr > Paragraph:StyleId:rPr > Paragraph:StyleId:pPr > default style > docDefault:rPr > Run:rPr:hint > Theme
		// 2. lang (if no rFont)
		// Run:rPr > Paragraph:StyleId:rPr > default style > docDefault:rPr
		var fontName = "";

		var fontType = CodePointRecognizer.GetFontType(text[0]);

		// *************
		// rFont, search Run first
		var runFonts = compositeElement.GetFirstDescendant<RunFonts>();
		if (runFonts != null)
		{
			// Special case handling (seems w:hint only exists in Run)
			if (fontType.UseEastAsiaIfhintIsEastAsia && runFonts.Hint?.Value == FontTypeHintValues.EastAsia)
				fontType.FontType = FontTypeEnum.EastAsian;

			// Search Run's direct formatting
			fontName = GetFontNameFromRunFontsByFontType(runFonts, fontType);

			// Search Run's rStyle
			if (string.IsNullOrEmpty(fontName) && compositeElement is Run r && r.RunProperties?.RunStyle != null)
				fontName = GetFontNameFromRunFontsByFontType(r.RunProperties?.RunStyle?.GetStyleById()?.GetEffectiveElement<RunFonts>(), fontType);
		}

		// No matched RunFonts from Run, search Paragraph pStyle
		if (string.IsNullOrEmpty(fontName) && compositeElement.Parent is Paragraph pg && pg.ParagraphProperties?.ParagraphStyleId != null)
			fontName = GetFontNameFromRunFontsByFontType(pg.ParagraphProperties?.ParagraphStyleId?.GetStyleById()?.GetEffectiveElement<RunFonts>(), fontType);

		// Search in default styles: character
		if (string.IsNullOrEmpty(fontName))
			fontName = GetFontNameFromRunFontsByFontType(docxDocument.MainDocumentPart?.GetDefaultStyle(DefaultStyleType.Character)?.GetFirstDescendant<RunFonts>(), fontType);

		// Default style, paragraph
		if (string.IsNullOrEmpty(fontName))
			fontName = GetFontNameFromRunFontsByFontType(docxDocument.MainDocumentPart?.GetDefaultStyle(DefaultStyleType.Paragraph)?.GetFirstDescendant<RunFonts>(), fontType);

		// Search in docDefault
		if (string.IsNullOrEmpty(fontName))
			fontName = GetFontNameFromRunFontsByFontType(docxDocument.MainDocumentPart?.GetDocDefaults<RunFonts>(DocDefaultsType.Character), fontType);

		// Still can't find, use w:hint
		if (!string.IsNullOrEmpty(fontName) || runFonts == null)
			return GetBaseFontByFontName(docxDocument, fontName, bold, italic, fontType, compositeElement.GetEffectiveElement<Languages>());

		if (runFonts.Hint != null)
			fontName = runFonts.Hint.InnerText ?? "";

		return GetBaseFontByFontName(docxDocument, fontName, bold, italic, fontType,
			compositeElement.GetEffectiveElement<Languages>());
	}

	/// <summary>
	/// Get font name from RunFonts, only from w:ascii(Theme)/w:hAnsi(Theme)/w:cs(Theme)/w:eastAsia(Theme) but not from
	/// w:hint.
	/// </summary>
	/// <param name="runFonts">Wordprocessing.RunFonts object.</param>
	/// <param name="fontType">FontTypeInfo object.</param>
	/// <returns>Return font name if matches, otherwise return null.</returns>
	public static string GetFontNameFromRunFontsByFontType(RunFonts? runFonts, FontTypeInfo fontType)
	{
		if (runFonts == null)
			return "";
		return fontType.FontType switch
		{
			FontTypeEnum.Ascii => runFonts.AsciiTheme != null
				? runFonts.AsciiTheme.InnerText ?? ""
				: runFonts.Ascii?.Value ?? "",
			FontTypeEnum.ComplexScript => runFonts.ComplexScriptTheme != null
				? runFonts.ComplexScriptTheme.InnerText ?? ""
				: runFonts.ComplexScript?.Value ?? "",
			FontTypeEnum.EastAsian => runFonts.EastAsiaTheme != null
				? runFonts.EastAsiaTheme.InnerText ?? ""
				: runFonts.EastAsia?.Value ?? "",
			FontTypeEnum.HighAnsi => runFonts.HighAnsiTheme != null
				? runFonts.HighAnsiTheme.InnerText ?? ""
				: runFonts.HighAnsi?.Value ?? "",
			_ => ""
		};
	}

	/// <summary>
	/// Get Pdf.BaseFont by font name. The font name was extracted from RunFonts.
	/// This method first try to create BaseFont by font name. If the font name is not normal font name and looks like
	/// eastAsia/cs/hAnsi/ascii or majorXX/minorXX, this method will try to search docDefaults and Theme.
	/// If all above procedures can't find font, it will try to use font type and w:lang to find the backup font.
	/// </summary>
	/// <param name="docxDocument">The word document</param>
	/// <param name="fontName">Name of the font</param>
	/// <param name="bold">True if font is bold</param>
	/// <param name="italic">True if font is italic</param>
	/// <param name="fontType">Type of the font</param>
	/// <param name="language">Language complement, will be used in case can't find BaseFont by font name.</param>
	/// <returns>The base font (can be null)</returns>
	public static BaseFontEx? GetBaseFontByFontName(WordprocessingDocument docxDocument, string fontName, bool bold, bool italic, FontTypeInfo fontType, LanguageType? language)
	{
		// Use font name to find font file path
		var result = FontFactory.CreateBaseFont(fontName, bold, italic);
		if (result == null)
		{
			// Cannot find font file path has two possibilities 
			// 1. fontName = (eastAsia/cs/hAnsi/ascii), search docDefault
			var docDefaultRunFonts =
				docxDocument.MainDocumentPart?.GetDocDefaults<RunFonts>(DocDefaultsType.Character);
			if (docDefaultRunFonts != null)
			{
				if (fontName.Contains("eastAsia", StringComparison.OrdinalIgnoreCase) && docDefaultRunFonts.EastAsia != null)
					result = FontFactory.CreateBaseFont(docDefaultRunFonts.EastAsia.Value ?? "", bold, italic);
				else if (fontName.Contains("cs", StringComparison.OrdinalIgnoreCase) && docDefaultRunFonts.ComplexScript != null)
					result = FontFactory.CreateBaseFont(docDefaultRunFonts.ComplexScript.Value ?? "", bold, italic);
				else if (fontName.Contains("hAnsi", StringComparison.OrdinalIgnoreCase) && docDefaultRunFonts.HighAnsi != null)
					result = FontFactory.CreateBaseFont(docDefaultRunFonts.HighAnsi.Value ?? "", bold, italic);
				else if (fontName.Contains("ascii", StringComparison.OrdinalIgnoreCase) && docDefaultRunFonts.Ascii != null)
					result = FontFactory.CreateBaseFont(docDefaultRunFonts.Ascii.Value ?? "", bold, italic);
			}

			// 2. fontName = majorXXX/minorXXX, search Theme
			if (result == null)
				result = FontFactory.CreateBaseFont(docxDocument.MainDocumentPart?.GetThemeFontByType(fontName) ?? "", bold, italic);
		}

		// *************
		// Can't find rFont, search for w:lang
		if (result != null || language == null)
			return result;

		var lang = GetLangFromLanguagesByFontType(language, fontType);

		// map lang (e.g. zh-TW) to script tag (e.g. Hant)
		var scriptTag = LangScriptTag.GetScriptTagByLocale(lang);
		return FontFactory.CreateBaseFont(docxDocument.MainDocumentPart?.GetThemeFontByScriptTag(scriptTag) ?? "", bold, italic);
	}

	static string GetLangFromLanguagesByFontType(LanguageType? langs, FontTypeInfo fontType)
	{
		if (langs == null)
			return "";

		return fontType.FontType switch
		{
			FontTypeEnum.EastAsian => langs.EastAsia?.Value ?? "",
			FontTypeEnum.ComplexScript => langs.Bidi?.Value ?? "",
			_ => langs.Val?.Value ?? ""
		};
	}
}
