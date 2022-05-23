using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using BootlegRealists.Reporting.Enumeration;
using BootlegRealists.Reporting.Extension;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class processes an OpenXML element and returns the PDF chunk
/// </summary>
public static class ChunkProcessor
{
	/// <summary>
	/// Processes the given chunk and applies the given effects and font
	/// </summary>
	/// <param name="docxDocument">Word processing document</param>
	/// <param name="compositeElement">Composite element referencing the chunk</param>
	/// <param name="chunk">Chunk to process</param>
	/// <param name="bold">Bold effect</param>
	/// <param name="italic">Italic effect</param>
	/// <param name="strike">Strike effect</param>
	/// <param name="caps">Caps effect</param>
	/// <param name="underline">Underline effect</param>
	/// <param name="verticalAlignment">Superscript/subscript effect</param>
	/// <param name="fontSize">Font size (in points)</param>
	/// <param name="fontSizeComplexScript">Complex script font size (in points)</param>
	/// <param name="color">Color</param>
	/// <returns>The processed chunk</returns>
	public static Chunk Process(WordprocessingDocument docxDocument, OpenXmlElement compositeElement, Chunk chunk, Bold? bold, Italic? italic,
		Strike? strike, OnOffType? caps, Underline? underline, VerticalAlignment verticalAlignment, float fontSize, float fontSizeComplexScript,
		Color? color)
	{
		if (Converter.OnOffToBool(caps))
			chunk = new Chunk(chunk.Content.ToUpper(CultureInfo.InvariantCulture));

		var baseFont = FontCreator.GetBaseFont(docxDocument, compositeElement, Converter.OnOffToBool(bold), Converter.OnOffToBool(italic), chunk.Content);
		if (baseFont != null)
		{
			var fontType = CodePointRecognizer.GetFontType(chunk.Content[0]);
			float ftSize;
			if (fontSizeComplexScript > 0.0f && fontType.FontType == FontTypeEnum.ComplexScript)
				ftSize = fontSizeComplexScript;
			else if (verticalAlignment.HasOffset())
			{
				var baseFont2 = (BaseFont)baseFont;
				if (verticalAlignment == VerticalAlignment.Superscript)
				{
					ftSize = baseFont2.GetFontDescriptor(BaseFont.SUPERSCRIPT_SIZE, fontSize);
					var offset = baseFont2.GetFontDescriptor(BaseFont.SUPERSCRIPT_OFFSET, fontSize);
					chunk.SetTextRise(offset);
				}
				else 
				{
					ftSize = baseFont2.GetFontDescriptor(BaseFont.SUBSCRIPT_SIZE, fontSize);
					var offset = baseFont2.GetFontDescriptor(BaseFont.SUBSCRIPT_OFFSET, fontSize);
					chunk.SetTextRise(offset);
				}
			}
			else
				ftSize = fontSize;
			chunk.Font = FontFactory.CreateFont(baseFont, ftSize, bold, italic, strike, color);
		}

		if (underline?.Val == null || underline.Val.Value == UnderlineValues.None) return chunk;
		var fntSize = chunk.Font.CalculatedSize;
		chunk.SetUnderline(0.07f * fntSize, -0.2f * fntSize);
		return chunk;
	}
}
