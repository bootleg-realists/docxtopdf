using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;
using Color = System.Drawing.Color;

// TODO: handle sectPr w:docGrid w:linePitch (how many lines per page)
// TODO: handle tblLook $17.7.6 (conditional formatting), it will define firstrow/firstcolum/...etc styles in styles.xml

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains main document part extension methods.
/// </summary>
public static class MainDocumentPartExtension
{
	static readonly Dictionary<DefaultStyleType, string> DefaultStyleName = new()
	{
		{ DefaultStyleType.Paragraph, "paragraph" },
		{ DefaultStyleType.Character, "character" },
		{ DefaultStyleType.Table, "table" },
		{ DefaultStyleType.Numbering, "numbering" }
	};

	/// <summary>
	/// Gets the printable width for the given object.
	/// </summary>
	/// <param name="obj">The object to act on.</param>
	/// <returns>The printable width in twips or 0.0f otherwise</returns>
	public static float GetPrintablePageWidth(this MainDocumentPart obj)
	{
		var body = obj.Document.Body;
		var section = body?.Descendants<SectionProperties>().FirstOrDefault();
		if (section == null) return 0.0f;
		var size = section.Descendants<PageSize>().FirstOrDefault();
		if (size == null || size.Width == (object?)null) return 0.0f;
		var margin = section.Descendants<PageMargin>().FirstOrDefault();
		if (margin == null || margin.Left == (object?)null || margin.Right == (object?)null) return 0.0f;
		return size.Width.Value - margin.Left.Value - margin.Right.Value;
	}

	/// <summary>
	/// Gets the default style for the given object and style type.
	/// </summary>
	/// <param name="obj">The object to act on.</param>
	/// <param name="type">Style type to get default style for.</param>
	/// <returns>The default style or null otherwise</returns>
	public static Style? GetDefaultStyle(this MainDocumentPart obj, DefaultStyleType type)
	{
		var stylesX = obj.StyleDefinitionsPart?.Styles;
		if (stylesX == null) return null;
		var styles = stylesX.Descendants<Style>();
		var docDefaults = stylesX.Descendants<DocDefaults>().FirstOrDefault();
		if (docDefaults == null) return null;

		var typeStr = DefaultStyleName[type];
		var result = styles.FirstOrDefault(x =>
		{
			var attrType = x.GetAttribute("type", docDefaults.NamespaceUri);
			if (attrType.Value == null)
				return false;

			if (attrType.Value.IndexOf(typeStr, 0, StringComparison.Ordinal) < 0)
				return false;

			var attrDefault = x.GetAttribute("default", docDefaults.NamespaceUri);
			return attrDefault.Value is "1";
		});
		return result;
	}

	/// <summary>
	/// Get object from docDefaults.
	/// </summary>
	/// <param name="obj">The object to act on.</param>
	/// <param name="type">Default type to get.</param>
	/// <typeparam name="T">Target class type.</typeparam>
	/// <returns>The object or null otherwise</returns>
	public static T? GetDocDefaults<T>(this MainDocumentPart obj, DocDefaultsType type) where T : OpenXmlElement
	{
		var docDefaults = obj.StyleDefinitionsPart?.Styles?.Descendants<DocDefaults>().FirstOrDefault();
		if (docDefaults == null) return default;
		return type switch
		{
			DocDefaultsType.Character => docDefaults.RunPropertiesDefault?.GetFirstDescendant<T>(),
			DocDefaultsType.Paragraph => docDefaults.ParagraphPropertiesDefault?.GetFirstDescendant<T>(),
			_ => default
		};
	}

	/// <summary>
	/// Get Theme font name by script tag (e.g. Hant). Only search from theme>fontScheme>minorFont because majorFont is
	/// meant to be used with Headings (Heading 1, etc.) and minorFont with "normal text".
	/// </summary>
	/// <param name="obj">The object to act on.</param>
	/// <param name="scriptTag"></param>
	/// <returns>The theme font or null otherwise.</returns>
	public static string GetThemeFontByScriptTag(this MainDocumentPart obj, string scriptTag)
	{
		var minorFont = obj.ThemePart?.Theme.ThemeElements?.FontScheme?.MinorFont;
		if (minorFont == null) return "";
		foreach (var fontValue in minorFont.Descendants<SupplementalFont>())
		{
			if (fontValue.Script?.Value?.IndexOf(scriptTag, StringComparison.OrdinalIgnoreCase) >= 0 &&
			    fontValue.Typeface?.HasValue == true)
				return fontValue.Typeface?.Value ?? "";
		}

		return "";
	}

	/// <summary>
	/// Get Theme font name by type. The possible values are majorBidi/minorBidi, majorHAnsi/minorHAnsi,
	/// majorEastAsia/minorEastAsia.
	/// </summary>
	/// <param name="obj">The object to act on.</param>
	/// <param name="fontType">
	/// Font type, the possible values are majorBidi/minorBidi, majorHAnsi/minorHAnsi,
	/// majorEastAsia/minorEastAsia.
	/// </param>
	/// <returns>The theme font or null otherwise</returns>
	public static string GetThemeFontByType(this MainDocumentPart obj, string fontType)
	{
		// http://blogs.msdn.com/b/officeinteroperability/archive/2013/04/22/office-open-xml-themes-schemes-and-fonts.aspx
		// major fonts are mainly for styles as headings, whereas minor fonts are generally applied to body and paragraph text

		var theme = obj.ThemePart;
		if (theme == null) return "";
		if (fontType.Contains("majorBidi", StringComparison.OrdinalIgnoreCase) && theme.Theme.ThemeElements?.FontScheme?.MajorFont?.ComplexScriptFont?.Typeface != null)
			return theme.Theme.ThemeElements?.FontScheme?.MajorFont?.ComplexScriptFont?.Typeface?.Value ?? "";
		if (fontType.Contains("minorBidi", StringComparison.OrdinalIgnoreCase) && theme.Theme.ThemeElements?.FontScheme?.MinorFont?.ComplexScriptFont?.Typeface != null)
			return theme.Theme.ThemeElements?.FontScheme?.MinorFont?.ComplexScriptFont?.Typeface?.Value ?? "";
		if (fontType.Contains("majorHAnsi", StringComparison.OrdinalIgnoreCase) && theme.Theme.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface != null)
			return theme.Theme.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value ?? "";
		if (fontType.Contains("minorHAnsi", StringComparison.OrdinalIgnoreCase) && theme.Theme.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface != null)
			return theme.Theme.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value ?? "";
		if (fontType.Contains("majorEastAsia", StringComparison.OrdinalIgnoreCase) && theme.Theme.ThemeElements?.FontScheme?.MajorFont?.EastAsianFont?.Typeface != null)
			return theme.Theme.ThemeElements?.FontScheme?.MajorFont?.EastAsianFont?.Typeface?.Value ?? "";
		if (fontType.Contains("minorEastAsia", StringComparison.OrdinalIgnoreCase) && theme.Theme.ThemeElements?.FontScheme?.MinorFont?.EastAsianFont?.Typeface != null)
			return theme.Theme.ThemeElements?.FontScheme?.MinorFont?.EastAsianFont?.Typeface?.Value ?? "";

		return "";
	}

	/// <summary>
	/// Get image by r:id
	/// </summary>
	/// <param name="obj">The object to act on.</param>
	/// <param name="id">Relationship ID.</param>
	/// <returns>Return a stream object if found, otherwise return null.</returns>
	public static Stream? GetImageById(this MainDocumentPart obj, string id)
	{
		if (string.IsNullOrEmpty(id)) return null;
		var part = obj.GetPartById(id);

		var img = obj.ImageParts.FirstOrDefault(c => c.Uri.Equals(part.Uri));
		if (img == null) return null;
		var stream = img.GetStream(FileMode.Open, FileAccess.Read);
		return stream;
	}
}
