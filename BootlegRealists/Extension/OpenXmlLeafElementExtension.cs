using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// TODO: handle sectPr w:docGrid w:linePitch (how many lines per page)
// TODO: handle tblLook $17.7.6 (conditional formatting), it will define firstrow/firstcolum/...etc styles in styles.xml

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains open xml leaf element extension methods.
/// </summary>
public static class OpenXmlLeafElementExtension
{
	/// <summary>
	/// Get Wordprocessing.Style by style ID. This method combines the LinkedStyle together if any.
	/// </summary>
	/// <param name="obj">Object to act on.</param>
	/// <returns>Return Wordprocessing.Style object if found otherwise return null.</returns>
	public static Style? GetStyleById(this OpenXmlLeafElement obj)
	{
		var styles = GetStyles(obj);

		string styleId;
		switch (obj)
		{
			case StringType st:
				styleId = st.Val?.Value ?? "";
				break;
			case String253Type st253:
				styleId = st253.Val?.Value ?? "";
				break;
			default:
				return null;
		}

		var style = InnerGetStyleById(styles, styleId);
		if (style == null)
			return null;

		var styleType = style.Type?.Value ?? StyleValues.Character;

		// Retrieve LinkedStyle only for Paragraph & Character
		if (styleType != StyleValues.Paragraph && styleType != StyleValues.Character || style.LinkedStyle == null)
			return style;

		var linkedStyle = InnerGetStyleById(styles, style.LinkedStyle.Val?.Value ?? "");
		if (linkedStyle?.StyleRunProperties == null)
			return style;

		if (styleType == StyleValues.Paragraph)
		{
			style.StyleRunProperties = (StyleRunProperties)linkedStyle.StyleRunProperties.CloneNode(true);
		}
		else
		{
			if (style.StyleRunProperties != null)
				linkedStyle.StyleRunProperties = (StyleRunProperties)style.StyleRunProperties.CloneNode(true);
			style = linkedStyle;
		}

		return style;
	}

	/// <summary>
	/// Gets the styles for the given object
	/// </summary>
	/// <param name="obj">Object to get styles for</param>
	/// <returns>Array of styles</returns>
	static Style[] GetStyles(OpenXmlElement obj)
	{
		IEnumerable<Style>? styles;
		if (obj.GetMainDocumentPart() is MainDocumentPart mainDocumentPart)
			styles = mainDocumentPart.StyleDefinitionsPart?.Styles?.Descendants<Style>();
		else
			styles = obj.Ancestors<Styles>().FirstOrDefault()?.Descendants<Style>();
		if (styles == null) return Array.Empty<Style>();
		return styles as Style[] ?? styles.ToArray();
	}

	/// <summary>
	/// Gets the style from a given array by id
	/// </summary>
	/// <param name="styles">Given style array</param>
	/// <param name="styleId">Identifier to search for</param>
	/// <returns>The style or null otherwise</returns>
	static Style? InnerGetStyleById(IEnumerable<Style> styles, string styleId)
	{
		return styles.FirstOrDefault(c => c.StyleId?.HasValue == true && c.StyleId.Value == styleId);
	}
}
