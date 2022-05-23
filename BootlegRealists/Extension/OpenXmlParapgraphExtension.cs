using DocumentFormat.OpenXml.Wordprocessing;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains paragraph (OpenXML) extension methods.
/// </summary>
public static class OpenXmlParagraphExtension
{
	/// <summary>
	/// Checks if the given object uses contextual spacing
	/// </summary>
	/// <param name="obj">The object to act on</param>
	/// <returns>True if it does and false otherwise</returns>
	public static bool UsesContextualSpacing(this Paragraph obj)
	{
		var result = true;
		var cs = obj.GetEffectiveElement<ContextualSpacing>();
		if (cs != null)
			result = Converter.OnOffToBool(cs);

		return result;
	}

	/// <summary>
	/// Check if the other paragraph's style is same as this paragraph's style
	/// </summary>
	/// <param name="obj">The object to act on</param>
	/// <param name="other">Other paragraph</param>
	/// <returns>True if it is and false otherwise</returns>
	public static bool HasSameStyle(this Paragraph obj, Paragraph other)
	{
		return other.ParagraphProperties?.ParagraphStyleId != null &&
		       obj.ParagraphProperties?.ParagraphStyleId != null &&
		       string.Equals(other.ParagraphProperties.ParagraphStyleId.Val,
			       obj.ParagraphProperties.ParagraphStyleId.Val, StringComparison.Ordinal);
	}
}
