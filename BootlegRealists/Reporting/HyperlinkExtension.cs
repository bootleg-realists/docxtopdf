using DocumentFormat.OpenXml.Wordprocessing;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains hyper link extension methods.
/// </summary>
public static class HyperlinkExtension
{
	/// <summary>
	/// Gets the URL for the given hyperlink
	/// </summary>
	/// <param name="obj">Object to act on.</param>
	/// <returns>The URL or null otherwise</returns>
	public static string GetUrl(this Hyperlink obj)
	{
		if (obj.Id == null) return "";

		var mainDocumentPart = obj.GetMainDocumentPart();
		if (mainDocumentPart == null) return "";
		var hyperlinkRelationships = mainDocumentPart.HyperlinkRelationships;
		var id = obj.Id.Value;
		var hr = hyperlinkRelationships.FirstOrDefault(c => c.Id == id);
		return hr?.IsExternal == true ? hr.Uri.OriginalString : "";
	}
}
