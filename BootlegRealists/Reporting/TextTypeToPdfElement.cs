using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class converts a text type to a PDF element
/// </summary>
public class TextTypeToPdfElement : XmlToPdfElement<Drawing>
{
	/// <summary>
	/// Initializes a new instance of the <see cref="TextTypeToPdfElement"/> class.
	/// </summary>
	/// <param name="sourceDocument">the source document</param>
	public TextTypeToPdfElement(WordprocessingDocument sourceDocument): base(sourceDocument) { }
	/// <inheritdoc />
	public override IEnumerable<IElement> Process(OpenXmlElement element)
	{
		if (element is not TextType text) return new List<IElement>();
		string str;
		if (text.Space != null)
			str = text.Space?.Value == SpaceProcessingModeValues.Preserve ? text.InnerText : text.InnerText.Trim();
		else
			str = text.InnerText.Trim();

		return !string.IsNullOrEmpty(str) ? new List<IElement> {new Chunk(str)} : new List<IElement>();
	}
}
