using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using iTextSharp.text.pdf.draw;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class converts a tab char to a PDF element
/// </summary>
public class TabCharToPdfElement : XmlToPdfElement<Drawing>
{
	/// <summary>
	/// Initializes a new instance of the <see cref="TabCharToPdfElement"/> class.
	/// </summary>
	/// <param name="sourceDocument">the source document</param>
	public TabCharToPdfElement(WordprocessingDocument sourceDocument): base(sourceDocument) { }
	/// <inheritdoc />
	public override IEnumerable<IElement> Process(OpenXmlElement element)
	{
		if (element is not TabChar) return new List<IElement>();

		var defaultTabStop = SourceDocument.MainDocumentPart?.DocumentSettingsPart?.Settings.Descendants<DefaultTabStop>().FirstOrDefault()?.Val?.Value ?? 720.0f;
		return new List<IElement> {new Chunk(new VerticalPositionMark(), Converter.TwipToPoint(defaultTabStop), false)};
	}
}
