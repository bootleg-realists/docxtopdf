using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using BootlegRealists.Reporting.Extension;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class converts a picture to a PDF element
/// </summary>
public class PictureToPdfElement : XmlToPdfElement<Picture>
{
	/// <summary>
	/// Initializes a new instance of the <see cref="PictureToPdfElement"/> class.
	/// </summary>
	/// <param name="sourceDocument">the source document</param>
	public PictureToPdfElement(WordprocessingDocument sourceDocument): base(sourceDocument) { }
	/// <inheritdoc />
	public override IEnumerable<IElement> Process(OpenXmlElement element)
	{
		Image? ret = null;

		var shapes = element.Descendants<Shape>();
		foreach (var shape in shapes)
		{
			var img = shape.Descendants<ImageData>().FirstOrDefault();
			if (img == null)
				continue;

			var bImg = SourceDocument.MainDocumentPart?.GetImageById(img.RelationshipId?.Value ?? "");
			if (bImg == null)
				continue;

			ret = Image.GetInstance(bImg, ImageFormat.Png);
			break;
		}

		return ret != null ? new List<IElement> {ret} : new List<IElement>();		
	}
}
