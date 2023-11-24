using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using BootlegRealists.Reporting.Extension;
using SkiaSharp;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class converts a drawing to a PDF element
/// </summary>
public class DrawingToPdfElement : XmlToPdfElement<Drawing>
{
	/// <summary>
	/// Initializes a new instance of the <see cref="DrawingToPdfElement"/> class.
	/// </summary>
	/// <param name="sourceDocument">the source document</param>
	public DrawingToPdfElement(WordprocessingDocument sourceDocument): base(sourceDocument) { }
	/// <inheritdoc />
	public override IEnumerable<IElement> Process(OpenXmlElement element)
	{
		var blipElement = element.Descendants<Blip>().First();
		var imageId = blipElement.Embed?.Value ?? "";

		var bImg = SourceDocument.MainDocumentPart?.GetImageById(imageId);
		if (bImg == null) return new List<IElement>();

		using var skBitmap = SKBitmap.Decode(bImg);
		if (skBitmap == null) return new List<IElement>();
		var ret = Image.GetInstance(skBitmap, SKEncodedImageFormat.Png);
		var extend = element.Descendants<Extent>().FirstOrDefault();
		if (extend == null) return new List<IElement>();
		const float inchIsEmu = 914400.0f;
		var newWidth = (extend.Cx?.Value ?? 0.0f) / inchIsEmu * 72.0f;
		var newHeight = (extend.Cy?.Value ?? 0.0f) / inchIsEmu * 72.0f;

		ret.ScaleAbsolute(newWidth, newHeight);

		return new List<IElement> {ret};
	}
}
