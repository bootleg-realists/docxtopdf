using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text;
using BootlegRealists.Reporting.Interface;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class contains the first implementation level for an OpenXML to PDF conversion
/// </summary>
/// <typeparam name="TSource">Type of the source (OpenXML)</typeparam>
public abstract class XmlToPdfElement<TSource> : IXmlToPdfElement where TSource : OpenXmlElement
{
	/// <summary>
	/// Initializes a new instance of the <see cref="XmlToPdfElement  &lt; TSource &gt;"/> class.
	/// </summary>
	/// <param name="sourceDocument">the source document</param>
	protected XmlToPdfElement(WordprocessingDocument sourceDocument) => SourceDocument = sourceDocument;
	/// <inheritdoc />
	public WordprocessingDocument SourceDocument { get; }
	/// <inheritdoc />
	public abstract IEnumerable<IElement> Process(OpenXmlElement element);

	/// <inheritdoc />
	public Type GetSourceType() => typeof(TSource);
}
