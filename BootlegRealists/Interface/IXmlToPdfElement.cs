using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text;

namespace BootlegRealists.Reporting.Interface;

/// <summary>
/// This interface represents a node conversion from OpenXML to PDF
/// </summary>
public interface IXmlToPdfElement
{
    /// <summary>
    /// The source document (OpenXML)
    /// </summary>
    WordprocessingDocument SourceDocument {get; }
    /// <summary>
    /// Processes the OpenXML node and returns a list of PDF elements
    /// </summary>
    /// <param name="element">Node to process</param>
    /// <returns>List of processed elements</returns>
    IEnumerable<IElement> Process(OpenXmlElement element);
    /// <summary>
    /// Gets the from type (OpenXML)
    /// </summary>
    /// <returns>The type</returns>
    Type GetSourceType();
}
