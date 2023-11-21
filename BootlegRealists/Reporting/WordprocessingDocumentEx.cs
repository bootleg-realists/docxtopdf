
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class extends the class WordprocessingDocument
/// </summary>
public static class WordprocessingDocumentEx
{
    /// <summary>
    /// Creates a new empty WordprocessingDocument.
    /// </summary>
    /// <returns>A new empty WordprocessingDocument.</returns>
    public static WordprocessingDocument CreateEmpty()
    {
        using var stream = new MemoryStream();
        return WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
    }
}