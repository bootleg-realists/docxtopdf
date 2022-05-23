using iTextSharp.text;

namespace BootlegRealists.Reporting.Function;

/// <summary>
/// This class contains chunk (PDF) functions.
/// </summary>
public static class ChunkFunction
{
    /// <summary>
    /// Gets a new line chunk
    /// </summary>
    /// <returns>The chunk</returns>
    public static Chunk NewLine { get; } = new("\n");
    /// <summary>
    /// Gets a next page chunk
    /// </summary>
    /// <returns>The chunk</returns>
    public static Chunk NextPage { get; } = new("");
}
