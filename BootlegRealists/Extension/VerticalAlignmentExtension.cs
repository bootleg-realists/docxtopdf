using BootlegRealists.Reporting.Enumeration;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class provides extension methods for the vertical alignment enumeration.
/// </summary>
public static class VerticalAlignmentExtension
{
    /// <summary>
    /// Checks if the given object has an offset.
    /// </summary>
    /// <param name="obj">Object to check</param>
    /// <returns>True if it has and false otherwise</returns>
    public static bool HasOffset(this VerticalAlignment obj)
    {
        return obj == VerticalAlignment.Superscript || obj == VerticalAlignment.Subscript;
    }
}