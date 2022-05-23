using iTextSharp.text;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains base color extension methods.
/// </summary>
public static class BaseColorExtension
{
    /// <summary>
    /// Checks if the given color is transparent
    /// </summary>
    /// <param name="obj">Object to act on</param>
    /// <returns>True if it is and false otherwise</returns>
    public static bool IsTransparent(this BaseColor obj)
    {
        var color = obj.ToArgb();
        var c = System.Drawing.Color.FromArgb(color);
        return c.A == 0;
    }
}