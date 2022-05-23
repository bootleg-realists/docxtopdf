using System.Diagnostics.CodeAnalysis;
using iTextSharp.text;
using iTextSharp.text.pdf.draw;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains (PDF) chunk extension methods.
/// </summary>
public static class PdfChunkExtension
{
	/// <summary>
	/// Gets the tab settings
	/// </summary>
	/// <param name="chunk">Chunk to get settings for</param>
	/// <param name="separator">[out] The drawInterface to use to draw the tab.</param>
	/// <param name="tabPosition">[out] The tab position (X coordinate)</param>
	/// <param name="newLine">[out] If true, a newline will be added if the tabPosition has already been reached.</param>
	/// <param name="adjustLeft">The extra adjustment to use (left)</param>
	/// <returns>True if tab settings are fetched and false otherwise</returns>
	public static bool GetTabSettings(this Chunk chunk, [MaybeNullWhen(false)] out IDrawInterface separator, out float tabPosition, out bool newLine, out int adjustLeft)
	{
		var attributes = chunk.Attributes;
		if (attributes?.ContainsKey(Chunk.TAB) != true || attributes[Chunk.TAB] is not object[] tab || tab.Length != 4)
		{
			separator = null;
			tabPosition = 0.0f;
			newLine = false;
			adjustLeft = 0;
			return false;
		}

		if (tab[0] is not IDrawInterface || tab[1] is not float || tab[2] is not bool || tab[3] is not int)
		{
			separator = null;
			tabPosition = 0.0f;
			newLine = false;
			adjustLeft = 0;
			return false;
		}
		separator = (IDrawInterface)tab[0];
		tabPosition = (float)tab[1];
		newLine = (bool)tab[2];
		adjustLeft = (int)tab[3];
		return true;
	}

	/// <summary>
	/// Sets the tab settings
	/// </summary>
	/// <param name="chunk">Chunk to set settings for</param>
	/// <param name="separator">The drawInterface to use to draw the tab.</param>
	/// <param name="tabPosition">The tab position (X coordinate)</param>
	/// <param name="newLine">If true, a newline will be added if the tabPosition has already been reached.</param>
	/// <param name="adjustLeft">The extra adjustment to use (left)</param>
	public static void SetTabSettings(this Chunk chunk, IDrawInterface separator, float tabPosition, bool newLine, int adjustLeft)
	{
		var attributes = chunk.Attributes;
		attributes[Chunk.TAB] = new object[] { separator, tabPosition, newLine, adjustLeft };
	}
}
