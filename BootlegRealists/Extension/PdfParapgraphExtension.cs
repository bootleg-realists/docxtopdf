using iTextSharp.text;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains paragraph (PDF) extension methods.
/// </summary>
public static class PdfParagraphExtension
{
	/// <summary>
	/// Gets the font for the given paragraph. It uses the hightest font size in its children.
	/// </summary>
	/// <param name="paragraph">Given paragraph</param>
	/// <returns>The font or null otherwise</returns>
	public static Font? GetCalculatedFont(this Paragraph paragraph)
	{
		var fontSize = 0.0f;
		Font? result = null;

		foreach (var element in paragraph)
		{
			if (element is not Chunk chunk || element is SpaceChunk || chunk.Font.CalculatedSize <= fontSize) continue;
			fontSize = chunk.Font.CalculatedSize;
			result = chunk.Font;
		}

		return result;
	}

	/// <summary>
	/// Processes breaks in a given paragraph.
	/// </summary>
	/// <param name="paragraph">Given paragraph</param>
	/// <returns>The paragraph broken into multiple paragraph based on the breaks</returns>
	public static IEnumerable<Paragraph> ProcessBreaks(this Paragraph paragraph)
	{
		var result = new List<Paragraph>();
		var begin = 0;
		var i = 0;
		while (i < paragraph.Chunks.Count)
		{
			if (paragraph.Chunks[i] is not Chunk chunk || !string.Equals(chunk.Content, "\n", StringComparison.Ordinal))
			{
				i++;
				continue;
			}

			// cut from begin to i
			var item1 = paragraph.CloneProperties();
			for (var j = begin; j < i; j++)
				item1.Add(paragraph.Chunks[j]);

			var item2 = paragraph.CloneProperties();
			// Breaks don't have spacing
			item2.SpacingBefore = 0.0f;
			item2.SpacingAfter = 0.0f;
			item2.Add(chunk);

			if (item1.Count > 0)
			{
				item1.SpacingAfter = 0.0f;
				result.Add(item1);
			}
			result.Add(item2);
			i++;
			begin = i;
		}
		if (result.Count == 0) return new List<Paragraph> { paragraph };
		if (begin >= paragraph.Chunks.Count) return result;

		var item3 = paragraph.CloneProperties();
		for (var j = begin; j < paragraph.Chunks.Count; j++)
			item3.Add(paragraph.Chunks[j]);
		if (item3.Count > 0)
			result.Add(item3);
		return result;
	}
	/// <summary>
	/// Processes tabs in a given paragraph.
	/// </summary>
	/// <param name="paragraph">Given paragraph</param>
	public static void ProcessTabs(this Paragraph paragraph)
	{
		var offset = 0.0f;
		foreach (var item in paragraph.Chunks)
		{
			if (item is not Chunk chunk || !chunk.GetTabSettings(out var separator, out var tabPosition, out var newLine, out var adjustLeft))
				continue;
			tabPosition += offset;
			chunk.SetTabSettings(separator, tabPosition, newLine, adjustLeft);
			offset = tabPosition;
		}
	}

	/// <summary>
	/// Clones the given paragraph (only the properties)
	/// </summary>
	/// <param name="paragraph">Given paragraph</param>
	/// <returns>The clone</returns>
	static Paragraph CloneProperties(this Paragraph paragraph)
	{
		var result = new Paragraph
		{
			Hyphenation = paragraph.Hyphenation,
			Font = paragraph.Font,
			SpacingBefore = paragraph.SpacingBefore,
			SpacingAfter = paragraph.SpacingAfter,
			KeepTogether = paragraph.KeepTogether,
			IndentationRight = paragraph.IndentationRight,
			IndentationLeft = paragraph.IndentationLeft,
			Alignment = paragraph.Alignment,
			ExtraParagraphSpace = paragraph.ExtraParagraphSpace,
			FirstLineIndent = paragraph.FirstLineIndent
		};
		result.SetLeading(paragraph.Leading, paragraph.MultipliedLeading);
		return result;
	}
}
