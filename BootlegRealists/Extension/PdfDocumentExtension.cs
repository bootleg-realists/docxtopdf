using iTextSharp.text;

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains (PDF) document extension methods.
/// </summary>
public static class PdfDocumentExtension
{
	/// <summary>
	/// Post processes and adds the elements to the given document.
	/// </summary>
	/// <param name="document">Given document</param>
	/// <param name="elements">Elements to add</param>
	/// <returns>The paragraph broken into multiple paragraph based on the breaks</returns>
	public static IEnumerable<IElement> Process(this Document document, IList<IElement> elements)
	{
		var newList = new List<IElement>();
		for (var i = 0; i < elements.Count; i++)
		{
			if (i == 0 && elements[i] is Paragraph)
			{
				// to allow spacing before
				newList.Add(new Paragraph(0, "\u00a0"));
			}

			if (elements[i] is Paragraph paragraph)
			{
				foreach (var p in paragraph.ProcessBreaks())
				{
					p.ProcessTabs();
					newList.Add(p);
				}
				continue;
			}
			newList.Add(elements[i]);
		}
		for (var i = 1; i < newList.Count; i++)
			UseSpacingAfter(newList[i - 1], newList[i]);
		return newList;
	}

	/// <summary>
	/// Use spacing after exclusively instead of spacing before
	/// </summary>
	/// <param name="previous">Previous element</param>
	/// <param name="current">Current element</param>
	static void UseSpacingAfter(IElement previous, IElement current)
	{
		if (previous is not Paragraph pp || current is not Paragraph cp)
			return;
			
		if (cp.SpacingBefore <= 0.0f)
			return;

		pp.SpacingAfter += cp.SpacingBefore;
		cp.SpacingBefore = 0.0f;
	}
}
