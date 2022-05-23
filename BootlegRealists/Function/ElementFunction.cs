using iTextSharp.text;
using iTextSharp.text.pdf;

namespace BootlegRealists.Reporting.Function;

/// <summary>
/// This class contains element functions.
/// </summary>
public static class ElementFunction
{
	/// <summary>
	/// Get the height of a set of IElements, must provide output width for reference.
	/// This function is called before pdfDoc created, so it must create a PDF document on the fly for calculation.
	/// </summary>
	/// <param name="elements">List of elements to get height for</param>
	/// <param name="width">Width to use</param>
	/// <returns>The height or 0.0f otherwise</returns>
	public static float CalculateHeight(IReadOnlyCollection<IElement> elements, float width)
	{
		var diff = 0f;

		if (elements.Count == 0) return diff;

		using var ms = new MemoryStream();
		var doc = new Document();
		var writer = PdfWriter.GetInstance(doc, ms);
		doc.Open();
		var ct = new ColumnText(writer.DirectContent);
		ct.SetSimpleColumn(0f, 0f, width, 1000f);
		foreach (var t in elements)
			ct.AddElement(t);

		var beforeY = ct.YLine;
		ct.Go();
		diff = beforeY - ct.YLine;
		doc.Close();

		return diff;
	}
}
