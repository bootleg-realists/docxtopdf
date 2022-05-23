using System.Reflection;
using iTextSharp.text.pdf;

namespace iTextSharp.text;

/// <summary>
/// This class extends the Document class by introducing new methods and properties
/// </summary>
public class DocumentEx : Document
{
	/// <inheritdoc />
	public DocumentEx()
	{
	}

	/// <inheritdoc />
	public DocumentEx(Rectangle pageSize, float marginLeft, float marginRight, float marginTop, float marginBottom)
		: base(pageSize, marginLeft, marginRight, marginTop, marginBottom)
	{
	}

	/// <summary>
	/// Page event
	/// </summary>
	public PdfPageEventHelperEx? PageEvent { get; set; }

	/// <summary>
	/// Gets the current height
	/// </summary>
	public float CurrentHeight => PageEvent?.PdfWriterEx.PdfDocument?.CurrentHeight ?? 0.0f;

	/// <summary>
	/// Gets the indent from the top
	/// </summary>
	public float IndentTop
	{
		get
		{
			var pdfDocument = PageEvent?.PdfWriterEx.PdfDocument;
			if (pdfDocument == null)
				return 0.0f;

			var indentTop = GetProperty(pdfDocument, nameof(IndentTop));
			if (indentTop != null) return (float)indentTop;
			return 0.0f;
		}
	}

	/// <summary>
	/// Gets the indent from the bottom
	/// </summary>
	public float IndentBottom
	{
		get
		{
			var pdfDocument = PageEvent?.PdfWriterEx.PdfDocument;
			if (pdfDocument == null)
				return 0.0f;

			var indentBottom = GetProperty(pdfDocument, nameof(IndentBottom));
			if (indentBottom != null) return (float)indentBottom;
			return 0.0f;
		}
	}

	/// <inheritdoc />
	public override bool Add(IElement element)
	{
		PageEvent?.SetElement(element);
		return base.Add(element);
	}

	static object? GetProperty(object obj, string propertyName)
	{
		var prop = Array.Find(obj.GetType().GetProperties(BindingFlags.Instance | BindingFlags.NonPublic),
			p => string.Equals(p.Name, propertyName, StringComparison.Ordinal));
		return prop?.GetValue(obj);
	}
}
