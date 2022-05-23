using System.Reflection;

namespace iTextSharp.text.pdf;

/// <summary>
/// This class extends the standard PdfWriter class with additional methods and properties.
/// </summary>
public sealed class PdfWriterEx : PdfWriter
{
	/// <inheritdoc />
	PdfWriterEx(PdfDocument document, Stream os) : base(document, os)
	{
	}

	/// <summary>The  PdfPageEvent for this document.</summary>
	public new IPdfPageEvent? PageEvent { get; set; }

	/// <summary>
	/// The PdfDocument instance
	/// </summary>
	public PdfDocument? PdfDocument { get; private set; }

	/// <summary>
	/// Gets the instance of this class
	/// </summary>
	/// <param name="document">Document to base the writer on</param>
	/// <param name="os">Stream to write to</param>
	/// <returns>The instance</returns>
	public static PdfWriterEx GetInstance(DocumentEx document, Stream os)
	{
		var pdf = CreatePdfDocument();
		document.AddDocListener(pdf);
		var writer = new PdfWriterEx(pdf, os)
		{
			PdfDocument = pdf
		};
		var pageEvent = new PdfPageEventHelperEx(writer);
		((PdfWriter)writer).PageEvent = pageEvent;
		document.PageEvent = pageEvent;
		AddWriter(pdf, writer);
		return writer;
	}

	/// <summary>
	/// Calls the AddWriter method of the given PdfDocument instance
	/// </summary>
	/// <param name="pdf">Instance to call method for</param>
	/// <param name="writer">argument of the method AddWriter</param>
	static void AddWriter(IElementListener pdf, PdfWriterEx writer)
	{
		var m = typeof(PdfDocument).GetMethod("AddWriter", BindingFlags.NonPublic | BindingFlags.Instance);
		m?.Invoke(pdf, new object[] { writer });
	}

	/// <summary>
	/// Creates an instance of PdfDocument
	/// </summary>
	/// <returns>PdfDocument instance</returns>
	static PdfDocument CreatePdfDocument()
	{
		return CreateInstance<PdfDocument>();
	}

	/// <summary>
	/// Creates an instance of a class.
	/// </summary>
	/// <param name="args">Constructor arguments</param>
	/// <typeparam name="T">Type of the class</typeparam>
	/// <returns>The instance</returns>
	static T CreateInstance<T>(params object[] args)
	{
		var type = typeof(T);
		var instance = type.Assembly.CreateInstance(type.FullName ?? string.Empty, false,
			BindingFlags.Instance | BindingFlags.NonPublic, null, args, null, null);
		return (T)instance!;
	}
}
