using BootlegRealists.Reporting.Extension;

namespace iTextSharp.text.pdf;

/// <summary>
/// This class extends PdfPageEventHelper and implements the advanced functionality (for example: ParagraphEx)
/// </summary>
public class PdfPageEventHelperEx : PdfPageEventHelper
{
	/// <summary>
	/// Paragraph to set properties (such as background color) for
	/// </summary>
	ParagraphEx? paragraph;

	/// <summary>
	/// The start position of the paragraph
	/// </summary>
	float startPosition;

	/// <summary>
	/// Constructor
	/// </summary>
	/// <param name="pdfWriterEx">Instance to use</param>
	public PdfPageEventHelperEx(PdfWriterEx pdfWriterEx)
	{
		PdfWriterEx = pdfWriterEx;
	}

	/// <summary>
	/// The pdf writer property
	/// </summary>
	public PdfWriterEx PdfWriterEx { get; }

	/// <summary>
	/// Sets the element for the extended functionality.
	/// </summary>
	/// <param name="element">Element to set</param>
	public void SetElement(IElement element)
	{
		paragraph = element as ParagraphEx;
	}

	/// <inheritdoc />
	public override void OnChapter(PdfWriter writer, Document document, float paragraphPosition, Paragraph title)
	{
		PdfWriterEx.PageEvent?.OnChapter(writer, document, paragraphPosition, title);
	}

	/// <inheritdoc />
	public override void OnChapterEnd(PdfWriter writer, Document document, float position)
	{
		PdfWriterEx.PageEvent?.OnChapterEnd(writer, document, position);
	}

	/// <inheritdoc />
	public override void OnCloseDocument(PdfWriter writer, Document document)
	{
		PdfWriterEx.PageEvent?.OnCloseDocument(writer, document);
	}

	/// <inheritdoc />
	public override void OnEndPage(PdfWriter writer, Document document)
	{
		PdfWriterEx.PageEvent?.OnEndPage(writer, document);
	}

	/// <inheritdoc />
	public override void OnGenericTag(PdfWriter writer, Document document, Rectangle rect, string text)
	{
		PdfWriterEx.PageEvent?.OnGenericTag(writer, document, rect, text);
	}

	/// <inheritdoc />
	public override void OnOpenDocument(PdfWriter writer, Document document)
	{
		PdfWriterEx.PageEvent?.OnOpenDocument(writer, document);
	}

	/// <inheritdoc />
	public override void OnParagraph(PdfWriter writer, Document document, float paragraphPosition)
	{
		if (paragraph?.BackgroundColor == null)
		{
			PdfWriterEx.PageEvent?.OnParagraph(writer, document, paragraphPosition);
			return;
		}

		startPosition = paragraphPosition;

		PdfWriterEx.PageEvent?.OnParagraph(writer, document, paragraphPosition);
	}

	/// <inheritdoc />
	public override void OnParagraphEnd(PdfWriter writer, Document document, float paragraphPosition)
	{
		if (paragraph?.BackgroundColor == null)
		{
			PdfWriterEx.PageEvent?.OnParagraphEnd(writer, document, paragraphPosition);
			return;
		}

		var cb = writer.DirectContentUnder;
		var indentLeft = paragraph.FirstLineIndent + paragraph.IndentationLeft;
		var descentHeight = 0.0f;
		var calcFont = paragraph.GetCalculatedFont();
		if (calcFont != null && calcFont.BaseFont != null)
		{
			var paragraphFontSize = calcFont.CalculatedSize;
			descentHeight = 0.0f - calcFont.BaseFont.GetFontDescriptor(BaseFont.DESCENT, paragraphFontSize);
		}

		var x = document.Left + indentLeft;
		var y = paragraphPosition + paragraph.SpacingAfter - descentHeight;
		var w = document.Right - document.Left - indentLeft - paragraph.IndentationRight;
		var h = startPosition - paragraphPosition - (paragraph.SpacingBefore + paragraph.SpacingAfter);
		cb.Rectangle(x, y, w, h);
		cb.SetColorFill(paragraph.BackgroundColor);
		cb.Fill();

		PdfWriterEx.PageEvent?.OnParagraphEnd(writer, document, paragraphPosition);
	}

	/// <inheritdoc />
	public override void OnSection(PdfWriter writer, Document document, float paragraphPosition, int depth,
		Paragraph title)
	{
		PdfWriterEx.PageEvent?.OnSection(writer, document, paragraphPosition, depth, title);
	}

	/// <inheritdoc />
	public override void OnSectionEnd(PdfWriter writer, Document document, float position)
	{
		PdfWriterEx.PageEvent?.OnSectionEnd(writer, document, position);
	}

	/// <inheritdoc />
	public override void OnStartPage(PdfWriter writer, Document document)
	{
		PdfWriterEx.PageEvent?.OnStartPage(writer, document);
	}
}
