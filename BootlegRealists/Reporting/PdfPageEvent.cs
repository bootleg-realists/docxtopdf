using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Extension;
using BootlegRealists.Reporting.Function;
using Pdf = iTextSharp.text.pdf;
using Text = iTextSharp.text;

namespace BootlegRealists.Reporting;

public partial class DocxToPdf
{
	/// <summary>
	/// For drawing header and footer, and background color
	/// </summary>
	class PdfPageEvent : Pdf.PdfPageEventHelper
	{
		/// <summary>
		/// The converter
		/// </summary>
		readonly DocxToPdf converter;

		public PdfPageEvent(DocxToPdf obj)
		{
			converter = obj;
		}

		/// <inheritdoc />
		public override void OnEndPage(Pdf.PdfWriter writer, Text.Document doc)
		{
			var body = converter.docxDocument.MainDocumentPart?.Document.Body;
			var section = body?.Descendants<SectionProperties>().FirstOrDefault();
			if (section == null) return;
			var margin = section.GetFirstElement<PageMargin>();
			if (margin == null) return;
			var pageWidth = doc.PageSize.Width - doc.LeftMargin - doc.RightMargin;

			// Draw header
			var contents = converter.BuildHeader(section);
			if (contents.Count > 0)
			{
				var margin2 = margin.Header != (object?)null ? Converter.TwipToPoint(margin.Header.Value) : 0f;
				var columnText = new Pdf.ColumnText(writer.DirectContent);
				// in iTextSharp page coordinate concept, for Y-axis, the top edge of
				// the page has the maximum value (doc.PageSize.Height) and the 
				// bottom edge of the page is zero
				columnText.SetSimpleColumn(doc.LeftMargin, doc.PageSize.Height - doc.TopMargin,
					doc.PageSize.Width - doc.RightMargin, doc.PageSize.Height - margin2);
				foreach (var t in contents) columnText.AddElement(t);

				columnText.Go();
			}

			// Draw footer
			contents = converter.BuildFooter(section);
			if (contents.Count == 0)
				return;

			{
				var margin2 = margin.Footer != (object?)null ? Converter.TwipToPoint(margin.Footer.Value) : 0f;
				var ct = new Pdf.ColumnText(writer.DirectContent);

				var height = ElementFunction.CalculateHeight(contents, pageWidth);

				// in iTextSharp page coordinate concept, for Y-axis, the top edge of
				// the page has the maximum value (doc.PageSize.Height) and the 
				// bottom edge of the page is zero
				ct.SetSimpleColumn(doc.LeftMargin, margin2, doc.PageSize.Width - doc.RightMargin, margin2 + height);
				foreach (var t in contents)
					ct.AddElement(t);

				ct.Go();
			}
		}
	}
}
