using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Extension;
using Pdf = iTextSharp.text.pdf;
using Text = iTextSharp.text;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class builds the pdf table.
/// </summary>
public static class TableBuilder
{
	/// <summary>
	/// Builds the cell height
	/// </summary>
	/// <param name="cell">Cell to build height for</param>
	/// <param name="elements">List of elements in the cell</param>
	/// <param name="minRowHeight">The minimum row height in points</param>
	/// <param name="exactRowHeight">The exact row height in points</param>
	public static void BuildCellHeight(Pdf.PdfPCell cell, IEnumerable<Text.IElement> elements, float minRowHeight, float exactRowHeight)
	{
		// add elements to cell
		var adjustPaddingDone = false; // magic: simulate Word cell padding
		//cell.AddElement(new Text.Paragraph(0f, "\u00A0")); // magic: to allow Paragraph.SpacingBefore take effect
		foreach (var element in elements)
		{
			if (element is Text.Paragraph pg)
			{
				var totalLeading = GetLeading(pg);

				pg.SetLeading(totalLeading * 0.9f, 0f); // magic: paragraph leading in table becomes 0.9

				// ------
				// magic: simulate Word cell padding
				if (!adjustPaddingDone)
				{
					cell.PaddingTop -= pg.TotalLeading * 0.25f;
					cell.PaddingBottom += pg.TotalLeading * 0.25f;
					var contentHeight = pg.TotalLeading + cell.PaddingTop + cell.PaddingBottom;
					if (float.IsNaN(minRowHeight) && float.IsNaN(exactRowHeight))
						cell.MinimumHeight = contentHeight;
					else if (!float.IsNaN(exactRowHeight))
						cell.FixedHeight = contentHeight > exactRowHeight ? contentHeight : exactRowHeight;
					else if (!float.IsNaN(minRowHeight))
						cell.MinimumHeight = contentHeight > minRowHeight ? contentHeight : minRowHeight;
					else
						cell.FixedHeight = contentHeight;
					adjustPaddingDone = true;
				}
			}

			cell.AddElement(element);
		}
	}

	static float GetLeading(Text.Paragraph pg)
	{
		const float wordDefaultAtLeastLineSpacing = 1.15f;
		var result = pg.GetCalculatedFont();
		if (result == null) return 16.0f;

		var totalLeading = result.CalculatedSize * wordDefaultAtLeastLineSpacing;
		return totalLeading;
	}

	/// <summary>
	/// Builds the table indentation and returns the table (maybe wrapped in a paragraph)
	/// </summary>
	/// <param name="wordTable">The Word table</param>
	/// <param name="pdfTable">The pdf table</param>
	/// <returns>The table or null otherwise</returns>
	public static Text.IElement BuildIndentation(Table wordTable, Pdf.PdfPTable pdfTable)
	{
		const float spacingBeforeMagic = 5.0f;
		// Table indentation
		var ind = wordTable.GetEffectiveElement<TableIndentation>();
		if (ind?.Type == null || ind.Type.Value != TableWidthUnitValues.Dxa || ind.Width?.HasValue != true)
		{
			pdfTable.SpacingBefore = spacingBeforeMagic;
			return pdfTable;
		}

		// Use the trick of wrap table into paragraph to achieve table indentation
		var width = Converter.TwipToPoint(ind.Width.Value);
		if (width <= 0.0f)
		{
			pdfTable.SpacingBefore = spacingBeforeMagic;
			return pdfTable;
		}

		var pg = new Text.Paragraph(0, "\u00a0") { IndentationLeft = width };
		pg.Add(pdfTable);

		pg.SpacingBefore = spacingBeforeMagic;
		return pg;
	}

	/// <summary>
	/// Gets the row padding for the given row identifier
	/// </summary>
	/// <param name="tableHelper">Table helper</param>
	/// <param name="rowId">Row identifier</param>
	/// <param name="topPadding">[out] The top padding in points</param>
	/// <param name="bottomPadding">[out] The bottom padding in points</param>
	public static void GetRowHeightPadding(TableHelper tableHelper, int rowId, out float topPadding,
		out float bottomPadding)
	{
		topPadding = 0.0f;
		bottomPadding = 0.0f;
		if (rowId < 0) return;
		foreach (var cellHelper in tableHelper.Cast<TableHelperCell>().Where(x => x.RowId == rowId))
		{
			var top = cellHelper.Cell?.GetEffectiveElement<TopMargin>();
			var currentTop = top?.Width != null ? Converter.TwipToPoint(top.Width.Value ?? "") : 0.0f;
			var bottom = cellHelper.Cell?.GetEffectiveElement<BottomMargin>();
			var currentBottom = bottom?.Width != null ? Converter.TwipToPoint(bottom.Width.Value ?? "") : 0.0f;

			BorderType? br = cellHelper.Borders?.TopBorder;
			if (br?.Val != null && br.Val.Value != BorderValues.Nil && br.Val.Value != BorderValues.None)
				currentTop += br.Size != (object?)null ? Converter.OneEighthPointToPoint(br.Size.Value) : 0.0f;
			br = cellHelper.Borders?.BottomBorder;
			if (br?.Val != null && br.Val.Value != BorderValues.Nil && br.Val.Value != BorderValues.None)
				currentBottom += br.Size != (object?)null ? Converter.OneEighthPointToPoint(br.Size.Value) : 0.0f;
			if (currentTop > topPadding) topPadding = currentTop;
			if (currentBottom > bottomPadding) bottomPadding = currentBottom;
		}
	}

	/// <summary>
	/// Gets the margin for the given cell
	/// </summary>
	/// <param name="cell">Cell to get margin for</param>
	/// <typeparam name="TMargin">Type of the margin</typeparam>
	/// <typeparam name="TTableCellMargin">Type of the table cell margin</typeparam>
	/// <returns>the margin or float.NaN otherwise</returns>
	public static float GetMargin<TMargin, TTableCellMargin>(TableCell cell) where TMargin : TableWidthType
		where TTableCellMargin : TableWidthDxaNilType
	{
		var margin = cell.GetEffectiveElement<TMargin>();
		if (margin?.Width != null)
			return Converter.TwipToPoint(margin.Width.Value ?? "");

		var tableCellMargin = cell.GetEffectiveElement<TTableCellMargin>();
		return tableCellMargin?.Width != (object?)null
			? Converter.TwipToPoint(tableCellMargin.Width.Value)
			: float.NaN;
	}

	/// <summary>
	/// Builds the table cell border
	/// </summary>
	/// <param name="cellHelper">Cell's helper reference</param>
	/// <param name="cell">Destination pdf cell</param>
	/// <param name="topPadding">Top padding to use</param>
	/// <param name="bottomPadding">Bottom padding to use</param>
	public static void BuildTableCellBorder(TableHelperCell cellHelper, Pdf.PdfPCell cell, float topPadding,
		float bottomPadding)
	{
		// Shading
		var sh = cellHelper.Cell?.GetEffectiveElement<Shading>();
		if (sh?.Fill?.HasValue == true && sh.Fill.Value != "auto")
			cell.BackgroundColor = new Text.BaseColor(Convert.ToInt32(sh.Fill.Value, 16));

		// Border
		//  top border
		BorderType? br = cellHelper.Borders?.TopBorder;
		if (br?.Val == null || br.Val.Value is BorderValues.Nil or BorderValues.None)
		{
			cell.Border &= ~Text.Rectangle.TOP_BORDER;
		}
		else
		{
			cell.BorderColorTop = br.Color != null && br.Color.Value != "auto"
				? new Text.BaseColor(Convert.ToInt32(br.Color.Value, 16))
				: null;
			cell.BorderWidthTop = br.Size != (object?)null ? Converter.OneEighthPointToPoint(br.Size.Value) : 0.0f;
		}

		//  bottom border
		br = cellHelper.Borders?.BottomBorder;
		if (br?.Val == null || br.Val.Value is BorderValues.Nil or BorderValues.None)
		{
			cell.Border &= ~Text.Rectangle.BOTTOM_BORDER;
		}
		else
		{
			cell.BorderColorBottom = br.Color != null && br.Color.Value != "auto"
				? new Text.BaseColor(Convert.ToInt32(br.Color.Value, 16))
				: null;
			cell.BorderWidthBottom =
				br.Size != (object?)null ? Converter.OneEighthPointToPoint(br.Size.Value) : 0.0f;
		}

		//  left border
		br = cellHelper.Borders?.LeftBorder;
		if (br?.Val == null || br.Val.Value is BorderValues.Nil or BorderValues.None)
		{
			cell.Border &= ~Text.Rectangle.LEFT_BORDER;
		}
		else
		{
			cell.BorderColorLeft = br.Color != null && br.Color.Value != "auto"
				? new Text.BaseColor(Convert.ToInt32(br.Color.Value, 16))
				: null;
			cell.BorderWidthLeft =
				br.Size != (object?)null ? Converter.OneEighthPointToPoint(br.Size.Value) : 0.0f;
		}

		//  right border
		br = cellHelper.Borders?.RightBorder;
		if (br?.Val == null || br.Val.Value is BorderValues.Nil or BorderValues.None)
		{
			cell.Border &= ~Text.Rectangle.RIGHT_BORDER;
		}
		else
		{
			cell.BorderColorRight = br.Color != null && br.Color.Value != "auto"
				? new Text.BaseColor(Convert.ToInt32(br.Color.Value, 16))
				: null;
			cell.BorderWidthRight =
				br.Size != (object?)null ? Converter.OneEighthPointToPoint(br.Size.Value) : 0.0f;
		}

		cell.PaddingTop = topPadding;
		cell.PaddingBottom = bottomPadding;
	}
}
