using System.Collections;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;
using BootlegRealists.Reporting.Extension;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class contains table helper functions.
/// </summary>
public class TableHelper : IEnumerable
{
	static readonly Dictionary<BorderValues, int> BorderNumber = new()
	{
		{ BorderValues.Single, 1 },
		{ BorderValues.Thick, 2 },
		{ BorderValues.Double, 3 },
		{ BorderValues.Dotted, 4 },
		{ BorderValues.Dashed, 5 },
		{ BorderValues.DotDash, 6 },
		{ BorderValues.DotDotDash, 7 },
		{ BorderValues.Triple, 8 },
		{ BorderValues.ThinThickSmallGap, 9 },
		{ BorderValues.ThickThinSmallGap, 10 },
		{ BorderValues.ThinThickThinSmallGap, 11 },
		{ BorderValues.ThinThickMediumGap, 12 },
		{ BorderValues.ThickThinMediumGap, 13 },
		{ BorderValues.ThinThickThinMediumGap, 14 },
		{ BorderValues.ThinThickLargeGap, 15 },
		{ BorderValues.ThickThinLargeGap, 16 },
		{ BorderValues.ThinThickThinLargeGap, 17 },
		{ BorderValues.Wave, 18 },
		{ BorderValues.DoubleWave, 19 },
		{ BorderValues.DashSmallGap, 20 },
		{ BorderValues.DashDotStroked, 21 },
		{ BorderValues.ThreeDEmboss, 22 },
		{ BorderValues.ThreeDEngrave, 23 },
		{ BorderValues.Outset, 24 },
		{ BorderValues.Inset, 25 }
	};

	readonly List<TableHelperCell> cells = new();

	// ------
	// Below variables are assigned in ParseTable()
	TableBorders? condFmtbr;
	Table? table;

	float[]? tableColumnsWidth; // after adjustTableColumnsWidth()
	// ------

	/// <summary>
	/// Get a float array which indicates all the table column width in points. This array is only available after
	/// ParseTable() got called.
	/// </summary>
	public float[]? TableColumnsWidth => tableColumnsWidth;

	/// <summary>
	/// /// Get table column length.
	/// </summary>
	public int ColumnLength { get; private set; }

	/// <summary>
	/// Get table row length.
	/// </summary>
	int RowLength { get; set; }

	/// <summary>
	/// Return useful cells (i.e. the cell has TableCell and Text.IElement elements).
	/// </summary>
	/// <returns></returns>
	public IEnumerator GetEnumerator()
	{
		if (cells.Count == 0)
			yield break;

		var maxId = cells[^1].CellId + 1;
		var id = 0;
		while (id < maxId)
			foreach (var cell in cells)
			{
				if (cell.CellId != id)
					continue;

				id++;
				yield return cell;
			}
	}

	/// <summary>
	/// Parses the given table
	/// </summary>
	/// <param name="t">Table to parse</param>
	public void ParseTable(Table t)
	{
		table = t;
		tableColumnsWidth = GetTableGridCols(table);
		ColumnLength = tableColumnsWidth?.Length ?? 0;

		var cellId = 0;
		var rowId = 0;
		foreach (var row in table.Elements<TableRow>())
		{
			var rowStart = true;

			// w:gridBefore
			var tmpGridBefore = row.GetEffectiveElement<GridBefore>();
			var skipGridsBefore = tmpGridBefore?.Val != (object?)null ? tmpGridBefore.Val.Value : 0;

			// w:gridAfter
			var tmpGridAfter = row.GetEffectiveElement<GridAfter>();
			var skipGridsAfter = tmpGridAfter?.Val != (object?)null ? tmpGridAfter.Val.Value : 0;

			var colId = 0;

			// gridBefore (with same cellId and set as blank)
			if (skipGridsBefore > 0)
			{
				for (var i = 0; i < skipGridsBefore; i++)
				{
					cells.Add(new TableHelperCell(this, cellId, rowId, colId));
					colId++; // increase for each cells.Add()
				}

				cellId++;
			}

			foreach (var col in row.Elements<TableCell>())
			{
				var currentCellId = cellId;
				var cellCount = 1;

				var basecell = new TableHelperCell(this, currentCellId, rowId, colId);
				if (rowStart)
				{
					basecell.RowStart = true;
					rowStart = false;
				}

				basecell.Row = row;
				basecell.Cell = col;

				// process rowspan and colspan
				if (col.TableCellProperties != null)
				{
					// colspan
					if (col.TableCellProperties.GridSpan != null)
						cellCount = col.TableCellProperties.GridSpan.Val ?? 0;

					// rowspan
					if (col.TableCellProperties.VerticalMerge != null)
					{
						// "continue": get cellId from (rowId-1):(colId)
						if (col.TableCellProperties.VerticalMerge.Val == null || col.TableCellProperties.VerticalMerge.Val == "continue")
						{
							currentCellId = cells[ColumnLength * (rowId - 1) + colId].CellId;
						}
						// "restart": the begin of rowspan
						else
						{
							currentCellId = cellId;
						}
					}
				}

				basecell.CellId = currentCellId;

				cells.Add(basecell);
				colId++; // increase for each cells.Add()

				for (var i = 1; i < cellCount; i++)
				{
					// Add spanned cells
					cells.Add(new TableHelperCell(this, currentCellId, rowId, colId));
					colId++; // increase for each cells.Add()
				}

				// The latest cellId was used, then we must increase it for future usage
				if (cellId == currentCellId) cellId++;
			}

			var rowEndIndex = cells.Count - 1;
			var rowEndCellId = cells[rowEndIndex].CellId;
			while (rowEndIndex > 0 && cells[rowEndIndex - 1].CellId == rowEndCellId) rowEndIndex--;

			cells[rowEndIndex].RowEnd = true;

			// gridAfter (with same cellId and set as blank)
			if (skipGridsAfter > 0 && colId < ColumnLength)
			{
				for (var i = 0; i < skipGridsAfter; i++)
				{
					cells.Add(new TableHelperCell(this, cellId, rowId, colId));
					colId++; // increase for each cells.Add()
				}

				cellId++;
			}

			rowId++;
		}

		RowLength = rowId;

		// ====== Adjust table columns width by their content ======

		AdjustTableColumnsWidth();

		// ====== Resolve cell border conflict ======

		// prepare table conditional formatting (border), which will be used in
		// applyCellBorders() so must be called before applyCellBorders()
		RollingUpTableBorders();
		for (var r = 0; r < RowLength; r++)
			// The following handles the situation where
			// if table innerV is set, and cnd format for first row specifies nil border, then nil border wins.
			// if table innerH is set, and cnd format for first columns specifies nil border, then table innerH wins.

			//// TODO: if row's cellspacing is not zero then bypass this row
			//Wordprocessing.TableCellSpacing tcspacing = _CvrtCell.GetTableRow(cells, r).TableRowProperties.Descendants<Wordprocessing.TableCellSpacing>().FirstOrDefault();
			//if (tcspacing.Type.Value != Wordprocessing.TableWidthUnitValues.Nil)
			//    continue;

		for (var c = 0; c < ColumnLength; c++)
		{
			var me = GetCell(r, c);
			if (me?.Blank != false)
				continue;

			me.Borders ??= ApplyCellBorders(me.Cell?.Descendants<TableCellBorders>().FirstOrDefault(),
				me.ColId == 0 || me.RowStart,
				me.ColId + GetColSpan(me.CellId) == ColumnLength || me.RowEnd,
				me.RowId == 0,
				me.RowId + GetRowSpan(me.CellId) == RowLength
			);

			var colspan = GetColSpan(me.CellId);
			var rowspan = GetRowSpan(me.CellId);

			// Process the cells at the right side of me
			//   Can bypass column-spanned cells because they never exist
			if (c + (colspan - 1) + 1 < ColumnLength) // not last column
			{
				var rights = new List<TableHelperCell>();
				for (var i = 0; i < rowspan; i++)
				{
					var tmp = GetCell(r + i, c + (colspan - 1) + 1);
					if (tmp?.Blank == false) rights.Add(tmp);
				}

				if (rights.Count > 0)
				{
					foreach (var right in rights)
					{
						right.Borders ??= ApplyCellBorders(
							right.Cell?.Descendants<TableCellBorders>().FirstOrDefault(),
							right.ColId == 0 || right.RowStart,
							right.ColId + GetColSpan(right.CellId) == ColumnLength || right.RowEnd,
							right.RowId == 0,
							right.RowId + GetRowSpan(right.CellId) == RowLength
						);

						var meWin = CompareBorder(me.Borders, right.Borders, CompareDirection.Horizontal);
						if (meWin) me.Borders.RightBorder?.CopyAttributesTo(right.Borders.LeftBorder);
					}

					me.Borders.RightBorder?.ClearAllAttributes();
				}
			}

			// Process the cells below me
			//   Can't bypass row-spanned cells because they still have tcBorders property
			if (r + 1 >= RowLength)
				continue;

			{
				var bottoms = new List<TableHelperCell>();
				for (var i = 0; i < colspan; i++)
				{
					var tmp = GetCell(r + 1, c + i);
					if (tmp?.Blank == false) bottoms.Add(tmp);
				}

				foreach (var bottom in bottoms)
				{
					bottom.Borders ??= ApplyCellBorders(
						bottom.Cell?.Descendants<TableCellBorders>().FirstOrDefault(),
						bottom.ColId == 0 || bottom.RowStart,
						bottom.ColId + GetColSpan(bottom.CellId) == ColumnLength || bottom.RowEnd,
						bottom.RowId == 0,
						bottom.RowId + GetRowSpan(bottom.CellId) == RowLength
					);

					var meWin = CompareBorder(me.Borders, bottom.Borders, CompareDirection.Vertical);
					if (meWin) me.Borders.BottomBorder?.CopyAttributesTo(bottom.Borders.TopBorder);
				}
			}
		}

		if (cells.Count == 0)
			return;

		{
			for (var i = 0; i < cells[^1].CellId; i++)
			{
				var me = GetCellByCellId(i);
				if (me?.Blank != false) // ignore gridBefore/gridAfter cells
					continue;

				if (me.RowSpan > 1)
				{
					// merge bottom border from the last cell of row-spanned cells
					var meRowEnd = GetCell(me.RowId + (me.RowSpan - 1), me.ColId);
					meRowEnd?.Borders?.BottomBorder?.CopyAttributesTo(me.Borders?.BottomBorder);
				}

				if (me.RowId + me.RowSpan >= RowLength)
					continue;

				// if me is not at the last row, compare the border with the cell below it
				var below = GetCell(me.RowId + me.RowSpan, me.ColId);
				if (below == null) continue;
				var bottom = GetCellByCellId(below.CellId);
				var meWin = CompareBorder(me.Borders, bottom?.Borders, CompareDirection.Vertical);
				if (!meWin) me.Borders?.BottomBorder?.ClearAllAttributes();
			}
		}
	}

	/// <summary>
	/// Compare border and return who is win.
	/// </summary>
	/// <param name="a"></param>
	/// <param name="b"></param>
	/// <param name="dir"></param>
	/// <returns>Return ture means a win, false means b win.</returns>
	static bool CompareBorder(TableCellBorders? a, TableCellBorders? b, CompareDirection dir)
	{
		if (a == null || b == null) return false;
		// compare line style
		int weight1 = 0, weight2 = 0;
		switch (dir)
		{
			case CompareDirection.Horizontal:
			{
				if (a.RightBorder?.Val != null)
					weight1 = BorderNumber.ContainsKey(a.RightBorder.Val) ? BorderNumber[a.RightBorder.Val] : 1;
				else if (a.InsideVerticalBorder?.Val != null)
					weight1 = BorderNumber.ContainsKey(a.InsideVerticalBorder.Val) ? BorderNumber[a.InsideVerticalBorder.Val] : 1;

				if (b.LeftBorder?.Val != null)
					weight2 = BorderNumber.ContainsKey(b.LeftBorder.Val) ? BorderNumber[b.LeftBorder.Val] : 1;
				else if (b.InsideVerticalBorder?.Val != null)
					weight2 = BorderNumber.ContainsKey(b.InsideVerticalBorder.Val) ? BorderNumber[b.InsideVerticalBorder.Val] : 1;

				break;
			}
			case CompareDirection.Vertical:
			{
				if (a.BottomBorder?.Val != null)
					weight1 = BorderNumber.ContainsKey(a.BottomBorder.Val) ? BorderNumber[a.BottomBorder.Val] : 1;
				else if (a.InsideHorizontalBorder?.Val != null)
					weight1 = BorderNumber.ContainsKey(a.InsideHorizontalBorder.Val) ? BorderNumber[a.InsideHorizontalBorder.Val] : 1;

				if (b.TopBorder?.Val != null)
					weight2 = BorderNumber.ContainsKey(b.TopBorder.Val) ? BorderNumber[b.TopBorder.Val] : 1;
				else if (b.InsideHorizontalBorder?.Val != null)
					weight2 = BorderNumber.ContainsKey(b.InsideHorizontalBorder.Val) ? BorderNumber[b.InsideHorizontalBorder.Val] : 1;

				break;
			}
		}

		if (weight1 > weight2) return true;

		if (weight2 > weight1) return false;

		// compare width
		float size1 = 0f, size2 = 0f;
		switch (dir)
		{
			case CompareDirection.Horizontal:
			{
				if (a.RightBorder?.Size?.HasValue == true)
					size1 = Converter.OneEighthPointToPoint(a.RightBorder.Size.Value);
				else if (a.InsideVerticalBorder?.Size?.HasValue == true)
					size1 = Converter.OneEighthPointToPoint(a.InsideVerticalBorder.Size.Value);

				if (b.LeftBorder?.Size?.HasValue == true)
					size2 = Converter.OneEighthPointToPoint(b.LeftBorder.Size.Value);
				else if (b.InsideVerticalBorder?.Size?.HasValue == true)
					size2 = Converter.OneEighthPointToPoint(b.InsideVerticalBorder.Size.Value);

				break;
			}
			case CompareDirection.Vertical:
			{
				if (a.BottomBorder?.Size?.HasValue == true)
					size1 = Converter.OneEighthPointToPoint(a.BottomBorder.Size.Value);
				else if (a.InsideHorizontalBorder?.Size?.HasValue == true)
					size1 = Converter.OneEighthPointToPoint(a.InsideHorizontalBorder.Size.Value);

				if (b.TopBorder?.Size?.HasValue == true)
					size2 = Converter.OneEighthPointToPoint(b.TopBorder.Size.Value);
				else if (b.InsideHorizontalBorder?.Size?.HasValue == true)
					size2 = Converter.OneEighthPointToPoint(b.InsideHorizontalBorder.Size.Value);

				break;
			}
		}

		if (size1 > size2) return true;

		if (size2 > size1) return false;

		// compare brightness
		//   TODO: current brightness implementation is based on Luminance 
		//   but ISO $17.4.66 defines the comparisons should be
		//   1. R+B+2G, 2. B+2G, 3. G
		float brightness1 = 0f, brightness2 = 0f;
		switch (dir)
		{
			case CompareDirection.Horizontal:
			{
				if (a.RightBorder?.Color?.HasValue == true)
					brightness1 = Tools.RgbBrightness(a.RightBorder?.Color?.Value ?? "");
				else if (a.InsideVerticalBorder?.Color?.HasValue == true)
					brightness1 = Tools.RgbBrightness(a.InsideVerticalBorder?.Color?.Value ?? "");

				if (b.LeftBorder?.Color?.HasValue == true)
					brightness2 = Tools.RgbBrightness(b.LeftBorder?.Color?.Value ?? "");
				else if (b.InsideVerticalBorder?.Color?.HasValue == true)
					brightness2 = Tools.RgbBrightness(b.InsideVerticalBorder?.Color?.Value ?? "");

				break;
			}
			case CompareDirection.Vertical:
			{
				if (a.BottomBorder?.Color?.HasValue == true)
					brightness1 = Tools.RgbBrightness(a.BottomBorder?.Color?.Value ?? "");
				else if (a.InsideHorizontalBorder?.Color?.HasValue == true)
					brightness1 = Tools.RgbBrightness(a.InsideHorizontalBorder?.Color?.Value ?? "");

				if (b.TopBorder?.Color?.HasValue == true)
					brightness2 = Tools.RgbBrightness(b.TopBorder?.Color?.Value ?? "");
				else if (b.InsideHorizontalBorder?.Color?.HasValue == true)
					brightness2 = Tools.RgbBrightness(b.InsideHorizontalBorder?.Color?.Value ?? "");

				break;
			}
		}

		// smaller brightness wins
		if (brightness1 < brightness2) return true;

		return brightness2 < brightness1 && false;
	}

	void AdjustTableColumnsWidth()
	{
		if (table?.GetMainDocumentPart() is not MainDocumentPart mainDocumentPart)
			return;
		var printablePageWidth = Converter.TwipToPoint(mainDocumentPart.GetPrintablePageWidth());

		// Get table total width
		var totalWidth = -1f;
		var autoWidth = false;
		var tableWidth = table.GetEffectiveElement<TableWidth>();
		if (tableWidth?.Type != null)
			switch (tableWidth.Type.Value)
			{
				default:
					autoWidth = true;
					break;
				case TableWidthUnitValues.Dxa:
					if (tableWidth.Width != null) totalWidth = Converter.TwipToPoint(tableWidth.Width.Value ?? "");

					break;
				case TableWidthUnitValues.Pct:
					if (tableWidth.Width != null)
						totalWidth = printablePageWidth * Tools.Percentage(tableWidth.Width.Value ?? "");

					//if (table.Parent.GetType() == typeof(Wordprocessing.Body))
					//    totalWidth = (float)((pdfDoc.PageSize.Width - pdfDoc.LeftMargin - pdfDoc.RightMargin) * percentage(tableWidth.Width.Value));
					//else
					//    totalWidth = this.getCellWidth(table.Parent as Wordprocessing.TableCell) * percentage(tableWidth.Width.Value);
					break;
			}

		if (!autoWidth)
			ScaleTableColumnsWidth(ref tableColumnsWidth, totalWidth);
		else
			totalWidth = tableColumnsWidth?.Sum() ?? 0.0f;

		for (var i = 0; i < RowLength; i++)
		{
			// Get all cells in this row
			var cellsInRow = cells.FindAll(c => c.RowId == i);
			if (cellsInRow.Count == 0) continue;

			// Get if any gridBefore & gridAfter
			int skipGridsBefore = 0, skipGridsAfter = 0;
			float skipGridsBeforeWidth = 0f, skipGridsAfterWidth = 0f;
			var head = cellsInRow.Find(c => c.RowStart);
			if (head?.Row?.TableRowProperties != null)
			{
				// w:gridBefore
				var tmpGridBefore = head.Row.TableRowProperties.Elements<GridBefore>().FirstOrDefault();
				if (tmpGridBefore?.Val != (object?)null) skipGridsBefore = tmpGridBefore.Val.Value;

				// w:wBefore
				var tmpGridBeforeWidth = head.Row.TableRowProperties.Elements<WidthBeforeTableRow>()
					.FirstOrDefault();
				if (tmpGridBeforeWidth?.Width != null)
					skipGridsBeforeWidth = Converter.TwipToPoint(Convert.ToInt32(tmpGridBeforeWidth.Width.Value, CultureInfo.InvariantCulture));

				// w:gridAfter
				var tmpGridAfter = head.Row.TableRowProperties.Elements<GridAfter>().FirstOrDefault();
				if (tmpGridAfter?.Val != (object?)null) skipGridsAfter = tmpGridAfter.Val.Value;

				// w:wAfter
				var tmpGridAfterWidth = head.Row.TableRowProperties.Elements<WidthAfterTableRow>()
					.FirstOrDefault();
				if (tmpGridAfterWidth?.Width != null)
					skipGridsAfterWidth = Converter.TwipToPoint(Convert.ToInt32(tmpGridAfterWidth.Width.Value, CultureInfo.InvariantCulture));
			}

			var j = 0;

			// -------
			// gridBefore
			var edgeEnd = skipGridsBefore;
			for (; j < edgeEnd; j++)
			{
				// deduce specific columns width from required width
				skipGridsBeforeWidth -= tableColumnsWidth?[j] ?? 0.0f;
			}

			if (skipGridsBeforeWidth > 0f)
			{
				// if required width is larger than the total width of specific columns,
				// the remaining required width adds to the last specific column 
				if (tableColumnsWidth != null)
					tableColumnsWidth[edgeEnd - 1] += skipGridsBeforeWidth;
			}

			// ------
			// cells
			while (j < cellsInRow.Count - skipGridsAfter)
			{
				var reqCellWidth = 0f;
				var cellWidth = cellsInRow[j].Cell?.GetEffectiveElement<TableCellWidth>();
				if (cellWidth?.Type != null)
				{
					switch (cellWidth.Type.Value)
					{
						case TableWidthUnitValues.Auto:
							//// TODO: calculate the items width
							//if (cellsInRow[j].elements.Count > 0)
							//{
							//    Text.IElement element = cellsInRow[j].elements[0];
							//}
							break;
						case TableWidthUnitValues.Dxa:
							if (cellWidth.Width != null)
								reqCellWidth = Converter.TwipToPoint(cellWidth.Width.Value ?? "");

							break;
						case TableWidthUnitValues.Pct:
							if (cellWidth.Width != null)
								reqCellWidth = Tools.Percentage(cellWidth.Width.Value ?? "") * totalWidth;

							break;
					}
				}

				// check row span
				var spanCount = 1;
				if (cellsInRow[j].Cell != null)
				{
					var tmpCell = cellsInRow[j].Cell;
					if (tmpCell?.TableCellProperties != null)
					{
						var span = tmpCell.TableCellProperties.Elements<GridSpan>().FirstOrDefault();
						spanCount = span?.Val != (object?)null ? span.Val.Value : 1;
					}
				}

				edgeEnd = j + spanCount;
				for (; j < edgeEnd; j++)
				{
					// deduce specific columns width from required width
					reqCellWidth -= tableColumnsWidth?[j] ?? 0.0f;
				}

				if (reqCellWidth <= 0f) continue;
				// if required width is larger than the total width of specific columns,
				// the remaining required width adds to the last specific column 
				if (tableColumnsWidth != null)
					tableColumnsWidth[edgeEnd - 1] += reqCellWidth;
			}

			// ------
			// gridAfter
			edgeEnd = j + skipGridsAfter;
			for (; j < edgeEnd; j++)
				// deduce specific columns width from required width
				skipGridsAfterWidth -= tableColumnsWidth?[j] ?? 0.0f;

			if (skipGridsAfterWidth > 0f)
				// if required width is larger than the total width of specific columns,
				// the remaining required width adds to the last specific column 
				if (tableColumnsWidth != null)
					tableColumnsWidth[edgeEnd - 1] += skipGridsAfterWidth;

			if (!autoWidth) // fixed table width, adjust width to fit in
				ScaleTableColumnsWidth(ref tableColumnsWidth, totalWidth);
			else // auto table width
				totalWidth = tableColumnsWidth?.Sum() ?? 0.0f;
		}
	}

	static void ScaleTableColumnsWidth(ref float[]? columns, float totalWidth)
	{
		if (columns == null) return;
		var sum = columns.Sum();
		if (sum <= totalWidth)
			return;

		var ratio = totalWidth / sum;
		for (var j = 0; j < columns.Length; j++) columns[j] *= ratio;
	}

	/// <summary>
	/// Get table gridCols information and convert to twip to Points.
	/// </summary>
	/// <param name="table"></param>
	/// <returns></returns>
	static float[]? GetTableGridCols(OpenXmlElement table)
	{
		float[]? grids;

		// Get grids and their width
		var grid = table.Elements<TableGrid>().FirstOrDefault();
		if (grid != null)
		{
			var gridCols = grid.Elements<GridColumn>().ToList();
			if (gridCols.Count == 0)
				return Array.Empty<float>();

			grids = new float[gridCols.Count];
			for (var i = 0; i < gridCols.Count; i++)
			{
				if (gridCols[i].Width != null)
					grids[i] = Converter.TwipToPoint(gridCols[i].Width?.Value ?? "");
				else
					grids[i] = 0f;
			}
		}
		else
		{
			if (table.GetMainDocumentPart() is not MainDocumentPart mainDocumentPart) return null;
			var body = mainDocumentPart.Document.Body;
			if (body == null) return null;
			var section = body.Descendants<SectionProperties>().FirstOrDefault();
			var size = section?.Descendants<PageSize>().FirstOrDefault();
			var margin = section?.Descendants<PageMargin>().FirstOrDefault();
			var printablePageWidth = 0.0f;
			if (size != null && size.Width != (object?)null && margin != null && margin.Left != (object?)null && margin.Right != (object?)null)
				printablePageWidth = Converter.TwipToPoint(size.Width.Value - margin.Left.Value - margin.Right.Value);
			var tableWidth = table.GetEffectiveElement<TableWidth>();
			if (tableWidth?.Type == null) return Array.Empty<float>();
			if (tableWidth.Type.Value != TableWidthUnitValues.Pct)
				return Array.Empty<float>();

			var totalWidth = printablePageWidth * Tools.Percentage(tableWidth.Width?.Value ?? "");
			var row = table.Elements<TableRow>().FirstOrDefault();
			if (row == null) return Array.Empty<float>();
			var cells = row.Elements<TableCell>().ToList();
			grids = new float[cells.Count];
			for (var i = 0; i < cells.Count; i++)
			{
				var cellWidth = cells[i].Descendants<TableCellWidth>().FirstOrDefault();
				if (cellWidth?.Type == null || cellWidth.Type.Value != TableWidthUnitValues.Pct)
					return null;
				grids[i] = totalWidth * Tools.Percentage(cellWidth.Width?.Value ?? "");
			}
		}

		return grids;
	}

	/// <summary>
	/// Rolling up table border property from TableProperties > TableStyle > Default style.
	/// </summary>
	void RollingUpTableBorders()
	{
		if (table == null) return;

		condFmtbr = new TableBorders();

		var borders = new List<TableBorders>();
		var tblPrs = table.Elements<TableProperties>().FirstOrDefault();
		if (tblPrs != null)
		{
			// get from table properties (priority 1)
			var tmp = tblPrs.Descendants<TableBorders>().FirstOrDefault();
			if (tmp != null) borders.Insert(0, tmp);

			// get from styles (priority 2)
			if (tblPrs.TableStyle?.Val != null)
			{
				var st = tblPrs.TableStyle.GetStyleById();
				while (st != null)
				{
					tmp = st.Descendants<TableBorders>().FirstOrDefault();
					if (tmp != null) borders.Insert(0, tmp);

					st = st.BasedOn?.Val != null
						? st.BasedOn.GetStyleById()
						: null;
				}
			}

			// get from default table style (priority 3)
			if (table.GetMainDocumentPart() is not MainDocumentPart mainDocumentPart)
				return;
			var defaultTableStyle = mainDocumentPart.GetDefaultStyle(DefaultStyleType.Table);
			if (defaultTableStyle != null)
			{
				tmp = defaultTableStyle.Descendants<TableBorders>().FirstOrDefault();
				if (tmp != null) borders.Insert(0, tmp);
			}
		}

		foreach (var border in borders)
		{
			if (border.TopBorder != null)
			{
				if (condFmtbr.TopBorder == null)
					condFmtbr.TopBorder = (TopBorder)border.TopBorder.CloneNode(true);
				else
					border.TopBorder.CopyAttributesTo(condFmtbr.TopBorder);
			}

			if (border.BottomBorder != null)
			{
				if (condFmtbr.BottomBorder == null)
					condFmtbr.BottomBorder = (BottomBorder)border.BottomBorder.CloneNode(true);
				else
					border.BottomBorder.CopyAttributesTo(condFmtbr.BottomBorder);
			}

			if (border.LeftBorder != null)
			{
				if (condFmtbr.LeftBorder == null)
					condFmtbr.LeftBorder = (LeftBorder)border.LeftBorder.CloneNode(true);
				else
					border.LeftBorder.CopyAttributesTo(condFmtbr.LeftBorder);
			}

			if (border.RightBorder != null)
			{
				if (condFmtbr.RightBorder == null)
					condFmtbr.RightBorder = (RightBorder)border.RightBorder.CloneNode(true);
				else
					border.RightBorder.CopyAttributesTo(condFmtbr.RightBorder);
			}

			if (border.InsideHorizontalBorder != null)
			{
				if (condFmtbr.InsideHorizontalBorder == null)
				{
					condFmtbr.InsideHorizontalBorder =
						(InsideHorizontalBorder)border.InsideHorizontalBorder.CloneNode(true);
				}
				else
				{
					border.InsideHorizontalBorder.CopyAttributesTo(condFmtbr.InsideHorizontalBorder);
				}
			}

			if (border.InsideVerticalBorder == null)
				continue;

			if (condFmtbr.InsideVerticalBorder == null)
			{
				condFmtbr.InsideVerticalBorder =
					(InsideVerticalBorder)border.InsideVerticalBorder.CloneNode(true);
			}
			else
			{
				border.InsideVerticalBorder.CopyAttributesTo(condFmtbr.InsideVerticalBorder);
			}
		}
	}

	/// <summary>
	/// Applies the cell borders.
	/// </summary>
	/// <param name="cellbr"></param>
	/// <param name="firstColumn">Is current cell at the first column?</param>
	/// <param name="lastColumn">Is current cell at the last column?</param>
	/// <param name="firstRow">Is current cell at the first row?</param>
	/// <param name="lastRow">Is current cell at the first row?</param>
	/// <returns></returns>
	TableCellBorders ApplyCellBorders(TableCellBorders? cellbr,
		bool firstColumn, bool lastColumn, bool firstRow, bool lastRow)
	{
		var ret = new TableCellBorders
		{
			TopBorder = new TopBorder(),
			BottomBorder = new BottomBorder(),
			LeftBorder = new LeftBorder(),
			RightBorder = new RightBorder(),
			InsideHorizontalBorder = new InsideHorizontalBorder(),
			InsideVerticalBorder = new InsideVerticalBorder(),
			TopLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder(),
			TopRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder()
		};

		// cell border first, if no cell border then conditional formatting (table border + style)
		if (cellbr?.TopBorder != null)
		{
			cellbr.TopBorder.CopyAttributesTo(ret.TopBorder);
		}
		else
		{
			condFmtbr?.TopBorder?.CopyAttributesTo(ret.TopBorder);

			if (!firstRow) condFmtbr?.InsideHorizontalBorder?.CopyAttributesTo(ret.TopBorder);
		}

		if (cellbr?.BottomBorder != null)
		{
			cellbr.BottomBorder.CopyAttributesTo(ret.BottomBorder);
		}
		else
		{
			condFmtbr?.BottomBorder?.CopyAttributesTo(ret.BottomBorder);

			if (!lastRow) condFmtbr?.InsideHorizontalBorder?.CopyAttributesTo(ret.BottomBorder);
		}

		if (cellbr?.LeftBorder != null)
		{
			cellbr.LeftBorder.CopyAttributesTo(ret.LeftBorder);
		}
		else
		{
			condFmtbr?.LeftBorder?.CopyAttributesTo(ret.LeftBorder);

			if (!firstColumn) condFmtbr?.InsideVerticalBorder?.CopyAttributesTo(ret.LeftBorder);
		}

		if (cellbr?.RightBorder != null)
		{
			cellbr.RightBorder.CopyAttributesTo(ret.RightBorder);
		}
		else
		{
			condFmtbr?.RightBorder?.CopyAttributesTo(ret.RightBorder);

			if (!lastColumn) condFmtbr?.InsideVerticalBorder?.CopyAttributesTo(ret.RightBorder);
		}

		cellbr?.TopLeftToBottomRightCellBorder?.CopyAttributesTo(ret.TopLeftToBottomRightCellBorder);

		cellbr?.TopRightToBottomLeftCellBorder?.CopyAttributesTo(ret.TopRightToBottomLeftCellBorder);

		return ret;
	}

	/// <summary>
	/// Get column span number of cell.
	/// </summary>
	/// <param name="cellId">Cell ID.</param>
	/// <returns></returns>
	public int GetColSpan(int cellId)
	{
		var colspan = 0;
		for (var i = 0; i < cells.Count; i++)
		{
			if (cells[i].CellId != cellId)
				continue;

			colspan++;

			var endOfRow = (i / ColumnLength + 1) * ColumnLength;
			while (i + 1 < endOfRow)
			{
				i++;
				if (cells[i].CellId == cellId)
					colspan++;
				else
					break;
			}

			break;
		}

		return colspan;
	}

	/// <summary>
	/// Get row span number of cell.
	/// </summary>
	/// <param name="cellId">Cell ID.</param>
	/// <returns></returns>
	public int GetRowSpan(int cellId)
	{
		var rowspan = 0;
		for (var i = 0; i < cells.Count; i++)
		{
			if (cells[i].CellId != cellId)
				continue;

			rowspan++;

			while (i + ColumnLength < cells.Count)
			{
				i += ColumnLength;
				if (cells[i].CellId == cellId)
					rowspan++;
				else
					break;
			}

			break;
		}

		return rowspan;
	}

	/// <summary>
	/// Get cell object by position (row x column).
	/// </summary>
	/// <param name="row">Row number.</param>
	/// <param name="col">Column number.</param>
	/// <returns></returns>
	TableHelperCell? GetCell(int row, int col)
	{
		var index = row * ColumnLength + col;
		return index < cells.Count ? cells[index] : null;
	}

	/// <summary>
	/// Get cell object by cellId.
	/// </summary>
	/// <param name="cellId">Cell Id</param>
	/// <returns></returns>
	TableHelperCell? GetCellByCellId(int cellId)
	{
		var cellsInRow = cells.FindAll(c => c.CellId == cellId).OrderBy(o => o.RowId).ToList();
		return cellsInRow.Count > 0
			? cellsInRow.FindAll(c => c.RowId == cellsInRow[0].RowId).OrderBy(o => o.ColId).ToList()[0]
			: null;
	}

	/// <summary>
	/// Get cell width, this should be called after ParseTable().
	/// </summary>
	/// <param name="cell"></param>
	/// <returns></returns>
	public float GetCellWidth(TableHelperCell cell)
	{
		var ret = 0f;

		for (var i = 0; i < cell.ColSpan; i++) ret += TableColumnsWidth?[cell.ColId + i] ?? 0.0f;

		return ret;
	}

	enum CompareDirection
	{
		Horizontal,
		Vertical
	}
}
