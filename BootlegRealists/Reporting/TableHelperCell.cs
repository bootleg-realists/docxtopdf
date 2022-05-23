using DocumentFormat.OpenXml.Wordprocessing;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class contains helper functions for the table cell.
/// </summary>
public class TableHelperCell
{
	readonly TableHelper owner;

	/// <summary>
	/// Constructor of class
	/// </summary>
	/// <param name="owner">owner of this class</param>
	/// <param name="cellId">Cell identifier</param>
	/// <param name="rowId">Row identifier</param>
	/// <param name="colId">Column identifier</param>
	public TableHelperCell(TableHelper owner, int cellId, int rowId, int colId)
	{
		this.owner = owner;
		CellId = cellId;
		RowId = rowId;
		ColId = colId;
	}

	/// <summary>
	/// Borders of the cell
	/// </summary>
	public TableCellBorders? Borders { get; set; }

	/// <summary>
	/// Points to the cell
	/// </summary>
	public TableCell? Cell { get; set; } // point to Wordprocessing.TableCell

	/// <summary>
	/// Cell ID.
	/// </summary>
	public int CellId { get; set; }

	/// <summary>
	/// Cell's column number.
	/// </summary>
	public int ColId { get; }

	/// <summary>
	/// Points to the row
	/// </summary>
	public TableRow? Row { get; set; } // point to Wordprocessing.TableRow

	/// <summary>
	/// Get whether this cell is the end of the row.
	/// rowEnd is set to the main cell of col-spanned cells (i.e. the first cell of col-spanned cells).
	/// gridAfter cells do not set as rowEnd.
	/// </summary>
	public bool RowEnd { get; set; }

	/// <summary>
	/// Cell's row number.
	/// </summary>
	public int RowId { get; }

	/// <summary>
	/// Get whether this cell is the start of the row.
	/// rowStart is set to the main cell of col-spanned cells (i.e. the first cell of col-spanned cells).
	/// gridBefore cells do not set as rowStart.
	/// </summary>
	public bool RowStart { get; set; }

	/// <summary>
	/// Get whether this cell is blank.
	/// </summary>
	public bool Blank => Cell == null;

	/// <summary>
	/// Get cell's row span.
	/// </summary>
	public int RowSpan => Blank ? 0 : owner.GetRowSpan(CellId);

	/// <summary>
	/// Get cell's column span.
	/// </summary>
	public int ColSpan => Blank ? 0 : owner.GetColSpan(CellId);
}
