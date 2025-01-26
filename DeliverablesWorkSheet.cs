using OfficeOpenXml;

public class DeliverablesWorkSheet : WorksheetBase
{
    public override string[] header => new string[]
    {
        "客戶名稱", "專案名稱"
    };
    public override int[] widthAlignment => new int[]
    {
        30, 30
    };

    public override int lastRow { get; set; }

    public DeliverablesWorkSheet(ExcelWorksheet sheet) : base(sheet)
    {
        lastRow = 1;
    }

    // 覆寫 Dispose 方法
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // 釋放子類別特定的資源（如果有）
        }

        // 呼叫基類的 Dispose 方法
        base.Dispose(disposing);
    }

    public override void WriteAndFormatHeader(int startRow, int endRow)
    {
        FormatCells(sheet.Cells[$"{startCol}{startRow}:{endCol}{endRow - 1}"], isHeader: true, isHair: true);
    }

    public override void FormatPhase()
    {
    }

    public override void WriteFooter()
    {
        AlignColumnWidth();
    }
}
