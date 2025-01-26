using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.Common;

public class TaskListWorksheet : WorksheetBase
{
    public override string[] header => new string[]
    {
        "工作編號", "大綱編號", "工作項目", "工作說明", "工天", "預計開始日期", "預計完成日期", "完成度", "負責單位", "交付文件", "備註"
    };

    public override int[] widthAlignment => new int[]
    {
        9, 9, 30, 75, 9, 15, 15, 15, 15, 15, 15
    };

    public override int lastRow { get; set; }

    public TaskListWorksheet(ExcelWorksheet sheet) : base(sheet)
    {
        taskListRefSheet = sheet.Name;
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

    // 寫入標題文字
    public override void WriteAndFormatHeader(int startRow, int endRow)
    {
        FormatCells(sheet.Cells[$"{startCol}{startRow}:{endCol}{endRow - 1}"], isHeader: true, isHair: true, isBorder: true);
    }

    public override void FormatPhase()
    {
        // 格式化內容
        MergeAndAlign(phaseStartRow + lastRow, 1, phaseEndRow + lastRow - 1, 1);
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseEndRow + lastRow - 1}"], isHair: true, isBorder: true);
        // 使用 RGB 自定義顏色格式化階段標題
        Color titleBgColor = Color.FromArgb(198, 224, 180);  // 淺綠色
        MergeAndAlign(phaseTitleRow + lastRow, 1, phaseTitleRow + lastRow, header.Length);
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseTitleRow + lastRow}"], bgColor: titleBgColor, isTitle: true);
    }

    // 寫入文字資料
    public override void WriteText(string text, int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 lastRow + sharedRow 來決定列數
        currentColumn = column;
        if (text=="無")
        {
            text = "";
        }        
        sheet.Cells[currentRow, column].Value = text;
        CenterText(column, isCenter);        
        referencedRow = currentRow;
    }

    // 寫入數字資料
    public override void WriteText(double number, int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 lastRow + sharedRow 來決定列數
        currentColumn = column;
        sheet.Cells[currentRow, column].Value = number;
        CenterText(column, isCenter);
        referencedRow = currentRow;
    }

    // 寫入公式
    public override void WriteFormula(string formula, int column)
    {
        // 使用 lastRow + sharedRow 來決定列數
        sheet.Cells[currentRow, column].Formula = formula;
        referencedRow = currentRow;
    }

    // 寫入頁腳文字
    public override void WriteFooter()
    {
        sheet.Cells[currentRow, 1].Value = $"總計工作天數";
        sheet.Cells[currentRow, 5].Formula = $"SUM(E3:E{currentRow-1})";
        sheet.Cells[currentRow, 5].Style.Font.Bold = true ;
        sheet.Cells[currentRow, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        Color footerBgColor = Color.FromArgb(248, 203, 173);    // 淺橘紅色
        FormatCells(sheet.Cells[$"{startCol}{currentRow}:{endCol}{currentRow}"], bgColor: footerBgColor, isBorder: true);
        // 在第 1 行插入 1 行空白行
        sheet.InsertRow(1, 1);
        sheet.Cells[1, 1].Style.Font.Size = 14;
        MergeAndAlign(1, 1, 1, header.Length, isCenter: true);
        AlignColumnWidth();
    }
}
