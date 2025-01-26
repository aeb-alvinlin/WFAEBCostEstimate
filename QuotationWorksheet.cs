using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data.Common;

public class QuotationWorksheet : WorksheetBase
{
    // 寫入標題行
    public override string[] header => new string[]
    {
        "項次", "工作項目", "工作說明", "數量", "總價(NT$)", "備註"
    };

    public override int[] widthAlignment => new int[] 
    { 
        10, 29, 54, 9, 18, 12 
    };

    public override int lastRow { get; set; }
    private bool includeArchitectColumn;
    int mandaySubtotal = 5; // 工作天數(小計)
    int projectManagerColumn = 6; // F 欄
    int architectColumn = default; // H 欄
    int deployerColumn = default; // J 或 H 欄
    int developerColumn = default; // L 或 J 欄
    int totalCostColumn = default; // N 或 L 欄

    public QuotationWorksheet(ExcelWorksheet sheet, bool withArchitectColumn) : base(sheet)
    {
        lastRow = 1;
        includeArchitectColumn = withArchitectColumn;
        architectColumn = includeArchitectColumn ? 8 : -1; // H 欄
        deployerColumn = includeArchitectColumn ? 10 : 8; // J 或 H 欄
        developerColumn = includeArchitectColumn ? 12 : 10; // L 或 J 欄
        totalCostColumn = includeArchitectColumn ? 14 : 12; // N 或 L 欄
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

    // 寫入標題行
    public override void WriteAndFormatHeader(int startRow, int endRow)
    {
        FormatCells(sheet.Cells[$"{startCol}{startRow}:{endCol}{endRow - 1}"], isHeader: true, isHair: true, isBorder: true);
    }

    public override void FormatPhase()
    {
        // 格式化內容
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseEndRow + lastRow - 1}"], isThin: true, isBorder: true);
        // 使用 RGB 自定義顏色格式化階段標題
        Color titleBgColor = Color.FromArgb(198, 224, 180);  // 淺綠色
        MergeAndAlign(phaseTitleRow + lastRow, 1, phaseTitleRow + lastRow, Array.IndexOf(header, "工作說明") + 1);
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseTitleRow + lastRow}"], bgColor: titleBgColor, isTitle: true);
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "數量") + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;  // 內文靠右
    }

    public void PhaseSumPrice(int phaseCount)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "數量") + 1].Value = 1;
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "總價(NT$)") + 1].Formula = $"SUM({projCostRefSheet}!{(char)('A' + totalCostColumn - 1)}{referProjCostRow}:{(char)('A' + totalCostColumn - 1)}{referProjCostRow + phaseCount})";
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.Numberformat.Format = "$#,##0";
    }

    public void WriteCostSumFormula(int column)
    {        
        sheet.Cells[currentRow, column].Style.Numberformat.Format = "$#,##0";
        sheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right; // 內文靠右
        sheet.Cells[currentRow, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }

    public override void WriteFooter()
    {
        int projectManagerColumn = 6; // F 欄
        int architectColumn = includeArchitectColumn ? 8 : -1; // H 欄
        int deployerColumn = includeArchitectColumn ? 10 : 8; // J 或 H 欄
        int developerColumn = includeArchitectColumn ? 12 : 10; // L 或 J 欄
        int totalCostColumn = includeArchitectColumn ? 14 : 12; // N 或 L 欄
        sheet.Cells[currentRow, 1].Value = $"總計";
        sheet.Cells[currentRow, 1].Style.Font.Bold = true;
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Formula = $"{projCostRefSheet}!{(char)('A' + totalCostColumn - 1)}{referProjCostRow+1}";
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.Font.Bold = true;
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.Numberformat.Format = "$#,##0";
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        Color footerBgColor = Color.FromArgb(248, 203, 173);    // 淺橘紅色
        FormatCells(sheet.Cells[$"{startCol}{currentRow}:{endCol}{currentRow}"], bgColor: footerBgColor, isBorder: true);
        // 深藍色
        footerBgColor = Color.FromArgb(7, 79, 105);     
        sheet.Cells[currentRow + 2, 1].Value = "專案成員人天分配：";
        MergeAndAlign(currentRow + 2, 1, currentRow + 2, 3, isCenter: false);
        FormatCells(sheet.Cells[$"A{currentRow + 2}:C{currentRow + 2}"], bgColor: footerBgColor, fontColor: Color.White, isBorder: true);

        // 動態設置每個成員的數據
        int allocationRow = currentRow + 3;
        sheet.Cells[allocationRow, 1].Value = "處理人員";
        sheet.Cells[allocationRow, 2].Value = "天數";
        sheet.Cells[allocationRow, 3].Value = "角色";
        allocationRow++;
        // 專案經理
        sheet.Cells[allocationRow, 1].Value = "A";
        sheet.Cells[allocationRow, 2].Formula = $"={projCostRefSheet}!{(char)('A' + projectManagerColumn - 1)}{referProjCostRow+1}";
        sheet.Cells[allocationRow, 3].Value = "專案經理";
        FormatCells(sheet.Cells[allocationRow, 2], isRight: true);
        allocationRow++;
        // 架構師"
        if (includeArchitectColumn)
        {
            sheet.Cells[allocationRow, 1].Value = "B";
            sheet.Cells[allocationRow, 2].Formula = $"={projCostRefSheet}!{(char)('A' + architectColumn - 1)}{referProjCostRow + 1}";
            sheet.Cells[allocationRow, 3].Value = "架構師";
            FormatCells(sheet.Cells[allocationRow, 2], isRight: true);
            allocationRow++;
        }
        // 部署者
        sheet.Cells[allocationRow, 1].Value = "C";
        sheet.Cells[allocationRow, 2].Formula = $"={projCostRefSheet}!{(char)('A' + deployerColumn - 1)}{referProjCostRow+1}";
        sheet.Cells[allocationRow, 3].Value = "部署者";
        FormatCells(sheet.Cells[currentRow + 6, 2], isRight: true);
        allocationRow++;
        // 開發者
        sheet.Cells[allocationRow, 1].Value = "D";
        sheet.Cells[allocationRow, 2].Formula = $"={projCostRefSheet}!{(char)('A' + developerColumn - 1)}{referProjCostRow+1}";
        sheet.Cells[allocationRow, 3].Value = "開發者";
        FormatCells(sheet.Cells[allocationRow, 2], isRight: true);
        // 格式化表格有表格格線
        FormatCells(sheet.Cells[$"A{currentRow + 3}:C{allocationRow}"], isThin: true, isBorder: true);
        // 淡藍色
        footerBgColor = Color.FromArgb(202, 237, 251);
        FormatCells(sheet.Cells[$"C{currentRow + 3}:C{allocationRow}"], bgColor: footerBgColor);
        // 淺藍色
        footerBgColor = Color.FromArgb(97, 203, 243);
        FormatCells(sheet.Cells[$"A{currentRow + 3}:C{currentRow + 3}"], bgColor: footerBgColor, isBorder: true);
        allocationRow += 2;
        // 深藍色
        footerBgColor = Color.FromArgb(7, 79, 105);
        sheet.Cells[allocationRow, 1].Value = "預估專案期間(週)";
        MergeAndAlign(allocationRow, 1, allocationRow, 3, isCenter: false);
        FormatCells(sheet.Cells[$"A{allocationRow}:C{allocationRow}"], bgColor: footerBgColor, fontColor: Color.White, isBorder: true);
        allocationRow++;
        //
        sheet.Cells[allocationRow, 1].Value = "預估開始日 :";
        MergeAndAlign(allocationRow, 1, allocationRow, 2, isCenter: false);
        allocationRow++;
        sheet.Cells[allocationRow, 1].Value = "預估結束日 :";
        MergeAndAlign(allocationRow, 1, allocationRow, 2, isCenter: false);
        allocationRow++;
        sheet.Cells[allocationRow, 1].Value = "預估週期 :";
        MergeAndAlign(allocationRow, 1, allocationRow, 2, isCenter: false);
        // 淺藍色
        FormatCells(sheet.Cells[$"A{allocationRow - 2}:C{allocationRow}"], isThin: true, isBorder: true);
        // 淺藍色
        footerBgColor = Color.FromArgb(97, 203, 243);
        FormatCells(sheet.Cells[$"A{allocationRow - 2}:B{allocationRow}"], bgColor: footerBgColor);
        // 在第 1 行插入 10 行空白行
        sheet.InsertRow(1, 10);
        sheet.Cells[1, 1].Value = "報價單 (供內部使用)";
        sheet.Cells[1, 1].Style.Font.Size = 24;
        MergeAndAlign(1, 1, 1, header.Length, isCenter: true);
        AlignColumnWidth();
    }
}
