using OfficeOpenXml;
using static OfficeOpenXml.ExcelErrorValue;

public class ProjectCostWorksheet : WorksheetBase
{
    public override string[] header => includeArchitectColumn
        ? new string[] { "專案階段", "工作編號", "工作項目", "工作說明", "工作天數(小計)", "專案經理", "", "架構師", "", "部署者", "", "開發者", "", "內部成本小計", "備註" }
        : new string[] { "專案階段", "工作編號", "工作項目", "工作說明", "工作天數(小計)", "專案經理", "", "部署者", "", "開發者", "", "內部成本小計", "備註" };

    public override int[] widthAlignment => includeArchitectColumn
        ? new int[] { 9, 14, 30, 45, 15, 8, 11, 8, 11, 8, 11, 8, 11, 15, 10 }
        : new int[] { 9, 14, 30, 45, 15, 8, 11, 8, 11, 8, 11, 15, 10 };

    public override int lastRow { get; set; }
    private bool includeArchitectColumn;

    public ProjectCostWorksheet(ExcelWorksheet sheet, bool withArchitectColumn) : base(sheet)
    {
        includeArchitectColumn = withArchitectColumn;
        projCostRefSheet = sheet.Name;
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

    // 寫入標題行
    public override void WriteAndFormatHeader(int startRow, int endRow)
    {
        string[] projcostheader = includeArchitectColumn
            ? new string[] { "", "", "", "", "", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "", "", "" }
            : new string[] { "", "", "", "", "", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "", "", "" };

        for (int col = 0; col < header.Length; col++)
        {
            sheet.Cells[lastRow, col + 1].Value = projcostheader[col];
        }
        // 每次寫入後遞增行
        lastRow++;

        // 動態合併欄位
        MergeAndAlign(startRow, Array.IndexOf(header, "專案經理") + 1, startRow, Array.IndexOf(header, "專案經理") + 2);
        if (includeArchitectColumn)
        {
            MergeAndAlign(startRow, Array.IndexOf(header, "架構師") + 1, startRow, Array.IndexOf(header, "架構師") + 2);
        }
        MergeAndAlign(startRow, Array.IndexOf(header, "部署者") + 1, startRow, Array.IndexOf(header, "部署者") + 2);
        MergeAndAlign(startRow, Array.IndexOf(header, "開發者") + 1, startRow, Array.IndexOf(header, "開發者") + 2);
        MergeAndAlign(startRow, Array.IndexOf(header, "專案階段") + 1, startRow + 1, Array.IndexOf(header, "專案階段") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作編號") + 1, startRow + 1, Array.IndexOf(header, "工作編號") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作項目") + 1, startRow + 1, Array.IndexOf(header, "工作項目") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作說明") + 1, startRow + 1, Array.IndexOf(header, "工作說明") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作天數(小計)") + 1, startRow + 1, Array.IndexOf(header, "工作天數(小計)") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "內部成本小計") + 1, startRow + 1, Array.IndexOf(header, "內部成本小計") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "備註") + 1, startRow + 1, Array.IndexOf(header, "備註") + 1);
        FormatCells(sheet.Cells[$"{startCol}{startRow}:{endCol}{endRow}"], isHeader: true, isHair: true, isBorder: true);
    }

    public override void FormatPhase()
    {
        // 格式化內容
        MergeAndAlign(phaseStartRow + lastRow, 1, phaseEndRow + lastRow - 1, 1);
        referProjCostRow = phaseTitleRow + lastRow;
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseEndRow + lastRow - 1}"], isHair: true, isBorder: true);
        // 使用 RGB-淺綠色-自定義顏色格式化階段標題
        Color titleBgColor = Color.FromArgb(198, 224, 180);  
        MergeAndAlign(phaseTitleRow + lastRow, 1, phaseTitleRow + lastRow, header.Length);
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseTitleRow + lastRow}"], bgColor: titleBgColor, isTitle: true);
    }

    public void WriteCostSumFormula(int column)
    {
        // 在寫入資料時根據 includeArchitectColumn 動態處理。
        if (includeArchitectColumn)
        {
            // 加總 FG 專案經理, HI 架構師, JK 部署者, LM 開發者	
            sheet.Cells[currentRow, column].Formula = $"SUM(F{currentRow}*G{currentRow},H{currentRow}*I{currentRow},J{currentRow}*K{currentRow},L{currentRow}*M{currentRow})";
        }
        else
        {
            // 加總 FG 專案經理, HI 部署者, JK 開發者	
            sheet.Cells[currentRow, column].Formula = $"SUM(F{currentRow}*G{currentRow},H{currentRow}*I{currentRow},J{currentRow}*K{currentRow})";
        }
        sheet.Cells[currentRow, column].Style.Numberformat.Format = "$#,##0";
    }

    public override void WriteFooter()
    {
        // 動態計算欄位索引
        int mandayTotal = 1; // 總計
        int mandaySubtotal = 5; // 工作天數(小計)
        int projectManagerColumn = 6; // F 欄
        int architectColumn = includeArchitectColumn ? 8 : -1; // H 欄
        int deployerColumn = includeArchitectColumn ? 10 : 8; // J 或 H 欄
        int developerColumn = includeArchitectColumn ? 12 : 10; // L 或 J 欄
        int totalCostColumn = includeArchitectColumn ? 14 : 12; // N 或 L 欄

        sheet.Cells[currentRow, mandayTotal].Value = $"總計";
        sheet.Cells[currentRow, mandaySubtotal].Formula = $"SUM({(char)('A' + mandaySubtotal - 1)}3:{(char)('A' + mandaySubtotal - 1)}{currentRow - 1})";
        CenterText(mandaySubtotal);
        sheet.Cells[currentRow, projectManagerColumn].Formula = $"SUM(${(char)('A' + projectManagerColumn - 1)}3:{(char)('A' + projectManagerColumn - 1)}{currentRow - 1})";
        CenterText(projectManagerColumn);
        if (includeArchitectColumn)
        {
            sheet.Cells[currentRow, architectColumn].Formula = $"SUM(${(char)('A' + architectColumn - 1)}3:${(char)('A' + architectColumn - 1)}{currentRow - 1})";
            CenterText(architectColumn);
        }
        sheet.Cells[currentRow, deployerColumn].Formula = $"SUM(${(char)('A' + deployerColumn - 1)}3:${(char)('A' + deployerColumn - 1)}{currentRow - 1})";
        CenterText(deployerColumn);
        sheet.Cells[currentRow, developerColumn].Formula = $"SUM(${(char)('A' + developerColumn - 1)}3:${(char)('A' + developerColumn - 1)}{currentRow - 1})";
        CenterText(developerColumn);
        sheet.Cells[currentRow, totalCostColumn].Formula = $"SUM(${(char)('A' + totalCostColumn - 1)}3:${(char)('A' + totalCostColumn - 1)}{currentRow - 1})";
        sheet.Cells[currentRow, totalCostColumn].Style.Numberformat.Format = "$#,##0";
        CenterText(totalCostColumn, isRight:true);

        // 淺橘紅色 248, 203, 173
        Color footerBgColor = Color.FromArgb(248, 203, 173);    
        FormatCells(sheet.Cells[$"{startCol}{currentRow}:{endCol}{currentRow}"], bgColor: footerBgColor, isBorder: true);
        referProjCostRow = phaseTitleRow + lastRow;

        // 深黑色 34,43,53      
        sheet.Cells[currentRow + 1, 1].Value = "專案成員人天分配：";
        MergeAndAlign(currentRow + 1, 1, currentRow + 1, 3);
        FormatCells(sheet.Cells[$"A{currentRow + 1}:C{currentRow + 4}"], isBorder: true, fontColor:Color.White, bgColor: Color.FromArgb(34, 43, 53));

        sheet.Cells[currentRow + 1, totalCostColumn].Value = "建議報價";
        MergeAndAlign(currentRow + 1, totalCostColumn, currentRow + 1, totalCostColumn + 1, isCenter: true);
        FormatCells(sheet.Cells[currentRow + 1, totalCostColumn, currentRow + 1, totalCostColumn + 1], isBorder: true, fontColor: Color.White, bgColor: Color.FromArgb(34, 43, 53));

        // 取得今天的日期，並格式化為 "西元年/月/日" 格式
        string getDatetimeToday = DateTime.Now.ToString("yyyy/MM/dd");

        // 淺綠色 198, 224, 180
        footerBgColor = Color.FromArgb(198, 224, 180);
        
        // 動態設置每個成員的數據
        int allocationRow = currentRow + 2;
        // 專案經理
        sheet.Cells[allocationRow, 1].Value = "A";
        sheet.Cells[allocationRow, 2].Value = "專案經理";
        sheet.Cells[allocationRow, 3].Formula = $"={(char)('A' + projectManagerColumn - 1)}{currentRow}";
        
        // 建議報價 - 加總
        sheet.Cells[allocationRow, totalCostColumn].Formula = $"={(char)('A' + totalCostColumn - 1)}{currentRow}";
        sheet.Cells[allocationRow, totalCostColumn].Style.Font.Bold = true;
        sheet.Cells[allocationRow, totalCostColumn].Style.Numberformat.Format = "$#,##0";
        MergeAndAlign(allocationRow, totalCostColumn, allocationRow, totalCostColumn + 1, isCenter: true);
        FormatCells(sheet.Cells[$"{(char)('A' + totalCostColumn - 1)}{allocationRow}:{(char)('A' + totalCostColumn)}{allocationRow}"], bgColor: footerBgColor, isBorder: true);
        
        // 架構師
        allocationRow++;
        if (includeArchitectColumn)
        {
            sheet.Cells[allocationRow, 1].Value = "B";
            sheet.Cells[allocationRow, 2].Value = "架構師";
            sheet.Cells[allocationRow, 3].Formula = $"={(char)('A' + architectColumn - 1)}{currentRow}";
            allocationRow++;
        }
        
        // 製表日期
        sheet.Cells[currentRow + 3, totalCostColumn].Value = $"製表日期：{getDatetimeToday}";
        FormatCells(sheet.Cells[$"{(char)('A' + totalCostColumn - 1)}{currentRow + 3}:{(char)('A' + totalCostColumn)}{currentRow + 3}"], isBorder: true);
        
        // 部署者
        sheet.Cells[allocationRow, 1].Value = "C";
        sheet.Cells[allocationRow, 2].Value = "部署者";
        sheet.Cells[allocationRow, 3].Formula = $"={(char)('A' + deployerColumn - 1)}{currentRow}";
        allocationRow++;
        // 開發者
        sheet.Cells[allocationRow, 1].Value = "D";
        sheet.Cells[allocationRow, 2].Value = "開發者";
        sheet.Cells[allocationRow, 3].Formula = $"={(char)('A' + developerColumn - 1)}{currentRow}";
        FormatCells(sheet.Cells[$"A{currentRow + 2}:C{allocationRow}"], bgColor: footerBgColor, isBorder: true, isThin: true);
        FormatCells(sheet.Cells[$"B{currentRow + 2}:B{allocationRow}"], bgColor: footerBgColor, isBorder: true);
        sheet.Cells[allocationRow + 2, 1].Value = "註1.本專案成本表所列的總工天為專員成員的工作總天數，非專案建置所需日曆天";
        // 在第 1 行插入 1 行空白行
        sheet.InsertRow(1, 1);
        sheet.Cells[1, 1].Style.Font.Size = 14;
        MergeAndAlign(1, 1, 1, header.Length, isCenter: true);
        AlignColumnWidth();
    }
}
