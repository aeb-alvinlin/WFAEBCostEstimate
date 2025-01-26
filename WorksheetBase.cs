using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Diagnostics;
using System;
using static System.Runtime.InteropServices.JavaScript.JSType;

public abstract class WorksheetBase : IDisposable
{
    public ExcelWorksheet sheet;
    private bool disposed = false;
    // 抽象屬性，讓子類必須實作
    public abstract string[] header { get; }
    public abstract int[] widthAlignment { get; }
    public abstract int lastRow { get; set; }
    // currentColumn 是主要的工作表，也就是[工作項目清單]工作表，用來給所有子表的參照欄號，它是靜態屬性，所有子類都能存取這個共享的參照欄號
    protected static int currentColumn = default;
    // sharedRow 是所有工作表共用的位移列號，它是靜態屬性，用來給所有子表的參照目前的位移列號
    protected static int sharedRow = default;
    // referencedRow 是主要的工作表列號，它是靜態屬性，用來給所有子表的參照目前的位移列號
    protected static int referencedRow = default;
    // referencedRow 是主要的工作表列號，它是靜態屬性，用來給所有子表的參照目前的位移列號
    protected static int referProjCostRow = default;
    // taskListRefSheet 是主要的工作表名稱，用來給所有子表的參照的工作表名稱
    protected static string taskListRefSheet = null!;
    // projCostRefSheet 是主要的工作表名稱，用來給所有子表的參照的工作表名稱
    protected static string projCostRefSheet = null!;
    // headerSize 是標題的字型大小
    byte headerSize = 11;
    // titleSize 是標題的字型大小
    byte titleSize = 11;
    // headerSize 是內文的字型大小
    byte contextSize = 11;
    // sharedRow 是所有工作表共用的位移列號，它是靜態屬性，用來給所有子表的參照目前的位移列號
    protected static int phaseTitleRow = default;
    // sharedRow 是所有工作表共用的位移列號，它是靜態屬性，用來給所有子表的參照目前的位移列號
    protected static int phaseStartRow = default;
    // sharedRow 是所有工作表共用的位移列號，它是靜態屬性，用來給所有子表的參照目前的位移列號
    protected static int phaseEndRow = default;
    // currentRow 是主要的工作表，也就是[工作項目清單]工作表，每個工作表的當前行數，加上 sharedRow 作為全域位移列，所有子類都能存取這個共享的參照列號
    public char startCol 
    {
        // 動態計算當前的欄號
        get { return (char)('A'); }
    }

    public char endCol
    {
        // 動態計算當前的欄號
        get { return (char)('A' + header.Length - 1); }
    }

    public int currentRow
    {
        // 動態計算當前的列號
        get { return lastRow + sharedRow; }  
    }

    public WorksheetBase(ExcelWorksheet sheet)
    {
        this.sheet = sheet;
    }
    // 抽象方法：寫入標題行

    public void WriteHeader()
    {
        int startRol = lastRow;
        // 寫入標題行
        for (int col = 0; col < header.Length; col++)
        {
            sheet.Cells[lastRow, col + 1].Value = header[col];
        }
        // 每次寫入後遞增行
        lastRow++;
        // 呼叫格式化標題
        WriteAndFormatHeader(startRol, lastRow);
    }

    public abstract void WriteAndFormatHeader(int startRow, int endRow);

    public static void SetPhaseTitle()
    {
        // 每次寫入後遞增列
        phaseTitleRow = sharedRow;
    }
    public static void SetPhaseStart()
    {
        // 每次寫入後遞增列
        phaseStartRow = sharedRow;
    }

    public static void SetPhaseEnd()
    {
        // 每次寫入後遞增列
        phaseEndRow = sharedRow;
    }

    public abstract void FormatPhase();

    public void WriteValue(double number, int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[currentRow, column].Value = number;
        CenterText(column, isCenter, isRight);
    }

    public virtual void WriteText(string text, int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[currentRow, column].Formula = $"={taskListRefSheet}!{(char)('A' + currentColumn - 1 )}{referencedRow}";
        CenterText(column, isCenter, isRight);
    }

    public virtual void WriteText(double number, int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[currentRow, column].Formula = $"={taskListRefSheet}!{(char)('A' + currentColumn - 1)}{referencedRow}";
        CenterText(column, isCenter, isRight);
    }

    public virtual void CenterText(int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[currentRow, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;            // 內文垂直置中
        if (isRight)
        {
            sheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;    // 內文水平靠右
        }
        else if (isCenter)
        {
            sheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;      // 內文水平置中
        }
        else
        {
            sheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;      // 內文水平靠左
        }
    }

    public void MergeText(int column, bool sheetCalculate = false)
    {
        // 透過 sheet.Calculate() 方法告訴 EPPlus 要計算所有公式
        if (sheetCalculate)
        {
            sheet.Calculate();
        }

        int startRow = phaseStartRow;

        // 使用 phaseEndRow 來決定列數的範圍
        while (startRow <= phaseEndRow)
        {
            int endRow = startRow;

            // Find consecutive rows with the same value in the specified column
            while (endRow + 1 <= phaseEndRow &&
                   sheet.Cells[lastRow + endRow, column].Value != null &&
                   sheet.Cells[lastRow + endRow, column].Value.Equals(sheet.Cells[lastRow + endRow + 1, column].Value))
            {
                endRow++;
            }

            // 合併區域包含從 startRow 到 endRow 的所有連續行
            if (endRow > startRow)
            {
                var range = sheet.Cells[lastRow + startRow, column, lastRow + endRow, column];
                if (!range.Merge) // 只在尚未合併的情況下合併
                {
                    range.Merge = true;
                }
            }

            // 移動到下一段非重複區域
            startRow = endRow + 1;
        }
    }

    public virtual void WriteNumeric(double number, int column, bool isCenter = true, bool isRight = false)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[currentRow, column].Value = number;
        sheet.Cells[currentRow, column].Style.Numberformat.Format = "$#,##0";
        CenterText(column, isCenter, isRight);
    }

    public virtual void WriteFormula(string formula, int column)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[currentRow, column].Formula = formula;
    }

    public abstract void WriteFooter();

    // 讓列數自動遞增
    public virtual void MoveRowToNext()
    {
        // 每次寫入後遞增列
        lastRow++;
    }
    // 全域性增加 sharedRow
    public static void MoveSharedRowToNext()
    {
        // 每次全域列變動時，增加 sharedRow
        sharedRow++; 
    }

    // 格式化儲存格，根據是否是標題、是否有內部格線等進行格式化
    protected void FormatCells(ExcelRange range, byte fontSize = default, Color? bgColor = null, Color? fontColor = null, bool isHeader = false, bool isTitle = false, bool isContext = false, bool isRight = false, bool isHair = false, bool isThin = false, bool isBorder = false)
    {
        // 如果 bgColor 和 fontColor 沒有提供，使用預設值：背景白色，字型黑色
        Color cellBgColor = bgColor ?? Color.White;
        Color cellFontColor = fontColor ?? Color.Black; 
        // 設定框線
        if (isHair)
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
        }
        // 設定框線
        if (isThin)
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }
        // 設定表頭列或標題列
        if (isHeader)
        {
            // 表頭列的特殊處理
            cellBgColor = Color.FromArgb(34, 43, 53);  // 深黑色
            cellFontColor = Color.White;
            range.Style.Font.Size = headerSize; // 標題或階段列字型大小
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // 標題居中
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Font.Bold = true;
        }
        if (isTitle)
        {
            // 標題列的特殊處理
            range.Style.Font.Size = titleSize; // 標題或階段列字型大小
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; // 標題居中
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Font.Bold = true;
        }
        if (isContext)
        {
            // 內文
            range.Style.Font.Size = contextSize; // 內文字型大小
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // 標題靠左
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Font.Bold = false;
        }
        if (isRight)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right; // 內文靠右
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
        if (isBorder)
        {
            range.Style.Border.BorderAround(ExcelBorderStyle.Medium);
        }
        if (fontSize > 0)
        {
            range.Style.Font.Size = fontSize;   // 字型大小
        }
        // 設定字型預設是 "Microsoft JhengHei" 中黑體
        range.Style.Font.Name = "Microsoft JhengHei";

        // 設定背景顏色和字體顏色
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(cellBgColor);
        range.Style.Font.Color.SetColor(cellFontColor);
    }
    
    // 標準方法：通過合併來設定儲存格對齊
    public void MergeAndAlign(int fromRow, int fromCol, int toRow, int toCol, bool isCenter = true)
    {
        var range = sheet.Cells[fromRow, fromCol, toRow, toCol];
        try
        {
            range.Merge = true;
        }
        catch
        {
        }

        if (isCenter)
        {
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        else
        {
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }
    }

    // 標準方法：通過合併來設定儲存格對齊
    public void AlignColumnWidth()
    {
        for (int i = 0; i < widthAlignment.Length; i++)
        {
            sheet.Column(i + 1).Width = widthAlignment[i];
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposed) return;

        if (disposing)
        {
            // 釋放託管資源
            if (sheet != null)
            {
                sheet.Dispose();  // 假設 EPPlus 支援 ExcelWorksheet 的 Dispose，否則不需要這行
                sheet = null!;
            }
        }
        disposed = true;
    }

    ~WorksheetBase()
    {
        Dispose(false);
    }
}