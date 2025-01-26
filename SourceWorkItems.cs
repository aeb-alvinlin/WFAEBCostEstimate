using System.Text.RegularExpressions;
using System.Collections.Generic;
// nuget package "EPPlus" Version 7.5.3
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Collections;
using System.Net.Mail;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.IO.Packaging;
using System;


public class SourceWorkItemRow
{
    public string TaskName { get; set; }            // 工作項目
    public string TaskDescription { get; set; }     // 工作說明
    public double TotalTaskDays { get; set; }       // 工作天數(小計)
    public double PrjManagerDays { get; set; }      // 專案經理
    public double ArchiectDays { get; set; }        // 架構師
    public double DeployerDays { get; set; }        // 部署者
    public double DeveloperDays { get; set; }       // 開發者

    public SourceWorkItemRow(string taskName, string taskDescription, double totalTaskDays, double projmanagerDays, double archiectDays, double deployerDays, double developerDays)
    {
        TaskName = taskName;
        TaskDescription = taskDescription;
        TotalTaskDays = totalTaskDays;
        PrjManagerDays = projmanagerDays;
        ArchiectDays = archiectDays;
        DeployerDays = deployerDays;
        DeveloperDays = developerDays;
    }
}
                                                            
public class ProjectItemRow
{
    public string CustomerName { get; set; }                // 客戶名稱
    public string ProjectName { get; set; }                 // 專案名稱
    public string SalesDepartmentName { get; set; }         // 業務部門
    public string SalesRepresentativeteName { get; set; }   // 業務代表
    public string SalesEmailAddress { get; set; }           // 業務電子信箱
    public string SalesPhoneExtension { get; set; }         // 業務電話分機
    public string TechDepartmentName { get; set; }          // 技術部門
    public string TechRepresentativeteName { get; set; }    // 技術代表
    public string TechEmailAddress { get; set; }            // 技術電子信箱
    public string TechPhoneExtension { get; set; }          // 技術電話分機
    public ProjectItemRow(string customerName, string projectName, string salesDpartment, string salesRepresentativete, string salesEmailAddress, string salesPhoneExtension, string techDpartment, string techRepresentativete, string techEmailAddress, string techPhoneExtension)
    {
        CustomerName = customerName;
        ProjectName = projectName;
        SalesDepartmentName = salesDpartment;
        SalesRepresentativeteName = salesRepresentativete;
        SalesEmailAddress = salesEmailAddress;
        SalesPhoneExtension = salesPhoneExtension;
        TechDepartmentName = techDpartment;
        TechRepresentativeteName = techRepresentativete;
        TechEmailAddress = techEmailAddress;
        TechPhoneExtension = techPhoneExtension;
    }
}

public class SourceWorkItems
{
    private bool archiectColumnFlag;
    public string errorMessage = string.Empty; 
    public Dictionary< string, List<SourceWorkItemRow> > phaseData;
    public Dictionary< string, int > phaseCount;
    // 新增屬性
    public ProjectItemRow? ProjectContext;

    public SourceWorkItems(bool isWithArchiectColumn)
    {
        archiectColumnFlag = isWithArchiectColumn;
        phaseData = new Dictionary <string, List<SourceWorkItemRow> >();
        phaseCount = new Dictionary <string, int >();
    }

    // 定義專案概觀裡的每個欄位名稱
    public string[] projectItemsHeader = {
        "客戶名稱",
        "專案名稱",
        "業務部門",
        "業務代表",
        "電子信箱s",
        "電話分機s",
        "技術部門",
        "部門代表",
        "電子信箱t",
        "電話分機t",
    };

    // 定義每個階段的工作表名稱
    protected string[] phaseNames = {
        "專案概觀 (Project Overview)",
        "第一階段-環境調查 (Envisioning)",
        "第二階段-設計規劃 (Planning)",
        "第三階段-發展階段 (Develop)",
        "第四階段-系統部署 (Deploying)",
        "第五階段-結案階段 (Ending)",
        "維護保固階段 (Maintance)",
    };

    private string[] workItemsHeaderA = {
        "工作項目", 
        "工作說明", 
        "工作天數(小計)", 
        "專案經理", 
        "架構師", 
        "部署者", 
        "開發者"
    };

    private string[] workItemsHeaderB = {
        "工作項目",
        "工作說明",
        "工作天數(小計)",
        "專案經理",
        "部署者",
        "開發者"
    };

    public void ReadSourceWorkItemsFile(string filePath)
    {
        string workItemsfilePath = filePath;
        using (ExcelPackage taskPackage = new ExcelPackage(new FileInfo(workItemsfilePath)))
        {
            // 讀取所有工作表名稱，存入字串陣列，並排序
            var worksheetNames = new List<string>();
            foreach (var sheet in taskPackage.Workbook.Worksheets)
            {
                worksheetNames.Add(sheet.Name); // 加入工作表名稱
            }

            // 建立一個 worksheetNames 的副本來進行驗證
            var verificationList = new List<string>(worksheetNames);

            // 將工作表名稱依照 phaseName 和編號排序 (無編號的在前, 有編號的按順序排列)
            var sortedWorksheetNames = worksheetNames
                .OrderBy(name =>
                {
                    // 匹配括號中的數字
                    var match = Regex.Match(name, @" \((\d+)\)$");
                    // 按照數字大小排序
                    return match.Success ? int.Parse(match.Groups[1].Value) : 0; 
                })
                // 無編號的排在前
                .ThenBy(name => name) 
                .ToList();

            // 讀取每個階段的工作表，並將其資料存入 Dictionary 中
            foreach (string phaseName in phaseNames)
            {
                // 使用正則表達式檢查是否存在與 phaseName 相似的工作表名稱
                string pattern = $"^{Regex.Escape(phaseName)}( \\(\\d+\\))?$";
                var matchingSheetNames = sortedWorksheetNames.Where(name => Regex.IsMatch(name, pattern)).ToList();

                if (matchingSheetNames.Count == 0)
                {
                    errorMessage += $"檔案中沒有名稱為 {phaseName} 的工作表！";
                    break;
                }

                foreach (var matchingSheetName in matchingSheetNames)
                {
                    string[] phaseNameParts = matchingSheetName.Split('-');
                    // 擷取分割後的第一部分並使用 Trim() 移除首尾空白字元
                    string phasePart = phaseNameParts[0].Trim();
                    // 找到的表從 verificationList 移除
                    verificationList.Remove(matchingSheetName);

                    var phaseSheet = taskPackage.Workbook.Worksheets[matchingSheetName];

                    if (matchingSheetName == "專案概觀 (Project Overview)")
                    {
                        var (errorField, projectRows) = ReadProjectOverviewInfoSheet(matchingSheetName, phaseSheet);
                        if (errorField != string.Empty)
                        {
                            errorMessage += $"讀取[專案概觀 (Project Overview)]工作表時發生錯誤，{errorField}！";
                        }
                        else
                        {
                            ProjectContext = projectRows;
                        }
                    }
                    else
                    {
                        var (errorField, phaseRows) = ReadPhaseSheet(matchingSheetName, phaseSheet);
                        // 檢查是否有資料
                        if (errorField != null)
                        {
                            errorMessage += $"{errorField}！\n";
                        }
                        else if (phaseRows.Count == 0)
                        {
                            errorMessage += $"工作表 {matchingSheetName} 中沒有有效資料！";
                        }
                        else
                        {
                            phaseData[matchingSheetName] = phaseRows;
                            // 檢查 phasePart 是否已經存在於 phaseCount 字典中
                            if (phaseCount.ContainsKey(phasePart))
                            {
                                // 如果存在，則累加行數
                                phaseCount[phasePart] += phaseRows.Count;
                            }
                            else
                            {
                                // 如果不存在，則初始化
                                phaseCount[phasePart] = phaseRows.Count;
                            }
                        }
                    }
                }
            }
            // 檢查是否有未被 remove 出的工作表名稱
            if (verificationList.Count > 0)
            {
                errorMessage += $"未能成功讀入所有工作表，以下工作表未被讀取: {string.Join(", ", verificationList)}";
            }
        }
    }
    // 用來讀取來源工作表裡到專案概觀定義裡標題每個欄位的第一列資料的內容
    private (string? errorField, ProjectItemRow rows) ReadProjectOverviewInfoSheet(string phaseName, ExcelWorksheet sheet)
    {
        // 預設錯誤訊息內容是空的
        string errorField = string.Empty;

        // 逐一確認來源工作表裡到專案概觀定義裡標題每個欄位的第一列資料的內容是否不是空的
        for (int col=0; col< projectItemsHeader.Length; col++)
        {
            if (sheet.Cells[$"{(char)('A' + col)}2"].Text == string.Empty)
            {
                errorField = $"{phaseName}工作表中的{projectItemsHeader[col]}欄位是空的！";
                break;
            }
        }

        // 逐一將來源工作表的專案概觀定義裡標題每個欄位的資料的內容放進相對應的變數裡
        string customerName = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "客戶名稱"))}2"].Text; 
        string projectName = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "專案名稱"))}2"].Text; 
        string salesDpartment = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "業務部門"))}2"].Text; 
        string salesRepresentativete = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "業務代表"))}2"].Text;
        string salesEmailAddress = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "電子信箱s"))}2"].Text;
        string salesPhoneExtension = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "電話分機s"))}2"].Text;
        string techDpartment = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "技術部門"))}2"].Text; 
        string techRepresentativete = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "部門代表"))}2"].Text;
        string techEmailAddress = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "電子信箱t"))}2"].Text;
        string techPhoneExtension = sheet.Cells[$"{(char)('A' + Array.IndexOf(projectItemsHeader, "電話分機t"))}2"].Text;
        var rows = new ProjectItemRow(customerName, projectName, salesDpartment, salesRepresentativete, salesEmailAddress, salesPhoneExtension, techDpartment, techRepresentativete, techEmailAddress, techPhoneExtension);

        // 如果沒有錯誤，回傳 null 和有效的資料列
        return (errorField, rows);
    }

    // 用來讀取每個階段的工作表並回傳該工作表中的每一行資料
    private (string? errorField, List<SourceWorkItemRow> rows) ReadPhaseSheet(string phaseName, ExcelWorksheet sheet)
    {
        // rows 為所有列的資料內容
        var rows = new List<SourceWorkItemRow>();

        // totalRows 為總資料列數計數
        int totalRows = sheet.Dimension.End.Row;

        // workItemsHeader 用 archiectColumnFlag 判斷應該讀取哪些資料是否包含架構師欄位
        string[] workItemsHeader = archiectColumnFlag
            ? workItemsHeaderA
            : workItemsHeaderB;
        // totalCols 為總資料欄數計數
        int totalCols = workItemsHeader.Length;

        // headerMap 用來讀取來源資料的第一列做為標題欄位，並與前述的標題建立對應的欄位名稱及對應順序號的字典
        var headerMap = new Dictionary<string, int>();

        // 從標題列第一欄開始讀取到最後一欄
        for (int col = 0; col < totalCols; col++)
        {
            // 建立讀取來源資料標題列欄位名稱及其對應的欄號
            string header = sheet.Cells[1, col+1].Text;
            if (!string.IsNullOrEmpty(header))
            {
                headerMap[header] = col; 
            }
        }

        // 如果讀取的來源資料表內的欄位缺少內建的必要欄位則回傳錯誤
        foreach (var header in workItemsHeader)
        {
            if (!headerMap.ContainsKey(header))
            {
                return ($"在 [{phaseName}] 表中缺少必要欄位: {header}", new List<SourceWorkItemRow>()); 
            }
        }

        // 如果總資料列大於等於 2 列, 也就是除了標題列至少有 1 行資料
        if (totalRows >= 2 && sheet.Cells[totalRows, 1].Value != null)
        {
            // 定義計算總天數
            double sumTaskDays = default;
            // 定義架構師總天數
            double archiectDays = default;

            // 從第 2 列, 也就是第 1 筆資料開始讀取（假設第 1 列是標題）讀取每一列對應到標題欄位名稱的資料
            for (int row = 2; row <= totalRows; row++)  // 行從第 2 行開始
            {
                // 回傳錯誤欄位和空的資料列
                string taskName = sheet.Cells[row, headerMap["工作項目"] + 1].Text;
                if (string.IsNullOrEmpty(taskName))
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [工作項目] 不是數值", new List<SourceWorkItemRow>());  
                }

                // 回傳錯誤欄位和空的資料列
                string taskDescription = sheet.Cells[row, headerMap["工作說明"] + 1].Text;
                if (string.IsNullOrEmpty(taskDescription))
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [工作說明] 不是數值", new List<SourceWorkItemRow>());  
                }

                // 回傳錯誤欄位和空的資料列
                bool totalTask = double.TryParse(sheet.Cells[row, headerMap["工作天數(小計)"] + 1].Text, out double totalTaskDays);
                if (!totalTask)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [工作天數(小計)] 不是數值", new List<SourceWorkItemRow>());  
                }

                // 回傳錯誤欄位和空的資料列
                bool projmanager = double.TryParse(sheet.Cells[row, headerMap["專案經理"] + 1].Text, out double projmanagerDays);
                if (!projmanager)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [專案經理] 不是數值", new List<SourceWorkItemRow>());  
                }

                // 回傳錯誤欄位和空的資料列
                bool deployer = double.TryParse(sheet.Cells[row, headerMap["部署者"] + 1].Text, out double deployerDays);
                if (!deployer)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [部署者] 不是數值", new List<SourceWorkItemRow>());
                }

                // 回傳錯誤欄位和空的資料列
                bool developer = double.TryParse(sheet.Cells[row, headerMap["開發者"] + 1].Text, out double developerDays);
                if (!developer)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [開發者] 不是數值", new List<SourceWorkItemRow>());
                }

                // 回傳錯誤欄位和空的資料列
                if (archiectColumnFlag)
                {
                    bool archiect = double.TryParse(sheet.Cells[row, headerMap["架構師"] + 1].Text, out archiectDays);
                    if (!archiect)
                    {
                        return ($"在 [{phaseName}] 表中的第 {row} 列的 [架構師] 不是數值", new List<SourceWorkItemRow>());
                    }
                }

                // 總天數加總
                sumTaskDays = projmanagerDays + archiectDays + deployerDays + developerDays;

                // 檢查總天數是否一致
                if (sumTaskDays != totalTaskDays)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [總工作天數] 不一致", new List<SourceWorkItemRow>());  
                }

                // 如果所有資料都正確，則加入到 rows 中
                rows.Add(new SourceWorkItemRow(taskName, taskDescription, totalTaskDays, projmanagerDays, archiectDays, deployerDays, developerDays));
            }
        }
        // 如果沒有錯誤，回傳 null 和有效的資料列
        return (null, rows);  
    }

    public void CreateSourceWorkItemsFile(string filePath, string CustomerName, string ProjectName)
    {
        string workItemsfilePath = filePath;
        string[] workItemsHeader = archiectColumnFlag
            ? workItemsHeaderA
            : workItemsHeaderB;

        using (ExcelPackage taskPackage = new ExcelPackage(new FileInfo(workItemsfilePath)))
        {
            // 為每個階段建立工作表
            foreach (var phaseName in phaseNames)
            {
                // 新增工作表
                var worksheet = taskPackage.Workbook.Worksheets.Add(phaseName);
                
                if (phaseName == "專案概觀 (Project Overview)")
                {
                    // 填入標題列
                    for (int i = 0; i < projectItemsHeader.Length; i++)
                    {
                        // 加粗標題列
                        worksheet.Cells[1, i + 1].Value = projectItemsHeader[i].Replace("s", "").Replace("t", "");
                        worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.White);
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                        worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    }

                    int additionRow = 2;
                    worksheet.Cells[additionRow, 1].Value = CustomerName;
                    worksheet.Cells[additionRow, 2].Value = ProjectName;
                    worksheet.Cells[additionRow, 3].Value = "業務部門";
                    worksheet.Cells[additionRow, 4].Value = "業務代表";
                    worksheet.Cells[additionRow, 5].Value = "電子信箱@aceraeb.com";
                    worksheet.Cells[additionRow, 6].Value = 1234;
                    worksheet.Cells[additionRow, 7].Value = "IE0T00";
                    worksheet.Cells[additionRow, 8].Value = "孫秋芳 Belris Sun";
                    worksheet.Cells[additionRow, 9].Value = "Berlis.Sun@aceraeb.com";
                    worksheet.Cells[additionRow, 10].Value = 5224;

                    
                    int expandLength = projectItemsHeader.Length;
                    // 加粗標題列
                    additionRow += 3;
                    worksheet.Cells[additionRow, 1].Value = "客戶需求描述";
                    worksheet.Cells[additionRow, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[additionRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Merge = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    worksheet.Cells[additionRow + 1, 1, additionRow + 1, expandLength].Merge = true;

                    // 加粗標題列
                    additionRow += 3;
                    worksheet.Cells[additionRow, 1].Value = "專案架構圖示意圖繪製";
                    worksheet.Cells[additionRow, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[additionRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Merge = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    worksheet.Cells[additionRow + 1, 1, additionRow + 1, expandLength].Merge = true;

                    // 加粗標題列
                    additionRow += 3;
                    worksheet.Cells[additionRow, 1].Value = "定價計算機 Azure 服務估算";
                    worksheet.Cells[additionRow, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[additionRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Merge = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    worksheet.Cells[additionRow + 1, 1, additionRow + 1, expandLength].Merge = true;

                    // 加粗標題列
                    additionRow += 3;
                    worksheet.Cells[additionRow, 1].Value = "蒐集客戶環境調整表";
                    worksheet.Cells[additionRow, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[additionRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Merge = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    worksheet.Cells[additionRow + 1, 1, additionRow + 1, expandLength].Merge = true;

                    // 加粗標題列
                    additionRow += 3;
                    worksheet.Cells[additionRow, 1].Value = "其他補充說明";
                    worksheet.Cells[additionRow, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[additionRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Merge = true;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[additionRow, 1, additionRow, expandLength].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    worksheet.Cells[additionRow + 1, 1].Value = @"你可以視需要複製同一個活頁簿中的工作表。以滑鼠右鍵按下工作表索引標籤，然後選取[移動或複製]。";
                    worksheet.Cells[additionRow + 1, 1, additionRow + 1, expandLength].Merge = true;
                    worksheet.Cells[additionRow + 2, 1].Value = @"選取[建立複本] 核取方塊。在[工作表之前] 底下，選取您要放置複本的位置。選取[確定]。";
                    worksheet.Cells[additionRow + 2, 1, additionRow + 2, expandLength].Merge = true;
                }
                else
                {
                    // 填入標題列
                    for (int i = 0; i < workItemsHeader.Length; i++)
                    {
                        // 加粗標題列
                        worksheet.Cells[1, i + 1].Value = workItemsHeader[i];
                        worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.White);
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                        worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 89, 89));
                    }
                    // 填入範例內容（第二行）
                    worksheet.Cells[2, 1].Value = "無";
                    worksheet.Cells[2, 2].Value = "無";
                    // 範例專案經理人天, (架構師人天), 部署者人天, 開發者人天
                    for (int i = 2; i < workItemsHeader.Length; i++)
                        worksheet.Cells[2, i + 1].Value = 0; // 範例天數
                }

                // 調整欄寬
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            }
            // 保存範本文件
            taskPackage.SaveAs(new FileInfo(workItemsfilePath));
        }
    }
}