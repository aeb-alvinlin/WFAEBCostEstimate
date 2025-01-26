using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
// nuget package "EPPlus" Version 7.5.3
using OfficeOpenXml; 
using OfficeOpenXml.Style;
using static System.Runtime.InteropServices.JavaScript.JSType;
// Date: 2025/1/25
// Time: 23:43 pm
// Version: 0.1c3 (整合有和沒有架構師版本)
// Author: alvin.lin@aceraeb.com
// License: MIT
/*
# EPPlus 7
## Announcement: new license model from version 5
EPPlus has from this new major version changed license from LGPL to [Polyform Noncommercial 1.0.0](https://polyformproject.org/licenses/noncommercial/1.0.0/).
With the new license EPPlus is still free to use in some cases, but will require a commercial license to be used in a commercial business.
This is explained in more detail [here](https://www.epplussoftware.com/Home/LgplToPolyform).
Commercial licenses, which includes support, can be purchased at (https://www.epplussoftware.com/).
The source code of EPPlus has moved to a [new github repository](https://github.com/EPPlusSoftware/EPPlus)
*/

namespace WFormProjEstimateApp1
{
    public partial class WFormProjEstimate : Form
    {
        string taskListSheetName = "工作項目清單";
        string projCostSheetName = "專案成本表";
        string quotationSheetName = "報價單(供內部使用)";
        string deliverablesSheetName = "專案文件交付清單";

        private string? workItemsFilePath = null;
        private string? targetFilePath = null;
        private static bool withArchiectColumn = false;
        int mandaySubtotal = 5; // 工作天數(小計)
        int projectManagerColumn = 6; // F 欄
        int architectColumn = default; // H 欄
        int deployerColumn = default; // J 或 H 欄
        int developerColumn = default; // L 或 J 欄
        int totalCostColumn = default; // N 或 L 欄
        private SourceWorkItems? workItems;

        public WFormProjEstimate(string? sourceExcelFilePath)
        {
            // Use EPPlus in a noncommercial context according to the Polyform Noncommercial license  
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            if (sourceExcelFilePath != null)
            {
                workItemsFilePath = sourceExcelFilePath;
                GenerateQuotationReport();
                MessageBox.Show($"已讀取「{workItemsFilePath}」並產生報價試算表「{targetFilePath}」！");
                return;
            }

            InitializeComponent();
        }

        private void WFormProjEstimate_Load(object sender, EventArgs e)
        {
            string[] deliverableItems = [
                "A1 產品簡報",
                "A2 專案建議書",
                "A3 專案成本表",
                "A4 工作項目(Action Item)",
                "A5 工作說明書(SOW)",
                "B1 系統環境調查表",
                "B2 啟動會議簡報",
                "C1 架構流程圖",
                "C2 系統規畫建議書",
                "C3 整理程序說明書",
                "C4 工作分解結構(WBS)",
                "D1 功能驗證報告書",
                "E1 問題處理清單",
                "E2 管理手冊",
                "E3 操作手冊",
                "E4 教育訓練手冊",
                "F1 專案結案報告書",
                "F2 結案會議簡報",
                "G1 工作紀錄/會議紀錄",
                "G2 進度報告",
                "G3 週報",
                "G4 其他(郵件、截圖)",
                "G5 合約",
            ];
            deliverableSelectionComboBox.Items.AddRange(deliverableItems);
        }

        private void CreateworkItemsFile(bool isWithArchiectColumn)
        {
            // 產生目標工作表 Excel 檔案名稱與路徑
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string messageArchiect = isWithArchiectColumn ? "(含)架構師" : "(無)架構師";
            string CustomerName = string.Empty;
            string ProjectName = string.Empty;
            do
            {
                // 使用 DateTime 取得當前日期和時間
                string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");

                // 目標檔案客戶名稱欄位
                CustomerName = customerNameTextBox.Text;
                if (string.IsNullOrEmpty(CustomerName))
                {
                    MessageBox.Show("請在客戶名稱欄位先輸入資料後再試一次！", "客戶名稱欄位沒有任何資料！");
                    return;
                }

                // 目標檔案專案名稱欄位
                ProjectName = projectNameTextBox.Text;
                if (string.IsNullOrEmpty(ProjectName))
                {
                    MessageBox.Show("請在專案名稱欄位先輸入資料後再試一次！", "專案名稱欄位沒有任何資料！");
                    return;
                }

                // 目標檔案名稱：工作清單(範本).xlsx - 產生新的空白的 Excel 檔案做為工作清單範本
                string workItemFileName = @$"[{CustomerName}-{ProjectName}]_工作清單{messageArchiect}範本{currentDate}.xlsx";

                // 使用取得目標檔案名稱和路徑
                targetFilePath = Path.Combine(desktopPath, workItemFileName);
                // 檢查直到確認此檔案並沒有相同名稱的檔名存在
            } while (File.Exists(targetFilePath));

            // 呼叫 workItemsContext 建立來源工作表
            workItems = new(isWithArchiectColumn);
            workItems.CreateSourceWorkItemsFile(targetFilePath!, CustomerName, ProjectName);

            string message = $"範本檔案已成功儲存到 [{targetFilePath}]！";

            // 顯示 MessageBox 並有 Yes / No 兩個選項
            DialogResult ifOpenExcelDirectly = MessageBox.Show(message + "\n\n是否同時開啟工作項目範本 Excel 檔案？", $"範本「{messageArchiect}」已成功儲存！", MessageBoxButtons.YesNo);

            // 是否要用 Excel 直接開啟檔案
            if (ifOpenExcelDirectly == DialogResult.Yes)
            {
                try
                {
                    // 建立 processExcel 
                    var processExcel = new Process();

                    // 開啟檔案
                    processExcel.StartInfo = new ProcessStartInfo(@$"{targetFilePath}"!)
                    {
                        UseShellExecute = true
                    };
                    processExcel.Start();
                }
                catch (Exception ex)
                {
                    // 如果開啟檔案失敗
                    MessageBox.Show($"錯誤!{ex.Message}", "無法開啟檔案！");
                    MessageBox.Show("無法開啟檔案！");
                }
            }
        }

        private void CreateSourceWorkItemA_Click(object sender, EventArgs e)
        {
            withArchiectColumn = false;
            CreateworkItemsFile(withArchiectColumn);
        }

        private void CreateSourceWorkItemB_Click(object sender, EventArgs e)
        {
            withArchiectColumn = true;
            CreateworkItemsFile(withArchiectColumn);
        }

        // 開啟來源工作清單 Excel 檔案
        private void OpenWorkItemsSource_Click(object sender, EventArgs e)
        {
            withArchiectColumn = false;
            OpenworkItemsFile(withArchiectColumn);
        }

        private void OpenWorkItemsWithArchiectSource_Click(object sender, EventArgs e)
        {
            withArchiectColumn = true;
            OpenworkItemsFile(withArchiectColumn);
        }

        private void OpenworkItemsFile(bool isWithArchiectColumn)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workItemsFilePath = openFileDialog.FileName;
                    StatusBarLabel.Text = $"已讀入檔案: {workItemsFilePath}";
                }
            }
            if (string.IsNullOrEmpty(workItemsFilePath))
            {
                MessageBox.Show("請先選擇來源檔案！", "沒有選取任何檔案！");
                return;
            }
            // 呼叫 workItemsContext 處理來源工作表
            workItems = new(isWithArchiectColumn);
            workItems.ReadSourceWorkItemsFile(workItemsFilePath);
            if (!string.IsNullOrEmpty(workItems.errorMessage))
            {
                MessageBox.Show(workItems.errorMessage, "請先排除下列錯誤後再重新讀取檔案！");
                string[] errorMessageParts = workItems.errorMessage.Split('\n');
                workItemsFilePath = string.Empty;
                if (errorMessageParts.Length > 2)
                {
                    if (!string.IsNullOrEmpty(errorMessageParts[^2]))
                    {
                        StatusBarLabel.Text = errorMessageParts[^2];
                    }
                    return;
                }
                StatusBarLabel.Text = "請先排除錯誤後再重新讀取檔案！";
                return;
            }
            customerNameTextBox.Text = workItems.ProjectContext!.CustomerName;
            projectNameTextBox.Text = workItems.ProjectContext!.ProjectName;
            salesDepartmentComboBox.Text = workItems.ProjectContext!.SalesDepartmentName;
            salesRepresentativeComboBox.Text = workItems.ProjectContext!.SalesRepresentativeteName;
            salesRepresentativeEmailAddress.Text = workItems.ProjectContext!.SalesEmailAddress;
            salesRepresentativePhoneExtension.Text = workItems.ProjectContext!.SalesPhoneExtension;
            techDepartmentComboBox.Text = workItems.ProjectContext!.TechDepartmentName;
            techRepresentativeComboBox.Text = workItems.ProjectContext!.TechRepresentativeteName;
            techRepresentativeEmailAddress.Text = workItems.ProjectContext!.TechEmailAddress;
            techRepresentativePhoneExtension.Text = workItems.ProjectContext!.TechPhoneExtension;
        }
        // 產生目標工作報表 Excel 檔案

        private void SaveQuotationReportTarget_Click(object sender, EventArgs e)
        {
            if (workItemsFilePath == null)
            {
                MessageBox.Show("請先透過選單的「讀取來源檔案」選取工作清單試算表！");
                return;
            }
            architectColumn = withArchiectColumn ? 8 : -1; // H 欄
            deployerColumn = withArchiectColumn ? 10 : 8; // J 或 H 欄
            developerColumn = withArchiectColumn ? 12 : 10; // L 或 J 欄
            totalCostColumn = withArchiectColumn ? 14 : 12; // N 或 L 欄

            GenerateQuotationReport();

            // 組合訊息
            string message = $"檔案「{targetFilePath}」已儲存！";

            // 顯示 MessageBox 並有 Yes / No 兩個選項
            DialogResult ifOpenExcelDirectly = MessageBox.Show(message + "\n\n是否同時開啟工作項目成本估算報價單 Excel 檔案？", $"已讀取「{Path.GetFileName(workItemsFilePath)}」並產生工作項目成本估算報價單！", MessageBoxButtons.YesNo);

            // 是否要用 Excel 直接開啟檔案
            if (ifOpenExcelDirectly == DialogResult.Yes)
            {
                try
                {
                    // 建立 processExcel 
                    var processExcel = new Process();

                    // 開啟檔案
                    processExcel.StartInfo = new ProcessStartInfo(@$"{targetFilePath}"!)
                    {
                        UseShellExecute = true
                    };
                    processExcel.Start();
                }
                catch (Exception ex)
                {
                    // 如果開啟檔案失敗
                    MessageBox.Show($"錯誤!{ex.Message}", "無法開啟檔案！");
                    MessageBox.Show("無法開啟檔案！");
                }
            }

            Application.Exit();
        }

        // 儲存目標工作報表 Excel 檔案
        private void GenerateQuotationReport()
        {
            // 產生目標工作表 Excel 檔案名稱與路徑
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            do
            {
                // 使用 DateTime 取得當前日期和時間
                string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");

                // 目標檔案名稱：專案成本表.xlsx - 產生新的空白的Excel檔案做為專案成本表
                string costFileName = @$"[{workItems!.ProjectContext!.CustomerName}-{workItems.ProjectContext!.ProjectName}]{projCostSheetName}_{currentDate}.xlsx";

                // 使用取得目標檔案名稱和路徑
                targetFilePath = Path.Combine(desktopPath, costFileName);
                // 檢查直到確認此檔案並沒有相同名稱的檔名存在
            } while (File.Exists(targetFilePath));

            // 開始主要程式
            using (ExcelPackage costPackage = new ExcelPackage())
            {
                // 啟始類別
                var quotationSheet = costPackage.Workbook.Worksheets.Add(quotationSheetName);
                var projCostSheet = costPackage.Workbook.Worksheets.Add(projCostSheetName);
                var taskListSheet = costPackage.Workbook.Worksheets.Add(taskListSheetName);
                var deliverablesSheet = costPackage.Workbook.Worksheets.Add(deliverablesSheetName);

                // 使用 WorksheetBase 的子類別來協助寫入各工作表的資料 將各工作表的物件放入 using 區塊，確保資源釋放
                using (QuotationWorksheet quotation = new QuotationWorksheet(quotationSheet, withArchiectColumn))
                using (ProjectCostWorksheet projectCost = new ProjectCostWorksheet(projCostSheet, withArchiectColumn))
                using (TaskListWorksheet taskList = new TaskListWorksheet(taskListSheet))
                using (DeliverablesWorkSheet deliverables = new DeliverablesWorkSheet(deliverablesSheet))
                {
                    // 定義 "階段編號" 從 1 開始計算
                    int phaseNumber = 1;

                    // 定義 "序號" 從 1 開始計算
                    int sequenceNumber = 1;

                    // 階段的順序切片
                    string lastPhasePart = null!;

                    // 寫入工作表標題
                    taskList.WriteHeader();
                    projectCost.WriteHeader();
                    quotation.WriteHeader();
                    deliverables.WriteHeader();

                    // 讀取的工作項目將工作簿裡的資料用各工作表各自的方法寫入新工作簿
                    foreach (var phase in workItems.phaseData)
                    {
                        // phaseName 為 Key 也就是階段名稱
                        string phaseName = phase.Key;

                        // 階段名稱用 "-" 來分割字串
                        string[] phaseNameParts = phaseName.Split('-');

                        // 擷取階段名稱分割後階段的第一部分順序切片出來，並使用 Trim() 移除首尾空白字元
                        string phasePart = phaseNameParts[0].Trim();

                        // taskLists 為 Value 也就是該階段以 TaskItemRow 型別儲存的資料清單
                        List<SourceWorkItemRow> taskLists = phase.Value;

                        // 比對目前的 階段的順序切片 是否就是前次的 階段的順序切片 相同的名稱
                        if (phasePart != lastPhasePart)
                        // 如果目前的 階段的順序切片 是新的階段 將以階段的開始做為寫入該階段的資料的形式
                        {
                            // 先以 SetPhaseTitle 設定 PhaseTitle 階段標題目前的位置
                            WorksheetBase.SetPhaseTitle();

                            // 寫入 "階段名稱"
                            taskList.WriteText(phaseName, 1);
                            projectCost.WriteText(phaseName, 1);
                            quotation.WriteText(phaseName, 1);

                            // 全部一起移到下一列
                            WorksheetBase.MoveSharedRowToNext();
                        }
                        // 定義 "階段編號" 後的 "點"->"大綱編號" 從 1 開始計算
                        int outlineNumber = 1;

                        // 先以 SetPhaseStart 設定 PhaseStart 階段項目目前的位置
                        WorksheetBase.SetPhaseStart();

                        // 開始從 taskLists 清單內逐一取出工作項目資料寫入 "階段項目" 
                        foreach (var item in taskLists)
                        {
                            // 如果目前 大綱編號 為 1 表示這是階段開頭的第一個編號
                            if (item.TaskName != "無")
                            {
                                if (outlineNumber == 1)
                                {
                                    // 只有在階段開頭的第一個編號時才寫入階段編號
                                    taskList.WriteText(phaseNumber, 1);
                                    projectCost.WriteText(phaseNumber, 1);
                                }

                                // 寫入大綱編號
                                taskList.WriteText($"{phaseNumber}.{outlineNumber}", 2);

                                // 寫入 序號
                                quotation.WriteValue(sequenceNumber, 1);
                                projectCost.WriteValue(sequenceNumber, 2);

                                // 寫入 工作天數
                                taskList.WriteText(item.TotalTaskDays, mandaySubtotal, isRight: false); ;
                                projectCost.WriteText(item.TotalTaskDays, mandaySubtotal, isRight: false); ;

                                // 寫入 專案經理
                                projectCost.WriteValue(item.PrjManagerDays, projectManagerColumn, isRight: false); ;
                                projectCost.WriteNumeric(8000, projectManagerColumn + 1);

                                // 寫入 架構師
                                if (withArchiectColumn)
                                {
                                    projectCost.WriteValue(item.ArchiectDays, architectColumn, isRight: false); ;
                                    projectCost.WriteNumeric(8000, architectColumn + 1);
                                }

                                // 寫入 部署者
                                projectCost.WriteValue(item.DeployerDays, deployerColumn, isRight: false); ;
                                projectCost.WriteNumeric(8000, deployerColumn + 1);

                                // 寫入 負責單位預設值為 "AEB"
                                taskList.WriteText("AEB", 9);

                                // 寫入 開發者
                                projectCost.WriteValue(item.DeveloperDays, developerColumn);
                                projectCost.WriteNumeric(8000, developerColumn + 1);
                                projectCost.WriteCostSumFormula(totalCostColumn);
                            }

                            // 寫入 工作項目
                            taskList.WriteText(item.TaskName, 3, isCenter: false);
                            projectCost.WriteText(item.TaskName, 3, isCenter: false);
                            quotation.WriteText(item.TaskName, 2, isCenter: false);

                            // 寫入 工作說明
                            taskList.WriteText(item.TaskDescription, 4, isCenter: false);
                            projectCost.WriteText(item.TaskDescription, 4, isCenter: false);
                            quotation.WriteText(item.TaskDescription, 3, isCenter: false);

                            // 移到下一列
                            WorksheetBase.MoveSharedRowToNext();

                            // 大綱編號加 1
                            outlineNumber++;

                            // 序號加 1
                            sequenceNumber++;
                        }
                        // 寫入 工作項目
                        WorksheetBase.SetPhaseEnd();

                        // 合併 工作項目
                        taskList.MergeText(3);
                        projectCost.MergeText(3, sheetCalculate: true);
                        quotation.MergeText(2, sheetCalculate: true);

                        // 使用自定義顏色格式化階段 
                        taskList.FormatPhase();

                        // 使用自定義顏色格式化階段 
                        projectCost.FormatPhase();

                        // 使用自定義顏色格式化階段 
                        quotation.FormatPhase();
                        if (phasePart != lastPhasePart)
                        {
                            quotation.PhaseSumPrice(workItems.phaseCount[phasePart]);
                        }

                        // 階段編號加 1
                        phaseNumber++;

                        // 整個階段完成後，將目前的 phasePart, 也就是階段名稱的階段順序切片部分 指派給 lastPhasePart。用來在下階段判斷是否還是相同階段順序
                        lastPhasePart = phasePart;
                    }
                    // 最後定位在表格的最後一行為 SetPhaseTitle。這是為了讓表尾的文字區段有位置的參考依據
                    WorksheetBase.SetPhaseTitle();

                    // 寫入表尾
                    taskList.WriteFooter();

                    // 寫入表尾
                    projectCost.WriteFooter();

                    // 寫入表尾
                    quotation.WriteFooter();

                    // 寫入表尾
                    deliverables.WriteFooter();

                    // 最後修飾 - 加入專案資訊
                    deliverablesSheet.Cells[2, 1].Value = customerNameTextBox.Text;
                    deliverablesSheet.Cells[2, 2].Value = projectNameTextBox.Text;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // 最後修飾 - 加入專案交付項目
                    if (deliverableListBox.Items.Count > 0)
                    {
                        deliverablesSheet.Cells[4, 1].Value = "交付項目";
                        deliverablesSheet.Cells[4, 1, 4, 2].Merge = true;
                        deliverablesSheet.Cells[4, 1].Style.Font.Bold = true;
                        deliverablesSheet.Cells[4, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        deliverablesSheet.Cells[4, 1].Style.Font.Color.SetColor(Color.White);
                        deliverablesSheet.Cells[4, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 43, 53));

                        int deliverablesRow = 4;
                        foreach (var listItem in deliverableListBox.Items)
                        {
                            deliverablesRow++;
                            deliverablesSheet.Cells[deliverablesRow, 1].Value = listItem;
                            deliverablesSheet.Cells[deliverablesRow, 1, deliverablesRow, 2].Merge = true;
                        }
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        deliverablesSheet.Cells[5, 1, deliverablesRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    }

                    // 最後修飾 - 加入專案資訊
                    taskListSheet.Cells[1, 1].Formula = $"={deliverablesSheetName}!A2&\"-\"&{deliverablesSheetName}!B2&\"-{taskListSheetName}\"";
                    projCostSheet.Cells[1, 1].Formula = $"={deliverablesSheetName}!A2&\"-\"&{deliverablesSheetName}!B2&\"-{projCostSheetName}\"";

                    // 最後修飾 - 加入專案資訊
                    string[] columnNames = ["客戶名稱", "專案名稱", "業務部門", "業務代表", "電子信箱s", "電話分機s", "報價日期", "報價單號", "技術部門", "部門代表", "電子信箱t", "電話分機t"];
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "客戶名稱") % 7), 1 + ((Array.IndexOf(columnNames, "客戶名稱") / 7) * 3)].Value = "    客戶名稱" + "：" + customerNameTextBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "專案名稱") % 7), 1 + ((Array.IndexOf(columnNames, "專案名稱") / 7) * 3)].Value = "    專案名稱" + "：" + projectNameTextBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "業務部門") % 7), 1 + ((Array.IndexOf(columnNames, "業務部門") / 7) * 3)].Value = "    業務部門" + "：" + salesDepartmentComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "業務代表") % 7), 1 + ((Array.IndexOf(columnNames, "業務代表") / 7) * 3)].Value = "    業務代表" + "：" + salesRepresentativeComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "電子信箱s") % 7), 1 + ((Array.IndexOf(columnNames, "電子信箱s") / 7) * 3)].Value = "    電子信箱" + "：" + salesRepresentativeEmailAddress.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "電話分機s") % 7), 1 + ((Array.IndexOf(columnNames, "客戶名稱") / 7) * 3)].Value = "    電話分機" + "：" + salesRepresentativePhoneExtension.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "報價日期") % 7), 1 + ((Array.IndexOf(columnNames, "報價日期") / 7) * 3)].Value = "    報價日期" + "：" + DateTime.Now.ToString("yyyy/MM/dd");
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "報價單號") % 7), 1 + ((Array.IndexOf(columnNames, "報價單號") / 7) * 3)].Value = "報價單號" + "：";
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "技術部門") % 7), 1 + ((Array.IndexOf(columnNames, "技術部門") / 7) * 3)].Value = "技術部門" + "：" + techDepartmentComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "部門代表") % 7), 1 + ((Array.IndexOf(columnNames, "部門代表") / 7) * 3)].Value = "部門代表" + "：" + techRepresentativeComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "電子信箱t") % 7), 1 + ((Array.IndexOf(columnNames, "電子信箱t") / 7) * 3)].Value = "電子信箱" + "：" + techRepresentativeEmailAddress.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "電話分機t") % 7), 1 + ((Array.IndexOf(columnNames, "電話分機t") / 7) * 3)].Value = "電話分機" + "：" + techRepresentativePhoneExtension.Text;

                    // 保存到目標檔案
                    costPackage.SaveAs(new FileInfo(targetFilePath!));
                    StatusBarLabel.Text = $"報價單已儲存: {targetFilePath}";
                    // 結束時將物件設為 null 垃圾回收
                }
            }
        }

        private void AddDeliverableButton_Click(object sender, EventArgs e)
        {
            string selection = deliverableSelectionComboBox.Text;
            // 檢查姓名是否為空字串
            if (selection == "")
            {
                MessageBox.Show("請選擇交付項目再按新增！");
                // 離開此事件處理函式
                return;
            }
            if (deliverableListBox.Items.Contains(selection))
            {
                MessageBox.Show("資料已存在!");
            }
            else
            {
                deliverableListBox.Items.Add(selection);
            }
        }

        private void ModifyDeliverableButton_Click(object sender, EventArgs e)
        {
            // 檢查是否選擇了 ListBox 中的項目
            if (deliverableListBox.SelectedItem == null)
            {
                MessageBox.Show("請選擇要修改的交付項目！");
                return;
            }

            // 獲取選中的項目
            string selectedItem = deliverableListBox.SelectedItem.ToString()!;

            // 提示使用者選擇新項目
            string newSelection = deliverableSelectionComboBox.Text;

            // 檢查新選項是否有效
            if (string.IsNullOrEmpty(newSelection))
            {
                MessageBox.Show("請選擇新的交付項目！");
                return;
            }

            // 檢查新項目是否已經存在於 ListBox 中
            if (deliverableListBox.Items.Contains(newSelection))
            {
                MessageBox.Show("選擇的交付項目已經存在！");
                return;
            }

            // 替換選中的項目
            int selectedIndex = deliverableListBox.SelectedIndex;
            deliverableListBox.Items[selectedIndex] = newSelection;
        }

        private void DeleteDeliverableButton_Click(object sender, EventArgs e)
        {
            // 檢查是否選擇了 ListBox 中的項目
            if (deliverableListBox.SelectedItem == null)
            {
                MessageBox.Show("請選擇要刪除的交付項目！");
                return;
            }

            // 確認刪除
            DialogResult result = MessageBox.Show("確定要刪除選中的交付項目嗎？", "刪除確認", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                // 刪除選中的項目
                deliverableListBox.Items.Remove(deliverableListBox.SelectedItem);
            }
        }

        private void LicensingStatement_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本程式除 EPPlus 7.5.3 外其餘部份為 MIT 授權。部份元件使用 EPPlus 7.5.3 套件開發。由於EPPlus 套件已變更其授權方式由 LGPL 改為 Polyform Noncommercial 1.0.0。在此新的授權許可方式下，EPPlus 在某些情況下仍然可以免費使用，但在商業業務中使用則需要商業許可證。本程式限定 EPPlus 套件免費版本使用範圍，請遵照其授權條款依其適用性來使用本軟體。");
        }

        private void VersionDescription_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本程式使用 C# .net 9 開發。目前版本為工作項目成本估算報價單整合架構師欄位版本。版本號 0.1c.3。");
        }

        private void AuthorInformation_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本程式由 alvin.lin@outlook.com 獨力開發。開發完成日期 2025-1-25 日");
        }
    }
}
