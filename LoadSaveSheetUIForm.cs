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
// Version: 0.1c3 (��X���M�S���[�c�v����)
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
        string taskListSheetName = "�u�@���زM��";
        string projCostSheetName = "�M�צ�����";
        string quotationSheetName = "������(�Ѥ����ϥ�)";
        string deliverablesSheetName = "�M�פ���I�M��";

        private string? workItemsFilePath = null;
        private string? targetFilePath = null;
        private static bool withArchiectColumn = false;
        int mandaySubtotal = 5; // �u�@�Ѽ�(�p�p)
        int projectManagerColumn = 6; // F ��
        int architectColumn = default; // H ��
        int deployerColumn = default; // J �� H ��
        int developerColumn = default; // L �� J ��
        int totalCostColumn = default; // N �� L ��
        private SourceWorkItems? workItems;

        public WFormProjEstimate(string? sourceExcelFilePath)
        {
            // Use EPPlus in a noncommercial context according to the Polyform Noncommercial license  
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            if (sourceExcelFilePath != null)
            {
                workItemsFilePath = sourceExcelFilePath;
                GenerateQuotationReport();
                MessageBox.Show($"�wŪ���u{workItemsFilePath}�v�ò��ͳ����պ��u{targetFilePath}�v�I");
                return;
            }

            InitializeComponent();
        }

        private void WFormProjEstimate_Load(object sender, EventArgs e)
        {
            string[] deliverableItems = [
                "A1 ���~²��",
                "A2 �M�׫�ĳ��",
                "A3 �M�צ�����",
                "A4 �u�@����(Action Item)",
                "A5 �u�@������(SOW)",
                "B1 �t�����ҽլd��",
                "B2 �Ұʷ|ĳ²��",
                "C1 �[�c�y�{��",
                "C2 �t�γW�e��ĳ��",
                "C3 ��z�{�ǻ�����",
                "C4 �u�@���ѵ��c(WBS)",
                "D1 �\�����ҳ��i��",
                "E1 ���D�B�z�M��",
                "E2 �޲z��U",
                "E3 �ާ@��U",
                "E4 �Ш|�V�m��U",
                "F1 �M�׵��׳��i��",
                "F2 ���׷|ĳ²��",
                "G1 �u�@����/�|ĳ����",
                "G2 �i�׳��i",
                "G3 �g��",
                "G4 ��L(�l��B�I��)",
                "G5 �X��",
            ];
            deliverableSelectionComboBox.Items.AddRange(deliverableItems);
        }

        private void CreateworkItemsFile(bool isWithArchiectColumn)
        {
            // ���ͥؼФu�@�� Excel �ɮצW�ٻP���|
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string messageArchiect = isWithArchiectColumn ? "(�t)�[�c�v" : "(�L)�[�c�v";
            string CustomerName = string.Empty;
            string ProjectName = string.Empty;
            do
            {
                // �ϥ� DateTime ���o��e����M�ɶ�
                string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");

                // �ؼ��ɮ׫Ȥ�W�����
                CustomerName = customerNameTextBox.Text;
                if (string.IsNullOrEmpty(CustomerName))
                {
                    MessageBox.Show("�Цb�Ȥ�W��������J��ƫ�A�դ@���I", "�Ȥ�W�����S�������ơI");
                    return;
                }

                // �ؼ��ɮױM�צW�����
                ProjectName = projectNameTextBox.Text;
                if (string.IsNullOrEmpty(ProjectName))
                {
                    MessageBox.Show("�Цb�M�צW��������J��ƫ�A�դ@���I", "�M�צW�����S�������ơI");
                    return;
                }

                // �ؼ��ɮצW�١G�u�@�M��(�d��).xlsx - ���ͷs���ťժ� Excel �ɮװ����u�@�M��d��
                string workItemFileName = @$"[{CustomerName}-{ProjectName}]_�u�@�M��{messageArchiect}�d��{currentDate}.xlsx";

                // �ϥΨ��o�ؼ��ɮצW�٩M���|
                targetFilePath = Path.Combine(desktopPath, workItemFileName);
                // �ˬd����T�{���ɮרèS���ۦP�W�٪��ɦW�s�b
            } while (File.Exists(targetFilePath));

            // �I�s workItemsContext �إߨӷ��u�@��
            workItems = new(isWithArchiectColumn);
            workItems.CreateSourceWorkItemsFile(targetFilePath!, CustomerName, ProjectName);

            string message = $"�d���ɮפw���\�x�s�� [{targetFilePath}]�I";

            // ��� MessageBox �æ� Yes / No ��ӿﶵ
            DialogResult ifOpenExcelDirectly = MessageBox.Show(message + "\n\n�O�_�P�ɶ}�Ҥu�@���ؽd�� Excel �ɮסH", $"�d���u{messageArchiect}�v�w���\�x�s�I", MessageBoxButtons.YesNo);

            // �O�_�n�� Excel �����}���ɮ�
            if (ifOpenExcelDirectly == DialogResult.Yes)
            {
                try
                {
                    // �إ� processExcel 
                    var processExcel = new Process();

                    // �}���ɮ�
                    processExcel.StartInfo = new ProcessStartInfo(@$"{targetFilePath}"!)
                    {
                        UseShellExecute = true
                    };
                    processExcel.Start();
                }
                catch (Exception ex)
                {
                    // �p�G�}���ɮץ���
                    MessageBox.Show($"���~!{ex.Message}", "�L�k�}���ɮסI");
                    MessageBox.Show("�L�k�}���ɮסI");
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

        // �}�Ҩӷ��u�@�M�� Excel �ɮ�
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
                openFileDialog.Filter = "Excel �ɮ� (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workItemsFilePath = openFileDialog.FileName;
                    StatusBarLabel.Text = $"�wŪ�J�ɮ�: {workItemsFilePath}";
                }
            }
            if (string.IsNullOrEmpty(workItemsFilePath))
            {
                MessageBox.Show("�Х���ܨӷ��ɮסI", "�S����������ɮסI");
                return;
            }
            // �I�s workItemsContext �B�z�ӷ��u�@��
            workItems = new(isWithArchiectColumn);
            workItems.ReadSourceWorkItemsFile(workItemsFilePath);
            if (!string.IsNullOrEmpty(workItems.errorMessage))
            {
                MessageBox.Show(workItems.errorMessage, "�Х��ư��U�C���~��A���sŪ���ɮסI");
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
                StatusBarLabel.Text = "�Х��ư����~��A���sŪ���ɮסI";
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
        // ���ͥؼФu�@���� Excel �ɮ�

        private void SaveQuotationReportTarget_Click(object sender, EventArgs e)
        {
            if (workItemsFilePath == null)
            {
                MessageBox.Show("�Х��z�L��檺�uŪ���ӷ��ɮסv����u�@�M��պ��I");
                return;
            }
            architectColumn = withArchiectColumn ? 8 : -1; // H ��
            deployerColumn = withArchiectColumn ? 10 : 8; // J �� H ��
            developerColumn = withArchiectColumn ? 12 : 10; // L �� J ��
            totalCostColumn = withArchiectColumn ? 14 : 12; // N �� L ��

            GenerateQuotationReport();

            // �զX�T��
            string message = $"�ɮסu{targetFilePath}�v�w�x�s�I";

            // ��� MessageBox �æ� Yes / No ��ӿﶵ
            DialogResult ifOpenExcelDirectly = MessageBox.Show(message + "\n\n�O�_�P�ɶ}�Ҥu�@���ئ������������ Excel �ɮסH", $"�wŪ���u{Path.GetFileName(workItemsFilePath)}�v�ò��ͤu�@���ئ������������I", MessageBoxButtons.YesNo);

            // �O�_�n�� Excel �����}���ɮ�
            if (ifOpenExcelDirectly == DialogResult.Yes)
            {
                try
                {
                    // �إ� processExcel 
                    var processExcel = new Process();

                    // �}���ɮ�
                    processExcel.StartInfo = new ProcessStartInfo(@$"{targetFilePath}"!)
                    {
                        UseShellExecute = true
                    };
                    processExcel.Start();
                }
                catch (Exception ex)
                {
                    // �p�G�}���ɮץ���
                    MessageBox.Show($"���~!{ex.Message}", "�L�k�}���ɮסI");
                    MessageBox.Show("�L�k�}���ɮסI");
                }
            }

            Application.Exit();
        }

        // �x�s�ؼФu�@���� Excel �ɮ�
        private void GenerateQuotationReport()
        {
            // ���ͥؼФu�@�� Excel �ɮצW�ٻP���|
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            do
            {
                // �ϥ� DateTime ���o��e����M�ɶ�
                string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");

                // �ؼ��ɮצW�١G�M�צ�����.xlsx - ���ͷs���ťժ�Excel�ɮװ����M�צ�����
                string costFileName = @$"[{workItems!.ProjectContext!.CustomerName}-{workItems.ProjectContext!.ProjectName}]{projCostSheetName}_{currentDate}.xlsx";

                // �ϥΨ��o�ؼ��ɮצW�٩M���|
                targetFilePath = Path.Combine(desktopPath, costFileName);
                // �ˬd����T�{���ɮרèS���ۦP�W�٪��ɦW�s�b
            } while (File.Exists(targetFilePath));

            // �}�l�D�n�{��
            using (ExcelPackage costPackage = new ExcelPackage())
            {
                // �ҩl���O
                var quotationSheet = costPackage.Workbook.Worksheets.Add(quotationSheetName);
                var projCostSheet = costPackage.Workbook.Worksheets.Add(projCostSheetName);
                var taskListSheet = costPackage.Workbook.Worksheets.Add(taskListSheetName);
                var deliverablesSheet = costPackage.Workbook.Worksheets.Add(deliverablesSheetName);

                // �ϥ� WorksheetBase ���l���O�Ө�U�g�J�U�u�@����� �N�U�u�@�������J using �϶��A�T�O�귽����
                using (QuotationWorksheet quotation = new QuotationWorksheet(quotationSheet, withArchiectColumn))
                using (ProjectCostWorksheet projectCost = new ProjectCostWorksheet(projCostSheet, withArchiectColumn))
                using (TaskListWorksheet taskList = new TaskListWorksheet(taskListSheet))
                using (DeliverablesWorkSheet deliverables = new DeliverablesWorkSheet(deliverablesSheet))
                {
                    // �w�q "���q�s��" �q 1 �}�l�p��
                    int phaseNumber = 1;

                    // �w�q "�Ǹ�" �q 1 �}�l�p��
                    int sequenceNumber = 1;

                    // ���q�����Ǥ���
                    string lastPhasePart = null!;

                    // �g�J�u�@����D
                    taskList.WriteHeader();
                    projectCost.WriteHeader();
                    quotation.WriteHeader();
                    deliverables.WriteHeader();

                    // Ū�����u�@���رN�u�@ï�̪���ƥΦU�u�@��U�۪���k�g�J�s�u�@ï
                    foreach (var phase in workItems.phaseData)
                    {
                        // phaseName �� Key �]�N�O���q�W��
                        string phaseName = phase.Key;

                        // ���q�W�٥� "-" �Ӥ��Φr��
                        string[] phaseNameParts = phaseName.Split('-');

                        // �^�����q�W�٤��Ϋᶥ�q���Ĥ@�������Ǥ����X�ӡA�èϥ� Trim() ���������ťզr��
                        string phasePart = phaseNameParts[0].Trim();

                        // taskLists �� Value �]�N�O�Ӷ��q�H TaskItemRow ���O�x�s����ƲM��
                        List<SourceWorkItemRow> taskLists = phase.Value;

                        // ���ثe�� ���q�����Ǥ��� �O�_�N�O�e���� ���q�����Ǥ��� �ۦP���W��
                        if (phasePart != lastPhasePart)
                        // �p�G�ثe�� ���q�����Ǥ��� �O�s�����q �N�H���q���}�l�����g�J�Ӷ��q����ƪ��Φ�
                        {
                            // ���H SetPhaseTitle �]�w PhaseTitle ���q���D�ثe����m
                            WorksheetBase.SetPhaseTitle();

                            // �g�J "���q�W��"
                            taskList.WriteText(phaseName, 1);
                            projectCost.WriteText(phaseName, 1);
                            quotation.WriteText(phaseName, 1);

                            // �����@�_����U�@�C
                            WorksheetBase.MoveSharedRowToNext();
                        }
                        // �w�q "���q�s��" �᪺ "�I"->"�j���s��" �q 1 �}�l�p��
                        int outlineNumber = 1;

                        // ���H SetPhaseStart �]�w PhaseStart ���q���إثe����m
                        WorksheetBase.SetPhaseStart();

                        // �}�l�q taskLists �M�椺�v�@���X�u�@���ظ�Ƽg�J "���q����" 
                        foreach (var item in taskLists)
                        {
                            // �p�G�ثe �j���s�� �� 1 ��ܳo�O���q�}�Y���Ĥ@�ӽs��
                            if (item.TaskName != "�L")
                            {
                                if (outlineNumber == 1)
                                {
                                    // �u���b���q�}�Y���Ĥ@�ӽs���ɤ~�g�J���q�s��
                                    taskList.WriteText(phaseNumber, 1);
                                    projectCost.WriteText(phaseNumber, 1);
                                }

                                // �g�J�j���s��
                                taskList.WriteText($"{phaseNumber}.{outlineNumber}", 2);

                                // �g�J �Ǹ�
                                quotation.WriteValue(sequenceNumber, 1);
                                projectCost.WriteValue(sequenceNumber, 2);

                                // �g�J �u�@�Ѽ�
                                taskList.WriteText(item.TotalTaskDays, mandaySubtotal, isRight: false); ;
                                projectCost.WriteText(item.TotalTaskDays, mandaySubtotal, isRight: false); ;

                                // �g�J �M�׸g�z
                                projectCost.WriteValue(item.PrjManagerDays, projectManagerColumn, isRight: false); ;
                                projectCost.WriteNumeric(8000, projectManagerColumn + 1);

                                // �g�J �[�c�v
                                if (withArchiectColumn)
                                {
                                    projectCost.WriteValue(item.ArchiectDays, architectColumn, isRight: false); ;
                                    projectCost.WriteNumeric(8000, architectColumn + 1);
                                }

                                // �g�J ���p��
                                projectCost.WriteValue(item.DeployerDays, deployerColumn, isRight: false); ;
                                projectCost.WriteNumeric(8000, deployerColumn + 1);

                                // �g�J �t�d���w�]�Ȭ� "AEB"
                                taskList.WriteText("AEB", 9);

                                // �g�J �}�o��
                                projectCost.WriteValue(item.DeveloperDays, developerColumn);
                                projectCost.WriteNumeric(8000, developerColumn + 1);
                                projectCost.WriteCostSumFormula(totalCostColumn);
                            }

                            // �g�J �u�@����
                            taskList.WriteText(item.TaskName, 3, isCenter: false);
                            projectCost.WriteText(item.TaskName, 3, isCenter: false);
                            quotation.WriteText(item.TaskName, 2, isCenter: false);

                            // �g�J �u�@����
                            taskList.WriteText(item.TaskDescription, 4, isCenter: false);
                            projectCost.WriteText(item.TaskDescription, 4, isCenter: false);
                            quotation.WriteText(item.TaskDescription, 3, isCenter: false);

                            // ����U�@�C
                            WorksheetBase.MoveSharedRowToNext();

                            // �j���s���[ 1
                            outlineNumber++;

                            // �Ǹ��[ 1
                            sequenceNumber++;
                        }
                        // �g�J �u�@����
                        WorksheetBase.SetPhaseEnd();

                        // �X�� �u�@����
                        taskList.MergeText(3);
                        projectCost.MergeText(3, sheetCalculate: true);
                        quotation.MergeText(2, sheetCalculate: true);

                        // �ϥΦ۩w�q�C��榡�ƶ��q 
                        taskList.FormatPhase();

                        // �ϥΦ۩w�q�C��榡�ƶ��q 
                        projectCost.FormatPhase();

                        // �ϥΦ۩w�q�C��榡�ƶ��q 
                        quotation.FormatPhase();
                        if (phasePart != lastPhasePart)
                        {
                            quotation.PhaseSumPrice(workItems.phaseCount[phasePart]);
                        }

                        // ���q�s���[ 1
                        phaseNumber++;

                        // ��Ӷ��q������A�N�ثe�� phasePart, �]�N�O���q�W�٪����q���Ǥ������� ������ lastPhasePart�C�ΨӦb�U���q�P�_�O�_�٬O�ۦP���q����
                        lastPhasePart = phasePart;
                    }
                    // �̫�w��b��檺�̫�@�欰 SetPhaseTitle�C�o�O���F���������r�Ϭq����m���ѦҨ̾�
                    WorksheetBase.SetPhaseTitle();

                    // �g�J���
                    taskList.WriteFooter();

                    // �g�J���
                    projectCost.WriteFooter();

                    // �g�J���
                    quotation.WriteFooter();

                    // �g�J���
                    deliverables.WriteFooter();

                    // �̫�׹� - �[�J�M�׸�T
                    deliverablesSheet.Cells[2, 1].Value = customerNameTextBox.Text;
                    deliverablesSheet.Cells[2, 2].Value = projectNameTextBox.Text;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    deliverablesSheet.Cells[1, 1, 2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // �̫�׹� - �[�J�M�ץ�I����
                    if (deliverableListBox.Items.Count > 0)
                    {
                        deliverablesSheet.Cells[4, 1].Value = "��I����";
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

                    // �̫�׹� - �[�J�M�׸�T
                    taskListSheet.Cells[1, 1].Formula = $"={deliverablesSheetName}!A2&\"-\"&{deliverablesSheetName}!B2&\"-{taskListSheetName}\"";
                    projCostSheet.Cells[1, 1].Formula = $"={deliverablesSheetName}!A2&\"-\"&{deliverablesSheetName}!B2&\"-{projCostSheetName}\"";

                    // �̫�׹� - �[�J�M�׸�T
                    string[] columnNames = ["�Ȥ�W��", "�M�צW��", "�~�ȳ���", "�~�ȥN��", "�q�l�H�cs", "�q�ܤ���s", "�������", "�����渹", "�޳N����", "�����N��", "�q�l�H�ct", "�q�ܤ���t"];
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�Ȥ�W��") % 7), 1 + ((Array.IndexOf(columnNames, "�Ȥ�W��") / 7) * 3)].Value = "    �Ȥ�W��" + "�G" + customerNameTextBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�M�צW��") % 7), 1 + ((Array.IndexOf(columnNames, "�M�צW��") / 7) * 3)].Value = "    �M�צW��" + "�G" + projectNameTextBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�~�ȳ���") % 7), 1 + ((Array.IndexOf(columnNames, "�~�ȳ���") / 7) * 3)].Value = "    �~�ȳ���" + "�G" + salesDepartmentComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�~�ȥN��") % 7), 1 + ((Array.IndexOf(columnNames, "�~�ȥN��") / 7) * 3)].Value = "    �~�ȥN��" + "�G" + salesRepresentativeComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�q�l�H�cs") % 7), 1 + ((Array.IndexOf(columnNames, "�q�l�H�cs") / 7) * 3)].Value = "    �q�l�H�c" + "�G" + salesRepresentativeEmailAddress.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�q�ܤ���s") % 7), 1 + ((Array.IndexOf(columnNames, "�Ȥ�W��") / 7) * 3)].Value = "    �q�ܤ���" + "�G" + salesRepresentativePhoneExtension.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�������") % 7), 1 + ((Array.IndexOf(columnNames, "�������") / 7) * 3)].Value = "    �������" + "�G" + DateTime.Now.ToString("yyyy/MM/dd");
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�����渹") % 7), 1 + ((Array.IndexOf(columnNames, "�����渹") / 7) * 3)].Value = "�����渹" + "�G";
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�޳N����") % 7), 1 + ((Array.IndexOf(columnNames, "�޳N����") / 7) * 3)].Value = "�޳N����" + "�G" + techDepartmentComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�����N��") % 7), 1 + ((Array.IndexOf(columnNames, "�����N��") / 7) * 3)].Value = "�����N��" + "�G" + techRepresentativeComboBox.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�q�l�H�ct") % 7), 1 + ((Array.IndexOf(columnNames, "�q�l�H�ct") / 7) * 3)].Value = "�q�l�H�c" + "�G" + techRepresentativeEmailAddress.Text;
                    quotationSheet.Cells[3 + (Array.IndexOf(columnNames, "�q�ܤ���t") % 7), 1 + ((Array.IndexOf(columnNames, "�q�ܤ���t") / 7) * 3)].Value = "�q�ܤ���" + "�G" + techRepresentativePhoneExtension.Text;

                    // �O�s��ؼ��ɮ�
                    costPackage.SaveAs(new FileInfo(targetFilePath!));
                    StatusBarLabel.Text = $"������w�x�s: {targetFilePath}";
                    // �����ɱN����]�� null �U���^��
                }
            }
        }

        private void AddDeliverableButton_Click(object sender, EventArgs e)
        {
            string selection = deliverableSelectionComboBox.Text;
            // �ˬd�m�W�O�_���Ŧr��
            if (selection == "")
            {
                MessageBox.Show("�п�ܥ�I���ئA���s�W�I");
                // ���}���ƥ�B�z�禡
                return;
            }
            if (deliverableListBox.Items.Contains(selection))
            {
                MessageBox.Show("��Ƥw�s�b!");
            }
            else
            {
                deliverableListBox.Items.Add(selection);
            }
        }

        private void ModifyDeliverableButton_Click(object sender, EventArgs e)
        {
            // �ˬd�O�_��ܤF ListBox ��������
            if (deliverableListBox.SelectedItem == null)
            {
                MessageBox.Show("�п�ܭn�ק諸��I���ءI");
                return;
            }

            // ����襤������
            string selectedItem = deliverableListBox.SelectedItem.ToString()!;

            // ���ܨϥΪ̿�ܷs����
            string newSelection = deliverableSelectionComboBox.Text;

            // �ˬd�s�ﶵ�O�_����
            if (string.IsNullOrEmpty(newSelection))
            {
                MessageBox.Show("�п�ܷs����I���ءI");
                return;
            }

            // �ˬd�s���جO�_�w�g�s�b�� ListBox ��
            if (deliverableListBox.Items.Contains(newSelection))
            {
                MessageBox.Show("��ܪ���I���ؤw�g�s�b�I");
                return;
            }

            // �����襤������
            int selectedIndex = deliverableListBox.SelectedIndex;
            deliverableListBox.Items[selectedIndex] = newSelection;
        }

        private void DeleteDeliverableButton_Click(object sender, EventArgs e)
        {
            // �ˬd�O�_��ܤF ListBox ��������
            if (deliverableListBox.SelectedItem == null)
            {
                MessageBox.Show("�п�ܭn�R������I���ءI");
                return;
            }

            // �T�{�R��
            DialogResult result = MessageBox.Show("�T�w�n�R���襤����I���ضܡH", "�R���T�{", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                // �R���襤������
                deliverableListBox.Items.Remove(deliverableListBox.SelectedItem);
            }
        }

        private void LicensingStatement_Click(object sender, EventArgs e)
        {
            MessageBox.Show("���{���� EPPlus 7.5.3 �~��l������ MIT ���v�C��������ϥ� EPPlus 7.5.3 �M��}�o�C�ѩ�EPPlus �M��w�ܧ����v�覡�� LGPL �אּ Polyform Noncommercial 1.0.0�C�b���s�����v�\�i�覡�U�AEPPlus �b�Y�Ǳ��p�U���M�i�H�K�O�ϥΡA���b�ӷ~�~�Ȥ��ϥΫh�ݭn�ӷ~�\�i�ҡC���{�����w EPPlus �M��K�O�����ϥνd��A�п�Ө���v���ڨ̨�A�ΩʨӨϥΥ��n��C");
        }

        private void VersionDescription_Click(object sender, EventArgs e)
        {
            MessageBox.Show("���{���ϥ� C# .net 9 �}�o�C�ثe�������u�@���ئ�������������X�[�c�v��쪩���C������ 0.1c.3�C");
        }

        private void AuthorInformation_Click(object sender, EventArgs e)
        {
            MessageBox.Show("���{���� alvin.lin@outlook.com �W�O�}�o�C�}�o������� 2025-1-25 ��");
        }
    }
}
