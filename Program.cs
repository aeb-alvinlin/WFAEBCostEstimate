using OfficeOpenXml;

namespace WFormProjEstimateApp1
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            // 來源檔案路徑預設為空的
            string? sourceExcelFilePath = null;
            // 檢查是否有傳遞參數
            if (args.Length > 0)
            {
                // 只取第一個參數 args[0]，並檢查檔案路徑以確保檔案確實存在
                if (File.Exists(args[0]))
                {
                    // 檔案確實存在的話就把 sourceExcelFilePath 作為用來開啟檔案的來源 Excel 路徑 sourceExcelFilePath
                    sourceExcelFilePath = args[0];
                }
            }

            // 將檔案路徑傳入 WinForm 啟動應用程式。如果沒有檔案路徑預設為空的
            Application.Run(new WFormProjEstimate(sourceExcelFilePath));     
        }
    }
}