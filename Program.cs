using System;
using WF = System.Windows.Forms;
using OfficeOpenXml;

namespace MyApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // ✅ EPPlus 7 이하: 비상업용 라이선스 설정
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            WF.Application.SetHighDpiMode(WF.HighDpiMode.SystemAware);
            WF.Application.EnableVisualStyles();
            WF.Application.SetCompatibleTextRenderingDefault(false);
            WF.Application.Run(new MainForm());
        }
    }
}
