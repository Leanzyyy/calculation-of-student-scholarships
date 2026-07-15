using System;
using System.Windows.Forms;
using OfficeOpenXml;

namespace PRAK1
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}