

using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Raqmana
{
    public partial class MyRibbon
    {
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            if (app.ActiveWorkbook == null) return;

            List<string> errorList = new List<string>();
            bool isProcessed;

            app.ScreenUpdating = false;
            try
            {
                SheetProcessor.ProcessWorkbook(app.ActiveWorkbook, errorList, out isProcessed);
                UIHelpers.ShowFinalReport(errorList, isProcessed, app.ActiveWorkbook);
            }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show("خطأ: " + ex.Message); }
            finally { app.ScreenUpdating = true; }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            UIHelpers.ShowAboutBox();
        }
    }
}