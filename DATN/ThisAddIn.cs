using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using DATN;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace DATN
{
    public partial class ThisAddIn
    {

        public void SumRange()
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection;
            double sum = 0.0;

            foreach (Excel.Range cell in selectedRange.Cells)
            {
                if (cell.Value2 != null)
                {
                    double value = 0.0;
                    if (double.TryParse(cell.Value2.ToString(), out value))
                    {
                        sum += value;
                    }
                }
            }

            MessageBox.Show("Tổng của các ô được chọn là: " + sum.ToString());
        }
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }
        public Excel.Worksheet getActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        public Excel.Workbook getActiveWorkbook()
        {
            return (Excel.Workbook)Application.ActiveWorkbook;

        }
        public Excel.Application getActiveApp()
        {
            return (Excel.Application)Application.Application;

        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
