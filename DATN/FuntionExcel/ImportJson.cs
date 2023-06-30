using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;           

namespace DATN.FuntionExcel
{
    internal class ImportJson
    {
        public static void ImportJsonToExcel(Excel.Application excellApp)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog.Title = "Open File";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string jsonContent = File.ReadAllText(openFileDialog.FileName);
            dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
            bool isVisible = (bool)jsonData["config"]["visible"];
            bool activateSheet = (bool)jsonData["config"]["activatesheet"];
            bool terminate = (bool)jsonData["config"]["terminate"];
            excellApp.Visible = isVisible;
            Excel.Workbook workbook = excellApp.Workbooks.Add();
            JArray sheets = (JArray)jsonData["sheets"];
            foreach (JObject sheetData in sheets)
            {
                string sheetName = (string)sheetData["name"];
                bool checkSheetName = false;
                Excel.Worksheet worksheet;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    
                    if (sheet.Name == sheetName) {
                        checkSheetName = true;
                        
                    }
                }
                // Tạo sheet mới trong workbook
                if (!checkSheetName)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = sheetName;
                }
                else
                {
                    worksheet = workbook.Worksheets[sheetName];
                }
                JArray cells = (JArray)sheetData["cells"];
                foreach (JObject cellData in cells)
                {
                    string cellPos = (string)cellData["pos"];
                    string cellValue = (string)cellData["value"];

                    Excel.Range cell = worksheet.Range[cellPos];
                    cell.Value = cellValue;
                }
            }

            // Kích hoạt sheet đầu tiên
            if (activateSheet)
            {
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Sheets[1];
                firstSheet.Activate();
            }
            if (terminate)
            {
                excellApp.Quit();
            }
        }

    }
}
