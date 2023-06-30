using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Spire.Xls;
using Spire.Xls.Collections;
using Spire.Xls.Core;
using Microsoft.Office.Tools.Excel;
using System.Text.Json.Nodes;

namespace DATN.FuntionExcel
{
    internal class ExportJson
    {
        public static void ExportExcelToJson(Excel.Application excellApp)
        {
            // Xác định phạm vi dữ liệu trong Excel (ví dụ: từ ô A1 đến ô C10)
            int startRow = 1;
            int startColumn = 1;
            int endRow = 10;
            int endColumn = 100;

            JObject json = new JObject();
            JObject config = new JObject();
            config["visible"] = excellApp.Visible;
            config["activatesheet"] = true;
            config["terminate"] = false;
            json["config"] = config;
            JArray sheets = new JArray();
            JArray cells = new JArray();
            Excel.Workbook workbook = excellApp.ActiveWorkbook;
            foreach (Excel.Worksheet sheet1 in workbook.Sheets)
            {
                JObject sheet = new JObject();
                sheet["name"] = sheet1.Name;
                sheet["visible"] = true;
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int column = startColumn; column <= endColumn; column++)
                    {
                       
                        Excel.Range cell = sheet1.Cells[row, column];
                        if (cell.Value2 != null)
                        {
                            JObject cell1 = new JObject();
                            if (column > 26)
                            {
                                cell1["pos"] = ((char)(column / 26 + 64)).ToString() + ((char)(column % 26 + 64)).ToString() + row;
                            }
                            else
                            {
                                cell1["pos"] = ((char)(column % 26 + 64)).ToString() + row;
                            }
                            cell1["value"] = cell.Value != null ? cell.Value.ToString() : string.Empty;
                            cells.Add(cell1);
                        }
                      
                    }
                }

                sheet["cells"] = cells;
                sheets.Add(sheet);
            }

            json["sheets"] = sheets;

            string jsonContent = json.ToString();

            try
            {

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
                saveFileDialog.Title = "Save File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    FileStream fileStream = File.Create(filePath);
                    fileStream.Close();
                    File.WriteAllText(filePath, jsonContent);
                                  
                }
                else
                {
                    MessageBox.Show("File creation canceled by the user.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating the file: " + ex.Message);
            }
        }

        public static void ExportExcelToJson1(Excel.Workbook workbook)
        {          
            JArray jsonArray = new JArray();
            DocumentProperties properties = workbook.BuiltinDocumentProperties;
            JObject jsonObject = new JObject();
            //foreach (DocumentProperty prop in properties)
            //{
            //    string propertyName = prop.Name;
            //    string propertyValue;
            //    //string ab=prop.Value;
            //    if (prop.Name == "Last print date"||prop.Name=="Creation date"||prop.Name=="Last save time"||prop.Name=="Total editing time"||prop.Name== "Security" || prop.Name.Contains("Number"))
            //    { propertyValue = prop.Value.ToString();
            //        string abc = "234";
            //    }
            //    else
            //        propertyValue = prop.Value;
            //    jsonObject.Add(propertyName, propertyValue);

            //    // Do something with the property name and value
            //}
            //jsonArray.Add(jsonObject);
            //Get the custom properties of the workbook
            //DocumentProperties customProperties = workbook.CustomDocumentProperties;
            //foreach (DocumentProperty customProp in customProperties)
            //{
            //    string propertyName = customProp.Name;
            //    string propertyValue = customProp.Value;
            //    jsonObject.Add(propertyName, propertyValue);
            //    // Do something with the custom property name and value
            //}
            ICustomDocumentProperties customProperties = workbook.CustomDocumentProperties;
            for (int i = 0; i < customProperties.Count; i++)
            {
                string propertyName = customProperties[i].Name;
                string propertyValue = customProperties[i].Text;
                jsonObject.Add(propertyName, propertyValue);
                // Do something with the custom property name and value
            }
            jsonArray.Add(jsonObject);
            string jsonContent = jsonArray.ToString();

            // Lưu chuỗi JSON vào tệp
            try
            {
                // Khởi tạo đối tượng SaveFileDialog
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
                saveFileDialog.Title = "Save File";

                // Mở hộp thoại Save File Dialog và chờ người dùng chọn vị trí và tên tệp
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Lấy đường dẫn và tên tệp từ SaveFileDialog
                    string filePath = saveFileDialog.FileName;

                    // Tạo tệp mới
                    FileStream fileStream = File.Create(filePath);
                    fileStream.Close();
                    File.WriteAllText(filePath, jsonContent);

                }
                else
                {
                    MessageBox.Show("File creation canceled by the user.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating the file: " + ex.Message);
            }

        }
        
    }
}