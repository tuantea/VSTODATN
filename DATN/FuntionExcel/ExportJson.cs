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
namespace DATN.FuntionExcel
{
    internal class ExportJson
    {
        public static void ExportExcelToJson(Excel.Worksheet worksheet)
        {
            // Xác định phạm vi dữ liệu trong Excel (ví dụ: từ ô A1 đến ô C10)
            int startRow = 1;
            int startColumn = 1;
            int endRow = 10;
            int endColumn = 3;

            // Tạo một đối tượng JArray để lưu trữ dữ liệu từ Excel
            JArray jsonArray = new JArray();

            // Lặp qua các ô trong phạm vi và thêm dữ liệu vào JArray
            for (int row = startRow; row <= endRow; row++)
            {
                JObject jsonObject = new JObject();
                for (int column = startColumn; column <= endColumn; column++)
                {
                    Excel.Range cell = worksheet.Cells[row, column];
                    string cellValue = cell.Value != null ? cell.Value.ToString() : string.Empty;
                    string columnName = ((char)(column + 64)).ToString()+row; // Chuyển đổi số cột thành chữ cái tương ứng (A, B, C, ...)
                    jsonObject.Add(columnName, cellValue);
                }
                jsonArray.Add(jsonObject);
            }

            // Chuyển đổi JArray thành chuỗi JSON
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