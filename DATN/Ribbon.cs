using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Excel;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Speech.Recognition;
using DATN.FuntionExcel;

namespace DATN
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        static private dynamic GetSellectedCells()
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (wb == null)
            {
                MessageBox.Show("Bạn phải mở một workbook.");
                return null;
            }

            /// Lấy sheet đang được active hiện thời
            Worksheet ws = wb.ActiveSheet;

            // Lẩy ra nội dung của cell hiện thời đang được chọn
            Range res = (Range)Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            return res;
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            if(actCell.Value2 != null )
            {
                string value =actCell.Value2;
                string sText = actCell.Text;
                System.Windows.Forms.MessageBox.Show(sText);
            }
        }
        private void button_ExChange(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            actCell.Value=ExchangeRate.GetExchangeRate();
        }
        private void button_ReadNumber(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            if (actCell.Value2 != null)
            {
                int sText = Int32.Parse(actCell.Text);
                
                string text=Number.NumberToText(sText);
                actCell.Value = actCell.Text + text;
            }
        }
        private void button_ExportExcelToJson(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excellApp = Globals.ThisAddIn.Application;
            ExportJson.ExportExcelToJson(excellApp);
        }
        private void button_ExportExcelToJson1(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            ExportJson.ExportExcelToJson1(activeWorkbook);
        }
        private void button_QRCoder(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            QRCoder.GenerateQRCode(activeWorksheet,actCell);
        }
        private void button_ImportJsonToExcel(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excellApp = Globals.ThisAddIn.Application;
            ImportJson.ImportJsonToExcel(excellApp);
        }
        private void buttonImage2Cells_Click(object sender, RibbonControlEventArgs e)
        {   //load image into cells - by MSc Tien
            Bitmap img;
            const int MAX_HEIGHT = 320;
            const int MAX_PIXEL = 82455; //chính xác đúng ngần này điểm

            /// Tạo dialog để chọn file ảnh
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                /// Chỉ chấp nhận file dạnh ảnh
                dialog.Filter = "image files (*.jpg)|*.jpg|*.png|*.png|*.bmp|*.bmp|All files (*.*)|*.*";
                dialog.FilterIndex = 1;

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                /// Mở file ảnh
                img = new Bitmap(dialog.FileName);

                /// Tự co giãn tỷ lệ theo số điểm tối đa            
                double ratio;
                int newwwidth = img.Width;
                int newheight = img.Height;
                if (img.Width * img.Height > MAX_PIXEL)
                {
                    ratio = Math.Sqrt((double)MAX_PIXEL / img.Width / img.Height);
                    newwwidth = (int)(img.Width * ratio);
                    newheight = (int)(img.Height * ratio);
                    img = ResizeBitmap(img, newwwidth, newheight);
                }

                /*
                /// Tự co giãn tỷ lệ theo chiều cao và chiều dọc để không vượt qua
                ratio = img.Width / img.Height;
                if (newheight > MAX_HEIGHT)
                {
                    newheight = MAX_HEIGHT;
                    newwwidth = (int)(newheight * ratio);
                }
                if (newwwidth > MAX_HEIGHT)
                {
                    newwwidth = MAX_HEIGHT;
                    newheight = (int)(newwwidth / ratio);
                }
                img = ResizeBitmap(img, newwwidth, newheight);
                */



                dialog.Dispose();
            }



            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (wb == null)
            {
                MessageBox.Show("Bạn phải mở một workbook.");
                img.Dispose();
                return;
            }

            /// Lấy sheet đang được active hiện thời
            Excel.Worksheet ws = wb.ActiveSheet;

            // Lẩy ra nội dung của cell hiện thời đang được chọn
            //((dynamic)ws.Application.ActiveCell).Value.ToString()

            Task.Run(() =>
            {
                /// Đưa chiều rộng của các cột bằng nhau và bằng 2
                for (int i = 1; i <= img.Width; i++)
                {
                    ws.Columns[i].ColumnWidth = 2;
                }
            });

            Task.Run(() =>
            {
                /// Biến đếm số điểm ảnh 
                int count = 0;
                /// Đặt cờ báo hiệu có lỗi trong quá trình convert
                bool error_flag = false;
                /// Đọc từng picel ảnh và qui đổi thành màu nền của cell
                for (int i = 0; i < img.Height; i++)
                {
                    for (int j = 0; j < img.Width; j++)
                    {
                        Color pixel = img.GetPixel(j, i);
                    retry:
                        try
                        {
                            ws.Cells[i + 1, j + 1].Interior.Color = pixel.R | (pixel.G << 8) | (pixel.B << 16);
                        }
                        catch (Exception ex)  //user click chuột vào 1 cell là sinh ngoại lệ và dừng ngay.
                        {
                            if ((uint)ex.HResult == 0x800a03ec)
                            {
                                error_flag = true;
                                goto _end_of_image;
                            }
                            else
                            {
                                Debug.WriteLine(ex.Message);
                                Debug.WriteLine("i={0}, j={1}", i, j);
                                Task.Delay(250);
                                goto retry;
                            }
                        }
                        count++;
                        if (count == 82455)
                        {
                            int x;
                            x = count;
                        }
                    }

                }
            _end_of_image:
                if (error_flag)
                {
                    MessageBox.Show("Excel cannot process too many different cell formats. Please create another workbook.", "Error 0x800a03ec");
                }
                else
                {
                    MessageBox.Show("Finish converting from image " + img.Height + "x" + img.Width + " pixels to " + count + " cells. Have fun!");
                }
                img.Dispose();
            }
            );

        }

        private void buttonSpeechVN_Click(object sender, RibbonControlEventArgs e)
        {
            /// Lấy ra vùng cell đang được select
            Range CurrentRange = GetSellectedCells();
            if (CurrentRange == null) return;

            /*
            if (Properties.Settings.Default.CELL_ORDER == "row")
            {
                CurrentRange.Sort(XlSortOrientation.xlSortRows);
                CurrentRange.ReadingOrder()
            }
            else if (Properties.Settings.Default.CELL_ORDER == "column")
            {
                CurrentRange.Sort(XlSortOrientation.xlSortColumns);
            }
            */

            Debug.WriteLine("Selected Range has {0} columns, {1} rows", CurrentRange.Columns.Count, CurrentRange.Rows.Count);

            if (Properties.Settings.Default.CELL_ORDER == "select")
            {
                /// Lần lượt đọc từng cell theo thứ tự select
                foreach (var mycell in CurrentRange.Cells)
                {
                    if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                    {
                        //SpeechProcessing.ReadMeByResponsiveVoice(((dynamic)(mycell)).Value.ToString());
                        SpeechProcessing.ReadMeByCotana(((dynamic)(mycell)).Value.ToString());
                    }
                }
            }
            else if (Properties.Settings.Default.CELL_ORDER == "row")
            {

                /// Lần lượt đọc từng cell
                for (int row = 1; row <= CurrentRange.Rows.Count; row++)
                    for (int col = 1; col <= CurrentRange.Columns.Count; col++)
                    {
                        var mycell = CurrentRange.Cells[row, col];
                        if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                        {
                            //SpeechProcessing.ReadMeByResponsiveVoice(((dynamic)(mycell)).Value.ToString());
                            SpeechProcessing.ReadMeByCotana(((dynamic)(mycell)).Value.ToString());
                        }
                    }
            }
            else if (Properties.Settings.Default.CELL_ORDER == "column")
            {

                /// Lần lượt đọc từng cell ưu tiên cột trước
                for (int col = 1; col <= CurrentRange.Columns.Count; col++)
                    for (int row = 1; row <= CurrentRange.Columns.Count; row++)
                    {
                        var mycell = CurrentRange.Cells[row, col];
                        if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                        {
                            //SpeechProcessing.ReadMeByResponsiveVoice(((dynamic)(mycell)).Value.ToString());
                            SpeechProcessing.ReadMeByCotana(((dynamic)(mycell)).Value.ToString());
                        }
                    }
            }

        }

        static Bitmap ResizeBitmap(Bitmap img, int width, int height)
        {
            Bitmap result = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(result))
            {
                g.DrawImage(img, 0, 0, width, height);
            }

            return result;
        }
    }
}
