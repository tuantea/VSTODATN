using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using ZXing;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using ZXing.QrCode.Internal;

namespace DATN.FuntionExcel
{
    internal class QRCoder
    {
        public static void GenerateQRCode(Excel.Worksheet worksheet, Excel.Range actCell)
        {

            string data = actCell.Value;
            BarcodeWriter barcodeWriter = new BarcodeWriter();
            barcodeWriter.Format = BarcodeFormat.QR_CODE;
            barcodeWriter.Options = new ZXing.Common.EncodingOptions
            {
                Width = 200,
                Height = 200
            };

            // Tạo bitmap từ mã QR
            if (data != null)
            {
                var bitmap = barcodeWriter.Write(data);
                string tempFile = Path.GetTempFileName();
                bitmap.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);
                Shapes shapes = worksheet.Shapes;
                Shape pictureShape = shapes.AddPicture(tempFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 100, 100, 100, 100);
                //Shape pictureShape = null;
                //pictureShape = worksheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, actCell.Left, actCell.Top, actCell.Width, actCell.Height);
                //pictureShape.Fill.UserPicture(tempFile);
                //System.Windows.Forms.Clipboard.SetImage(bitmap);
                //worksheet.Paste(worksheet.Range["A1"]);

                //worksheet.Pictures[worksheet.Pictures.Count].ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                //worksheet.Cells.EntireColumn.ColumnWidth = 15;
            }
            else
            {
                MessageBox.Show("Cell không có nội dung");
            }
        }
        public static string GenerateQRCode1(string data)
        {

            
            BarcodeWriter barcodeWriter = new BarcodeWriter();
            barcodeWriter.Format = BarcodeFormat.QR_CODE;
            barcodeWriter.Options = new ZXing.Common.EncodingOptions
            {
                Width = 200,
                Height = 200
            };
            
                var bitmap = barcodeWriter.Write(data);
                string tempFile = Path.GetTempFileName();
            bitmap.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);
            return tempFile;
        }

        [ExcelDna.Integration.ExcelFunction(Description = "Tạo mã QRCode")]
        public static object QRCode(
            
            [ExcelDna.Integration.ExcelArgument(Description = "Tên của Shape sẽ chứa ảnh QRCode (xem bằng Selection Pane). Nếu shape chưa tồn tại, hàm sẽ tự tạo mới. Ví dụ: tl123")]
            string ShapeName,
            [ExcelDna.Integration.ExcelArgument(Description = "Văn bản cần chuyển thành QRcode. Ví dụ: xin chào bạn")]
            string Text)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return "";

            Worksheet ws = wb.ActiveSheet;

            Shape MyShape = null;

            /// Tìm xem có Shape nào có tên như tham số vào không
            foreach (Shape shape in ws.Shapes)
                if (shape.Name == ShapeName)
                {
                    MyShape = shape;
                }

            /// Nếu chưa có Shape thì tự tạo  mới luôn
            if (MyShape == null)
            {
                MyShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, xlApp.ActiveCell.Left, xlApp.ActiveCell.Top, xlApp.ActiveCell.Width, xlApp.ActiveCell.Height);
                MyShape.Name = ShapeName;
                MyShape.Line.Transparency = (float)(1.0);
                MyShape.Fill.Solid();
                MyShape.Fill.ForeColor.RGB = 0xEEEEEE;
            };
            /// Và đặt vào đó hình ảnh QRCode
            {
                try
                {

                    MyShape.Fill.UserPicture(GenerateQRCode1(Text));
                }
                catch
                {
                    return "Disconnect";
                }
            }
            return Text;
        }
    }
}
