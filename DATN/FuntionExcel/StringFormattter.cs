using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DATN.FuntionExcel
{
    internal class StringFormattter
    {   public static string LayTenSV(string ten)
        {
            int i = ten.LastIndexOf(" ");

            string ho = ten.Substring(i + 1);
            return ho;
        }
        private static readonly string[] VietnameseSigns = new string[]
        {

            "aAeEoOuUiIdDyY",

            "áàạảãâấầậẩẫăắằặẳẵ",

            "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ",

            "éèẹẻẽêếềệểễ",

            "ÉÈẸẺẼÊẾỀỆỂỄ",

            "óòọỏõôốồộổỗơớờợởỡ",

            "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ",

            "úùụủũưứừựửữ",

            "ÚÙỤỦŨƯỨỪỰỬỮ",

            "íìịỉĩ",

            "ÍÌỊỈĨ",

            "đ",

            "Đ",

            "ýỳỵỷỹ",

            "ÝỲỴỶỸ"
        };
        static public string RemoveAccent(string text)
        {
            for (int i = 1; i < VietnameseSigns.Length; i++)
            {
                for (int j = 0; j < VietnameseSigns[i].Length; j++)
                    text = text.Replace(VietnameseSigns[i][j], VietnameseSigns[0][i - 1]);
            }
            return text;
        }
        public static string LayHoSV(string ten)
        {
            string[] temp= ten.Split(' ');

            //string ho = ten.Substring(i + 1);
            return temp[0];
        }
        /// <summary>
        /// Return original text and change cell background color
        /// </summary>
        /// <param name="text"></param>
        /// <param name="CellBackgroundColor"></param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Change cell background color")]
        public static object StringCellFormatter(
            [ExcelDna.Integration.ExcelArgument(Description = "Original text")] string text,
            [ExcelDna.Integration.ExcelArgument(Description = "cell background color")] string CellBackgroundColor)
        {
            var result = Color.FromName(CellBackgroundColor);
            if (result.IsKnownColor) //check if the color string is valid
            {

            }
            return text;
        }
        static public string LayCacChuCaiDau(string FullName)
        {
            int i;
            int len;
            string ShorthenName = String.Empty;

            /// Loại bỏ kí tự trống ở cả 2 đầu cuối nếu có
            FullName = FullName.Trim();

            len = FullName.Length;

            /// Nếu chỉ có 1 kí tự thì xong luôn
            if (len == 1)
            {
                ShorthenName = FullName;
            }
            else
            {
                ShorthenName = ShorthenName + FullName[0];
                for (i = 0; i < len; i++)
                {
                    if ((FullName[i] == ' ') && (i < (len - 1)) && (FullName[i + 1] != ' '))
                    {
                        ShorthenName = ShorthenName + FullName[i + 1];
                        i++;
                    }
                }
            }
            return ShorthenName;
        }
        static public string LayDemSV(string FullName)
        {
            int pos_end;
            int pos_start;
            int len;

            /// Loại bỏ kí tự trống ở cả 2 đầu cuối nếu có
            FullName = FullName.Trim();

            len = FullName.Length;

            /// Tìm vị trí kí tự space đầu tiên
            for (pos_start = 1; pos_start < len; pos_start++)
            {
                if (FullName[pos_start] == ' ')
                    break;
            }
            if (pos_start >= len)
            {
                return String.Empty;
            }

            /// Bỏ qua các kí tự trống
            for (pos_start++; pos_start < len && (FullName[pos_start] == ' '); pos_start++) { };

            /// Tìm vị trí kí tự space cuối cùng
            for (pos_end = len - 1; pos_end > pos_start; pos_end--)
            {
                if (FullName[pos_end] == ' ')
                    break;
            }
            if (pos_end < pos_start)
            {
                return String.Empty;
            }
            else
            {
                return FullName.Substring(pos_start, pos_end - pos_start);
            }
        }
        /// <summary>
        /// Trả về Tên của Sinh Viên
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về Tên của Sinh Viên")]
        public static string LayTen(
            [ExcelDna.Integration.ExcelArgument(Description = "Original text")] string text)
        {
            return LayTenSV(text);
        }
        /// <summary>
        /// Trả về Họ của Sinh Viên
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về Họ của Sinh Viên")]
        public static string LayHo(
            [ExcelDna.Integration.ExcelArgument(Description = "Original text")] string text)
        {
            return LayHoSV(text);
        }
    }
}
