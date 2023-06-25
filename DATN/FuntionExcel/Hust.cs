using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DATN.FuntionExcel
{
    internal class Hust
    {
        /// <summary>
        ///         Hàm Excel EmailSinhVien
        /// </summary>
        /// <param name="HoVaTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <param name="MaSoSinhVien">Mã số SV do trường cấp. Ví dụ 20002987</param>
        /// <returns></returns>
        /// <remarks> ExcelDna.Integration.ExcelFunction(Name = ...)  sẽ qui định tên hàm để gọi ra trong Excel </remarks>
        [ExcelDna.Integration.ExcelFunction(Description = "Tính địa chỉ email HUST của sinh viên, giảng viên dựa theo tên và mã số sinh viên", Name = "EmailSinhVien", Category = "Text")]
        public static object StudentEmail(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen,
            [ExcelDna.Integration.ExcelArgument(Description = "Mã số SV do trường cấp. Bỏ trống nếu là giảng viên. Ví dụ 20002987.")] string MaSoSinhVien)
        {
            string TenKhongDau = StringFormattter.RemoveAccent(HoVaTen);
            string Email;
            if (MaSoSinhVien.Length == 8)
            {
                string ChuCaiDau = StringFormattter.LayCacChuCaiDau(TenKhongDau);
                Email = StringFormattter.LayTenSV(TenKhongDau) + "." + ChuCaiDau.Substring(0, ChuCaiDau.Length - 1) + MaSoSinhVien.Substring(2, MaSoSinhVien.Length - 2) + "@sis.hust.edu.vn";
            }
            else if (MaSoSinhVien.Length == 0)
            {
                Email = StringFormattter.LayTenSV(TenKhongDau) + "." + StringFormattter.LayHoSV(TenKhongDau) + StringFormattter.LayDemSV(TenKhongDau).Trim() + "@hust.edu.vn";
            }
            else
            {
                Email = "Mã số không hợp lệ";
            }
            return Email;
        }
        /// <summary>
        ///         KipThi function, return the starting time of the exam
        /// </summary>
        /// <param name="Kip">Kíp thi từ 1 đến 4</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Return starting time of the exam", Name = "KipThi")]
        public static object KipThi(
            [ExcelDna.Integration.ExcelArgument(Description = "Số thứ tự kíp thi, là 1 | 2 | 3 | 4. Bách Khoa Hà Nội chỉ có 4 buổi thi mỗi ngày.")] int Kip
            )
        {
            string startingTime;
            switch (Kip)
            {
                case 1: startingTime = "7:00"; break;
                case 2: startingTime = "9:30"; break;
                case 3: startingTime = "12:30"; break;
                case 4: startingTime = "15:00"; break;
                default: startingTime = "Invalid"; break;
            }
            return startingTime;
        }
    }
}
