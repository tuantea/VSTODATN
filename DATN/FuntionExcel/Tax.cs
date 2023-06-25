using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DATN.FuntionExcel
{
    internal class Tax
     {
            public double SalaryRange;
            public double SalaryRate;
            /// <summary>
            ///     Bảng thuế suất biểu luỹ tiến từng phần, theo thông tư 111/2013 của Bộ tài chính
            /// </summary>
            static Tax[] list_TT1112013;
            static Tax()
            {
                /// Tạo bảng thuế suất theo thông tư 111/2013.  7 bậc
                list_TT1112013 = new Tax[7 + 1]; //+1 là dummy, phải gán về 0
                list_TT1112013[0] = new Tax(); list_TT1112013[0].SalaryRange = 0; list_TT1112013[0].SalaryRate = 0;
                list_TT1112013[1] = new Tax(); list_TT1112013[1].SalaryRange = 60 * Math.Pow(10, 6); list_TT1112013[1].SalaryRate = 0.05;
                list_TT1112013[2] = new Tax(); list_TT1112013[2].SalaryRange = 120 * Math.Pow(10, 6); list_TT1112013[2].SalaryRate = 0.10;
                list_TT1112013[3] = new Tax(); list_TT1112013[3].SalaryRange = 216 * Math.Pow(10, 6); list_TT1112013[3].SalaryRate = 0.15;
                list_TT1112013[4] = new Tax(); list_TT1112013[4].SalaryRange = 384 * Math.Pow(10, 6); list_TT1112013[4].SalaryRate = 0.20;
                list_TT1112013[5] = new Tax(); list_TT1112013[5].SalaryRange = 624 * Math.Pow(10, 6); list_TT1112013[5].SalaryRate = 0.25;
                list_TT1112013[6] = new Tax(); list_TT1112013[6].SalaryRange = 960 * Math.Pow(10, 6); list_TT1112013[6].SalaryRate = 0.30;
                list_TT1112013[7] = new Tax(); list_TT1112013[7].SalaryRange = Math.Pow(10, 15); list_TT1112013[7].SalaryRate = 0.35;

                /// Tạo bảng thuế suất theo thông tư ...........
                /// 
            }

            /// <summary>
            ///     Trả về bảng thuế suất
            /// </summary>
            /// <param name="version">The version.</param>
            static public Tax[] SelectThueSuat(int version)
            {
                return list_TT1112013;
            }
        /// <summary>
        /// Thues the ca nhan.
        /// </summary>
        /// <param name="Salary">The salary.</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Tiền thuế cần nộp của cả năm", Category = "Financial")]
        public static double ThueThuNhapCaNhan(
            [ExcelDna.Integration.ExcelArgument(Description = "Phần lương chịu thuế, sau khi đã miễn trừ")] double Salary
            )
        {
            /// Bảng thuế suất được sử dụng trong các tính toán thuế. Mặc định là mới nhất.
            Tax[] list;
            list = SelectThueSuat(0);

            /// Phân thuế cần phải tìm
            double Tax = 0;
            double TaxPerStep;
            // Phân tiền lương còn phải đóng thuế ở bậc tiếp theo
            double SalaryRemain = Salary;
            for (int index = 1; index < list.Length; index++)
            {
                if (Salary > list[index].SalaryRange)
                {
                    //Thuế phải nộp = phần chênh giữa 2 bậc nhân với thuế suất
                    TaxPerStep = (list[index].SalaryRange - list[index - 1].SalaryRange) * list[index].SalaryRate;
                    Tax += TaxPerStep;
                }
                else
                {
                    //Thuế phải nộp = phần chệnh giữa 2 bậc nhân với thuế suất
                    TaxPerStep = (Salary - list[index - 1].SalaryRange) * list[index].SalaryRate;
                    Tax += TaxPerStep;
                    break;
                }
            };
            // Trả kết quả về hàm
            return Tax;
        }

        /// <summary>
        /// Thues the ca nhan.
        /// </summary>
        /// <param name="Salary">The salary.</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Chi tiết tiền thuế cần nộp của cả năm, chi tiết theo từng bậc", Category = "Financial")]
        public static double[] ThueThuNhapCaNhanTheoBac(
            [ExcelDna.Integration.ExcelArgument(Description = "Phần lương chịu thuế, sau khi đã miễn trừ")] double Salary
            )
        {
            /// Bảng thuế suất được sử dụng trong các tính toán thuế. Mặc định là mới nhất.
            Tax[] list;
            list = SelectThueSuat(0);

            /// Phân thuế cần phải tìm
            double[] TaxPerStep = new double[list.Length - 1];
            // Phân tiền lương còn phải đóng thuế ở bậc tiếp theo
            double SalaryRemain = Salary;
            for (int index = 1; index < list.Length; index++)
            {
                if (Salary > list[index].SalaryRange)
                {
                    //Thuế phải nộp = phần chênh giữa 2 bậc nhân với thuế suất
                    TaxPerStep[index - 1] = (list[index].SalaryRange - list[index - 1].SalaryRange) * list[index].SalaryRate;
                }
                else
                {
                    //Thuế phải nộp = phần chệnh giữa 2 bậc nhân với thuế suất
                    TaxPerStep[index - 1] = (Salary - list[index - 1].SalaryRange) * list[index].SalaryRate;

                    //Gan phần còn lại về 0
                    for (index++; index < list.Length; index++)
                    {
                        TaxPerStep[index - 1] = 0;
                    }
                    break;
                }
            };
            // Trả kết quả về hàm
            return TaxPerStep;
        }

        /// <summary>
        ///     Hiển thị bảng thuế suất
        /// </summary>
        /// <param name="Salary">The salary.</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Hiển thị bảng thuế suất", Category = "Financial")]
        public static object BangThueSuat()
        {
            /// Bảng thuế suất được sử dụng trong các tính toán thuế. Mặc định là mới nhất.
            Tax[] list;
            list = SelectThueSuat(0);

            /// Chuyển đổi về dạng mảng 2 chiều, trong đó loại đi Bậc 0
            var res = new object[list.Length - 1, 2];
            for (int index = 1; index < list.Length; index++)
            {
                res[index - 1, 0] = list[index].SalaryRange;
                res[index - 1, 1] = list[index].SalaryRate;
            }
            // Trả kết quả về hàm
            return res;
        }
    }
}


      




