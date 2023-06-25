using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DATN.FuntionExcel
{
    internal class Finance
    {
        public static double OptimizeInterestRate1(double principal, int loanTerm, double targetProfit)
        {
            double minInterestRate = 0; // Lãi suất tối thiểu
            double maxInterestRate = 100; // Lãi suất tối đa
            double step = 0.01; // Bước nhảy

            double optimalInterestRate = minInterestRate;
            double optimalProfit = CalculateProfit(principal, loanTerm, optimalInterestRate);

            // Tìm lãi suất tối ưu dựa trên lợi nhuận mục tiêu
            for (double interestRate = minInterestRate; interestRate <= maxInterestRate; interestRate += step)
            {
                double profit = CalculateProfit(principal, loanTerm, interestRate);
                if (profit >= targetProfit)
                {
                    optimalInterestRate = interestRate;
                    break;
                }
                else if (profit > optimalProfit)
                {
                    optimalInterestRate = interestRate;
                    optimalProfit = profit;
                }
            }

            return optimalInterestRate;
        }

        public static double CalculateProfit(double principal, int loanTerm, double interestRate)
        {
            double monthlyInterestRate = interestRate / 12 / 100; // Lãi suất hàng tháng

            // Tính lãi suất hàng tháng
            double monthlyPayment = principal * monthlyInterestRate / (1 - Math.Pow(1 + monthlyInterestRate, -loanTerm));

            // Tính tổng lợi nhuận sau kỳ trả
            double totalProfit = monthlyPayment * loanTerm - principal;

            return totalProfit;
        }
        /// <summary>
        ///         KipThi function, trả về lãi suất tối ưu
        /// <param name="Kip">Kíp thi từ 1 đến 4</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về lãi suất tối ưu", Name = "OptimizeInterestRate")]
        public static object OptimizeInterestRate(
            [ExcelDna.Integration.ExcelArgument(Description = "Số tiền vay")] double principal,
            [ExcelDna.Integration.ExcelArgument(Description = "Số kỳ trả (tháng)")] int loanTerm,
            [ExcelDna.Integration.ExcelArgument(Description = "Lợi nhuận mục tiêu")] double targetProfit
            )
        {
            return OptimizeInterestRate1(principal, loanTerm, targetProfit);
        }
    }
}
