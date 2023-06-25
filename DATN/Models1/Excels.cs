using Microsoft.Office.Interop.Excel;
using Microsoft.Object.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
namespace DATN.Models{
    public static class Excel
    {
        public static object Data{get;set;}
        public static string PathToM7DOpen
        {
            get{ return "C:\\Users\\nguye\\OneDrive\\Documents\\datn.xlsx";}
        }
        public static string PathToM7DModel
        {
            get { return "C:\\Users\\nguye\\OneDrive\\Documents\\datn.xlsx";}
        }
        public static DateTime date=DateTime.Today;
    }
}