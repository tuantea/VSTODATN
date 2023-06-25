using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using System.Linq;


namespace DATN.FuntionExcel
{
    public class 
        
        
        
        
        
        
        
        
        
        
        AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var functions = ExcelRegistration.GetExcelFunctions().ToList();
            if (HasNativeXMatch())
            {
                foreach (var func in functions)
                {
                    func.FunctionAttribute.Name = "HUST." + func.FunctionAttribute.Name;
                }
            }
            functions.RegisterFunctions();

            ///Cho phép hiển thị các dòng gợi ý hàm và gợi ý tham số của các thuộc tính
            ///<see cref="ExcelDna.Integration.ExcelFunctionAttribute"/> và <see cref="ExcelDna.Integration.ExcelArgumentAttribute"/>
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            //Gỡ bỏ
            IntelliSenseServer.Uninstall();
        }

        bool HasNativeXMatch()
        {
            int xlfXMatch = 620;
            var retval = XlCall.TryExcel(xlfXMatch, out var _, 1, 1);
            return (retval == XlCall.XlReturn.XlReturnSuccess);
        }
    }
}

