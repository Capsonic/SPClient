using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPClient
{
    public interface IExcelUpdateProcess
    {
        bool Execute(ExcelPackage p);
        string ErrorMessage { get; set; }
    }
}
