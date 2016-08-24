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
        string Title { get; }
        bool Execute(ExcelPackage p);
        string ErrorMessage { get; set; }
        bool IsConfigured { get; set; }
        bool Configure();
    }
}
