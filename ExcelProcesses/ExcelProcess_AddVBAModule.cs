using System;
using OfficeOpenXml;
using OfficeOpenXml.VBA;

namespace SPClient
{
    /**
     * Add vba module/class, replace if exists
     * */
    public class ExcelProcess_AddVBAModule : IExcelUpdateProcess
    {
        public string ErrorMessage { get; set; }

        public ExcelProcess_AddVBAModule(string name, string vba_content, ModuleType type)
        {
            Name = name;
            VBA_Content = vba_content;
            Type = type;
        }

        public bool IsConfigured { get; set; }

        public string VBA_Content { get; set; }
        public string Name { get; set; }
        public ModuleType Type { get; set; }

        public string Title { get { return "Add VBA Module"; } }
        public int MyProperty { get; set; }

        public enum ModuleType
        {
            MODULE,
            CLASS
        }

        public bool Execute(ExcelPackage p)
        {
            ExcelVBAModule module;
            if (p.Workbook.VbaProject.Modules.Exists(Name))
            {
                module = p.Workbook.VbaProject.Modules[Name];
            }
            else
            {
                if (Type == ModuleType.CLASS)
                {
                    module = p.Workbook.VbaProject.Modules.AddClass(Name, true);
                }
                else
                {
                    module = p.Workbook.VbaProject.Modules.AddModule(Name);
                }
            }
            module.Code = VBA_Content;
            return true;
        }

        public bool Configure()
        {
            throw new NotImplementedException();
        }
    }
}
