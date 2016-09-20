using System;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.VBA;
using System.Text;

namespace SPClient.ExcelProcesses
{
    public partial class ExcelProcess_AddVBAModule : Form, IExcelUpdateProcess
    {
        public ExcelProcess_AddVBAModule()
        {
            InitializeComponent();
        }

        public string ErrorMessage { get; set; }

        public bool IsConfigured { get; set; }

        public string Title { get { return "Update VBA code."; } }

        public bool Configure()
        {
            ShowDialog();
            return IsConfigured;
        }

        public bool Execute(ExcelPackage p)
        {
            if (p.Workbook.VbaProject.Modules.Exists(txtName.Text.Trim()))
            {
                var module = p.Workbook.VbaProject.Modules[txtName.Text.Trim()];
                module.Code = txtContent.Text;
            }
            else
            {
                ErrorMessage = "Non-existent module: " + txtName.Text.Trim();
                return false;
            }
            return true;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            IsConfigured = true;
            Hide();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            IsConfigured = false;
            Hide();
        }
    }
}
