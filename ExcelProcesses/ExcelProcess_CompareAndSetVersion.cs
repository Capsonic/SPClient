using System;
using System.Windows.Forms;
using OfficeOpenXml;

namespace SPClient
{
    public partial class frmCompareAndSetVersion : Form, IExcelUpdateProcess
    {
        public frmCompareAndSetVersion()
        {
            InitializeComponent();
        }

        public string ErrorMessage { get; set; }

        public bool IsConfigured { get; set; }

        public string Title { get { return "Compare and Set Version"; } }

        public bool Execute(ExcelPackage p)
        {
            ErrorMessage = "";
            bool bBEXU_VersionExists = false;
            int? iVersionThan = null;
            int? iCurrentVersion = null;
            string currentVersion = (string)p.Workbook.Properties.GetCustomPropertyValue("BEXU_Version");
            if (currentVersion != null)
            {
                bBEXU_VersionExists = true;
            }

            switch (cboComparatorMethod.Text)
            {
                case "Without BEXU_Version":
                    if (bBEXU_VersionExists)
                    {
                        ErrorMessage = "Not the version specified.";
                        return false;
                    }
                    break;
                case "Equals To":
                    if (currentVersion == null)
                    {
                        ErrorMessage = "This file has no version";
                        return false;
                    }
                    if (currentVersion.ToLower().Trim() != txtCompareVersion.Text.ToLower().Trim())
                    {
                        ErrorMessage = "Not the version specified.";
                        return false;
                    }
                    break;
                case "Less Than":
                    if (currentVersion == null)
                    {
                        ErrorMessage = "This file has no version";
                        return false;
                    }
                    iCurrentVersion = int.Parse(currentVersion);
                    iVersionThan = txtCompareVersion.Text != "" ? int.Parse(txtCompareVersion.Text) : (int?)null;
                    if (iCurrentVersion >= iVersionThan)
                    {
                        ErrorMessage = "Not the version specified.";
                        return false;
                    }
                    break;
                case "Greater Than":
                    if (currentVersion == null)
                    {
                        ErrorMessage = "This file has no version";
                        return false;
                    }
                    iCurrentVersion = int.Parse(currentVersion);
                    iVersionThan = txtCompareVersion.Text != "" ? int.Parse(txtCompareVersion.Text) : (int?)null;
                    if (iCurrentVersion <= iVersionThan)
                    {
                        ErrorMessage = "Not the version specified.";
                        return false;
                    }
                    break;
                default:
                    throw new Exception("Invalid Comparator-Method for version.");
            }

            p.Workbook.Properties.SetCustomPropertyValue("BEXU_Version", txtNewVersion.Text);
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

        public bool Configure()
        {
            ShowDialog();
            return IsConfigured;
        }
    }
}
