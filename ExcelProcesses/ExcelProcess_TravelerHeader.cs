using System;
using System.Windows.Forms;
using OfficeOpenXml;

namespace SPClient.ExcelProcesses
{
    public partial class ExcelProcess_TravelerHeader : Form, IExcelUpdateProcess
    {
        public ExcelProcess_TravelerHeader()
        {
            InitializeComponent();
        }

        public string ErrorMessage { get; set; }

        public bool IsConfigured { get; set; }

        public string Title { get { return "Traveler Header Update"; } }

        public bool Configure()
        {
            ShowDialog();
            return IsConfigured;
        }

        public bool Execute(ExcelPackage p)
        {
            ErrorMessage = "";

            //1. Verify workbook has a log sheet.
            bool bLogExists = false;
            foreach (var worksheet in p.Workbook.Worksheets)
            {
                if (worksheet.Name.Trim().ToLower() == "log")
                {
                    bLogExists = true;
                    break;
                }
            }
            if (!bLogExists)
            {
                ErrorMessage = "Log sheet not found.";
                return false;
            }

            //2. Verify workbook has at least two sheets.
            if (p.Workbook.Worksheets.Count <= 1)
            {
                ErrorMessage = "At least two sheets must exist.";
                return false;
            }

            //3. Try to get sheet with a different name than "log".
            ExcelWorksheet ws = null;
            foreach (var worksheet in p.Workbook.Worksheets)
            {
                if (worksheet.Name.Trim().ToLower() != "log")
                {
                    ws = worksheet;
                    break;
                }
            }
            if (ws == null)
            {
                ErrorMessage = "Could not found a sheet with a different name than \"log\"";
                return false;
            }

            //Updates:
            ws.Cells["N2"].Style.Numberformat.Format = ""; //Empty == General format
            ws.Cells["N4"].Style.Numberformat.Format = "";
            ws.Cells["N5"].Style.Numberformat.Format = "";
            ws.Row(1).Height = 55;
            ws.Row(2).Height = 70;

            //Part Number:
            ws.Cells["D1"].Style.Numberformat.Format = "";
            ws.Cells["D1"].Formula = "=N2";
            ws.Cells["D1"].Style.Font.Size = 45;
            ws.Cells["D1"].Style.Font.Bold = true;
            ws.Cells["D1"].Style.Font.Name = "Geneva";

            //Part Number Bar Code:
            ws.Cells["D2"].Style.Numberformat.Format = "";
            ws.Cells["D2"].Formula = "=\"*\" & N2 & \"*\"";
            ws.Cells["D2"].Style.Font.Bold = false;
            ws.Cells["D2"].Style.Font.Name = "Free 3 of 9 Extended";
            ws.Cells["D2"].Style.Font.Size = 80;

            //# de M.O value
            ws.Cells["D4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

            //# de M.O. bar code
            ws.Cells["E5"].Value = null;
            ws.Cells["D5:E5"].Merge = true;
            ws.Cells["D5:E5"].Style.Numberformat.Format = "";
            ws.Cells["D5:E5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["D5:E5"].Formula = "=\"*\" & D4 & \"*\"";
            ws.Cells["D5:E5"].Style.Font.Size = 80;
            ws.Cells["D5:E5"].Style.Font.Name = "Free 3 of 9 Extended";

            //MO Ln Num Title
            ws.Cells["F3"].Style.Numberformat.Format = "";
            ws.Cells["F3"].Value = "MO Ln Num.";
            ws.Cells["F3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["F3"].Style.Font.Size = 22;
            ws.Cells["F3"].Style.Font.Bold = true;

            //MO Ln Num value
            ws.Cells["F4:G4"].Merge = true;
            ws.Cells["F4:G4"].Style.Numberformat.Format = "@";
            ws.Cells["F4:G4"].Style.Font.Size = 48;
            ws.Cells["F4:G4"].Style.Font.Name = "Geneva";
            ws.Cells["F4:G4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells["F4:G4"].Style.Font.Bold = true;


            //MO Ln Num Bar Code
            ws.Cells["F5"].Style.Numberformat.Format = "";
            ws.Cells["F5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["F5"].Formula = "=\"*\" & F4 & \"*\"";
            ws.Cells["F5"].Style.Font.Name = "Free 3 of 9 Extended";
            ws.Cells["F5"].Style.Font.Size = 80;


            //Cantidad Pzas
            ws.Cells["G5"].Value = null;
            ws.Cells["H3:I3"].Merge = true;
            ws.Cells["H3:I3"].Style.Numberformat.Format = "";
            ws.Cells["H3:I3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells["H3:I3"].Value = "Cant. Pzas.";
            ws.Cells["H3:I3"].Style.Font.Size = 22;
            ws.Cells["H3:I3"].Style.Font.Bold = true;


            //# serie title
            ws.Cells["J5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            //# SERIE BAR CODE
            ws.Cells["K5"].Style.Font.Size = 65;
            ws.Cells["K5"].Style.Font.Name = "Free 3 of 9 Extended";
            ws.Cells["K5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["K5"].Style.Font.Bold = false;

            //Description
            ws.Cells["F2:J2"].Merge = true;
            ws.Cells["F2:J2"].Formula = "=N3";
            ws.Cells["F2:J2"].Style.Font.Size = 24;

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
