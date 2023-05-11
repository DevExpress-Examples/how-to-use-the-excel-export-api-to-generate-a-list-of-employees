using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace XLExportExample {
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        List<EmployeeData> employees = EmployeesRepository.CreateEmployees();
        List<string> departments = EmployeesRepository.CreateDepartments();
        XlCellFormatting headerRowFormatting;
        XlCellFormatting evenRowFormatting;
        XlCellFormatting oddRowFormatting;

        public Form1() {
            InitializeComponent();
            InitializeFormatting();
        }

        void InitializeFormatting() {
            // Specify formatting settings for the even rows.
            evenRowFormatting = new XlCellFormatting();
            evenRowFormatting.Font = new XlFont();
            evenRowFormatting.Font.Name = "Century Gothic";
            evenRowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;
            evenRowFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center);

            // Specify formatting settings for the odd rows.
            oddRowFormatting = new XlCellFormatting();
            oddRowFormatting.CopyFrom(evenRowFormatting);
            oddRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, -0.15));

            // Specify formatting settings for the header row.
            headerRowFormatting = new XlCellFormatting();
            headerRowFormatting.CopyFrom(evenRowFormatting);
            headerRowFormatting.Font.Bold = true;
            headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
            headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));
            headerRowFormatting.Border = new XlBorder();
            headerRowFormatting.Border.TopColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0);
            headerRowFormatting.Border.TopLineStyle = XlBorderLineStyle.Medium;
            headerRowFormatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0);
            headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;
        }

        // Export the document to XLSX format.
        void btnExportToXLSX_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xlsx))
                ShowFile(fileName);
        }

        // Export the document to XLS format.
        void btnExportToXLS_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xls))
                ShowFile(fileName);
        }

        // Export the document to CSV format.
        void btnExportToCSV_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Csv))
                ShowFile(fileName);
        }

        string GetSaveFileName(string filter, string defaultName) {
            saveFileDialog1.Filter = filter;
            saveFileDialog1.FileName = defaultName;
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return null;
            return saveFileDialog1.FileName;
        }

        void ShowFile(string fileName) {
            if (!File.Exists(fileName))
                return;
            DialogResult dResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName),
                this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dResult == DialogResult.Yes)
                Process.Start(fileName);
        }

        bool ExportToFile(string fileName, XlDocumentFormat documentFormat) {
            try {
                using (FileStream stream = new FileStream(fileName, FileMode.Create)) {
                    // Create an exporter with the specified formula parser.
                    IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());
                    // Create a new document and begin to write it to the specified stream. 
                    using (IXlDocument document = exporter.CreateDocument(stream)) {
                        // Generate the document content. 
                        GenerateDocument(document);
                    }
                }
                return true;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        void GenerateDocument(IXlDocument document) {
            // Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture;

            // Add a new worksheet to the document.
            using (IXlSheet sheet = document.CreateSheet()) {
                // Specify the worksheet name.
                sheet.Name = "Employees";

                // Specify print settings for the worksheet.
                SetupPageParameters(sheet);

                // Generate worksheet columns.
                GenerateColumns(sheet);

                // Add the title to the documents exported to the XLSX and XLS formats.  
                if (document.Options.DocumentFormat != XlDocumentFormat.Csv)
                    GenerateTitle(sheet);

                // Create the header row.
                GenerateHeaderRow(sheet);

                int firstDataRowIndex = sheet.CurrentRowIndex;

                // Create the data rows.
                for (int i = 0; i < employees.Count; i++)
                    GenerateDataRow(sheet, employees[i], (i + 1) == employees.Count);

                // Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange;

                // Create data validation criteria for the documents exported to the XLSX and XLS formats.
                if(document.Options.DocumentFormat != XlDocumentFormat.Csv)
                    GenerateDataValidations(sheet, firstDataRowIndex);
            }

            // Create the hidden worksheet containing source data for the data validation drop-down list.
            if (document.Options.DocumentFormat != XlDocumentFormat.Csv) {
                using (IXlSheet sheet = document.CreateSheet()) {
                    sheet.Name = "Departments";
                    sheet.VisibleState = XlSheetVisibleState.Hidden;

                    foreach (string department in departments) {
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell())
                                cell.Value = department;
                        }
                    }
                }
            }

        }

        void GenerateColumns(IXlSheet sheet) {
            // Create the "Employee ID" column and set its width.
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 110;

            // Create the "Employee Name" column and set its width.
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 200;

            XlNumberFormat numberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";

            // Create the "Salary" and "Bonus" columns and set the specific number format for their cells.
            for (int i = 0; i < 2; i++) {
                using (IXlColumn column = sheet.CreateColumn()) {
                    column.WidthInPixels = 100;
                    column.ApplyFormatting(numberFormat);
                }
            }

            // Create the "Department" column and set its width.
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 140;
        }

        void GenerateTitle(IXlSheet sheet) {
            // Specify formatting settings for the document title.
            XlCellFormatting formatting = new XlCellFormatting();
            formatting.Font = new XlFont();
            formatting.Font.Name = "Calibri Light";
            formatting.Font.SchemeStyle = XlFontSchemeStyles.None;
            formatting.Font.Size = 24;
            formatting.Font.Color = XlColor.FromTheme(XlThemeColor.Dark1, 0.35);
            formatting.Border = new XlBorder();
            formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.35);
            formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;

            // Add the document title.
            using (IXlRow row = sheet.CreateRow()) {
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = "LIST OF EMPLOYEES";
                    cell.Formatting = formatting;
                }
                // Create four empty cells with the title formatting.
                row.BlankCells(4, formatting);
            }

            // Skip one row before starting to generate data rows.
            sheet.SkipRows(1);
        }

        void GenerateHeaderRow(IXlSheet sheet) {
            string[] columnNames = new string[] { "Employee ID", "Employee Name", "Salary", "Bonus", "Department" };
            // Create the header row and set its height.
            using (IXlRow row = sheet.CreateRow()) {
                row.HeightInPixels = 28;

                // Create required cells in the header row and apply specific formatting settings to them. 
                foreach(string columnName in columnNames) {
                    using (IXlCell cell = row.CreateCell()) {
                        cell.Value = columnName;
                        cell.ApplyFormatting(headerRowFormatting);
                    }
                }
            }
        }

        void GenerateDataRow(IXlSheet sheet, EmployeeData employee, bool isLastRow) {
            // Create the data row to display the employee's information.
            using (IXlRow row = sheet.CreateRow()) {
                row.HeightInPixels = 28;

                // Specify formatting settings to be applied to the data rows to shade alternate rows. 
                XlCellFormatting formatting = new XlCellFormatting();
                formatting.CopyFrom((row.RowIndex % 2 == 0) ? evenRowFormatting : oddRowFormatting);
                // Set the bottom border for the last data row.
                if (isLastRow) {
                    formatting.Border = new XlBorder();
                    formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0);
                    formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;
                }

                // Create the cell containing the employee's ID. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = employee.Id;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the employee's name.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = employee.Name;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the employee's salary.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = employee.Salary;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing information about bonuses.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = employee.Bonus;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the department name.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = employee.Department;
                    cell.ApplyFormatting(formatting);
                }
            }
        }

        void SetupPageParameters(IXlSheet sheet) {
            // Specify the header and footer for the odd-numbered pages.
            sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR("NorthWind Inc.", null, XlHeaderFooter.Date);
            sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR("List of employees", null, XlHeaderFooter.PageNumber + " of " + XlHeaderFooter.PageTotal);

            // Specify page margins.
            sheet.PageMargins = new XlPageMargins();
            sheet.PageMargins.PageUnits = XlPageUnits.Centimeters;
            sheet.PageMargins.Left = 2.0;
            sheet.PageMargins.Right = 1.0;
            sheet.PageMargins.Top = 1.4;
            sheet.PageMargins.Bottom = 1.4;
            sheet.PageMargins.Header = 0.7;
            sheet.PageMargins.Footer = 0.7;

            // Specify page settings.
            sheet.PageSetup = new XlPageSetup();
            // Select the paper size.
            sheet.PageSetup.PaperKind = PaperKind.A4;
            // Scale the print area to fit to one page wide.
            sheet.PageSetup.FitToPage = true;
            sheet.PageSetup.FitToWidth = 1;
            sheet.PageSetup.FitToHeight = 0;
        }

        void GenerateDataValidations(IXlSheet sheet, int firstDataRowIndex) {
            // Restrict data entry in the "Employee ID" column using criteria calculated by a worksheet formula (Employee ID must be a 5-digit number).
            XlDataValidation validation = new XlDataValidation();
            validation.Ranges.Add(XlCellRange.FromLTRB(0, firstDataRowIndex, 0, sheet.CurrentRowIndex - 1));
            validation.Type = XlDataValidationType.Custom;
            validation.Criteria1 = string.Format("=AND(ISNUMBER(A{0}),LEN(A{0})=5)", firstDataRowIndex + 1);
            validation.InputPrompt = "Please enter a 5-digit number.";
            validation.PromptTitle = "Employee ID";
            sheet.DataValidations.Add(validation);

            // Restrict data entry in the "Salary" column to a whole number from 600 to 2000.
            validation = new XlDataValidation();
            validation.Ranges.Add(XlCellRange.FromLTRB(2, firstDataRowIndex, 2, sheet.CurrentRowIndex - 1));
            validation.Type = XlDataValidationType.Whole;
            validation.Operator = XlDataValidationOperator.Between;
            validation.Criteria1 = 600;
            validation.Criteria2 = 2000;
            // Specify the error message.
            validation.ErrorMessage = "Salary must be greater than $600 and less than $2000.";
            validation.ErrorTitle = "Warning";
            validation.ErrorStyle = XlDataValidationErrorStyle.Warning;
            // Specify the input message.
            validation.InputPrompt = "Please enter a whole number in the range 600...2000.";
            validation.PromptTitle = "Salary";
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            sheet.DataValidations.Add(validation);

            // Restrict data entry in the "Bonus" column to a decimal number within the specified limits (bonus cannot be greater than 10% of the salary.)
            validation = new XlDataValidation();
            validation.Ranges.Add(XlCellRange.FromLTRB(3, firstDataRowIndex, 3, sheet.CurrentRowIndex - 1));
            validation.Type = XlDataValidationType.Whole;
            validation.Operator = XlDataValidationOperator.Between;
            validation.Criteria1 = 0;
            validation.Criteria2 = string.Format("=C{0}*0.1", firstDataRowIndex + 1);
            // Specify the error message.
            validation.ErrorMessage = "Bonus cannot be greater than 10% of the salary.";
            validation.ErrorTitle = "Information";
            validation.ErrorStyle = XlDataValidationErrorStyle.Information;
            validation.ShowErrorMessage = true;
            sheet.DataValidations.Add(validation);

            // Restrict data entry in the "Department" column to values in a drop-down list obtained from the cell range in the hidden "Departments" worksheet.
            validation = new XlDataValidation();
            validation.Ranges.Add(XlCellRange.FromLTRB(4, firstDataRowIndex, 4, sheet.CurrentRowIndex - 1));
            validation.Type = XlDataValidationType.List;
            XlCellRange sourceRange = XlCellRange.FromLTRB(0, 0, 0, departments.Count - 1).AsAbsolute();
            sourceRange.SheetName = "Departments";
            validation.Criteria1 = sourceRange;
            sheet.DataValidations.Add(validation);
        }
    }
}
