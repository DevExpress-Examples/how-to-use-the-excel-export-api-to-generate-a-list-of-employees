Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing.Printing
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Namespace XLExportExample

    Public Partial Class Form1
        Inherits DevExpress.XtraEditors.XtraForm

        Private employees As List(Of EmployeeData) = CreateEmployees()

        Private departments As List(Of String) = CreateDepartments()

        Private headerRowFormatting As XlCellFormatting

        Private evenRowFormatting As XlCellFormatting

        Private oddRowFormatting As XlCellFormatting

        Public Sub New()
            InitializeComponent()
            InitializeFormatting()
        End Sub

        Private Sub InitializeFormatting()
            ' Specify formatting settings for the even rows.
            evenRowFormatting = New XlCellFormatting()
            evenRowFormatting.Font = New XlFont()
            evenRowFormatting.Font.Name = "Century Gothic"
            evenRowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
            evenRowFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center)
            ' Specify formatting settings for the odd rows.
            oddRowFormatting = New XlCellFormatting()
            oddRowFormatting.CopyFrom(evenRowFormatting)
            oddRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, -0.15))
            ' Specify formatting settings for the header row.
            headerRowFormatting = New XlCellFormatting()
            headerRowFormatting.CopyFrom(evenRowFormatting)
            headerRowFormatting.Font.Bold = True
            headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
            headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))
            headerRowFormatting.Border = New XlBorder()
            headerRowFormatting.Border.TopColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0)
            headerRowFormatting.Border.TopLineStyle = XlBorderLineStyle.Medium
            headerRowFormatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0)
            headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Medium
        End Sub

        ' Export the document to XLSX format.
        Private Sub btnExportToXLSX_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLSX.Click
            Dim fileName As String = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx")
            If String.IsNullOrEmpty(fileName) Then Return
            If ExportToFile(fileName, XlDocumentFormat.Xlsx) Then ShowFile(fileName)
        End Sub

        ' Export the document to XLS format.
        Private Sub btnExportToXLS_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLS.Click
            Dim fileName As String = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls")
            If String.IsNullOrEmpty(fileName) Then Return
            If ExportToFile(fileName, XlDocumentFormat.Xls) Then ShowFile(fileName)
        End Sub

        ' Export the document to CSV format.
        Private Sub btnExportToCSV_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToCSV.Click
            Dim fileName As String = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv")
            If String.IsNullOrEmpty(fileName) Then Return
            If ExportToFile(fileName, XlDocumentFormat.Csv) Then ShowFile(fileName)
        End Sub

        Private Function GetSaveFileName(ByVal filter As String, ByVal defaultName As String) As String
            saveFileDialog1.Filter = filter
            saveFileDialog1.FileName = defaultName
            If saveFileDialog1.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return saveFileDialog1.FileName
        End Function

        Private Sub ShowFile(ByVal fileName As String)
            If Not File.Exists(fileName) Then Return
            Dim dResult As DialogResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dResult = DialogResult.Yes Then Call Process.Start(fileName)
        End Sub

        Private Function ExportToFile(ByVal fileName As String, ByVal documentFormat As XlDocumentFormat) As Boolean
            Try
                Using stream As FileStream = New FileStream(fileName, FileMode.Create)
                    ' Create an exporter with the specified formula parser.
                    Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())
                    ' Create a new document and begin to write it to the specified stream. 
                    Using document As IXlDocument = exporter.CreateDocument(stream)
                        ' Generate the document content. 
                        GenerateDocument(document)
                    End Using
                End Using

                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        Private Sub GenerateDocument(ByVal document As IXlDocument)
            ' Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture
            ' Add a new worksheet to the document.
            Using sheet As IXlSheet = document.CreateSheet()
                ' Specify the worksheet name.
                sheet.Name = "Employees"
                ' Specify print settings for the worksheet.
                SetupPageParameters(sheet)
                ' Generate worksheet columns.
                GenerateColumns(sheet)
                ' Add the title to the documents exported to the XLSX and XLS formats.  
                If document.Options.DocumentFormat <> XlDocumentFormat.Csv Then GenerateTitle(sheet)
                ' Create the header row.
                GenerateHeaderRow(sheet)
                Dim firstDataRowIndex As Integer = sheet.CurrentRowIndex
                ' Create the data rows.
                For i As Integer = 0 To employees.Count - 1
                    GenerateDataRow(sheet, employees(i), i + 1 = employees.Count)
                Next

                ' Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange
                ' Create data validation criteria for the documents exported to the XLSX and XLS formats.
                If document.Options.DocumentFormat <> XlDocumentFormat.Csv Then GenerateDataValidations(sheet, firstDataRowIndex)
            End Using

            ' Create the hidden worksheet containing source data for the data validation drop-down list.
            If document.Options.DocumentFormat <> XlDocumentFormat.Csv Then
                Using sheet As IXlSheet = document.CreateSheet()
                    sheet.Name = "Departments"
                    sheet.VisibleState = XlSheetVisibleState.Hidden
                    For Each department As String In departments
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = department
                            End Using
                        End Using
                    Next
                End Using
            End If
        End Sub

        Private Sub GenerateColumns(ByVal sheet As IXlSheet)
            ' Create the "Employee ID" column and set its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 110
            End Using

            ' Create the "Employee Name" column and set its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 200
            End Using

            Dim numberFormat As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
            ' Create the "Salary" and "Bonus" columns and set the specific number format for their cells.
            For i As Integer = 0 To 2 - 1
                Using column As IXlColumn = sheet.CreateColumn()
                    column.WidthInPixels = 100
                    column.ApplyFormatting(numberFormat)
                End Using
            Next

            ' Create the "Department" column and set its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 140
            End Using
        End Sub

        Private Sub GenerateTitle(ByVal sheet As IXlSheet)
            ' Specify formatting settings for the document title.
            Dim formatting As XlCellFormatting = New XlCellFormatting()
            formatting.Font = New XlFont()
            formatting.Font.Name = "Calibri Light"
            formatting.Font.SchemeStyle = XlFontSchemeStyles.None
            formatting.Font.Size = 24
            formatting.Font.Color = XlColor.FromTheme(XlThemeColor.Dark1, 0.35)
            formatting.Border = New XlBorder()
            formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.35)
            formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium
            ' Add the document title.
            Using row As IXlRow = sheet.CreateRow()
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "LIST OF EMPLOYEES"
                    cell.Formatting = formatting
                End Using

                ' Create four empty cells with the title formatting.
                row.BlankCells(4, formatting)
            End Using

            ' Skip one row before starting to generate data rows.
            sheet.SkipRows(1)
        End Sub

        Private Sub GenerateHeaderRow(ByVal sheet As IXlSheet)
            Dim columnNames As String() = New String() {"Employee ID", "Employee Name", "Salary", "Bonus", "Department"}
            ' Create the header row and set its height.
            Using row As IXlRow = sheet.CreateRow()
                row.HeightInPixels = 28
                ' Create required cells in the header row and apply specific formatting settings to them. 
                For Each columnName As String In columnNames
                    Using cell As IXlCell = row.CreateCell()
                        cell.Value = columnName
                        cell.ApplyFormatting(headerRowFormatting)
                    End Using
                Next
            End Using
        End Sub

        Private Sub GenerateDataRow(ByVal sheet As IXlSheet, ByVal employee As EmployeeData, ByVal isLastRow As Boolean)
            ' Create the data row to display the employee's information.
            Using row As IXlRow = sheet.CreateRow()
                row.HeightInPixels = 28
                ' Specify formatting settings to be applied to the data rows to shade alternate rows. 
                Dim formatting As XlCellFormatting = New XlCellFormatting()
                formatting.CopyFrom(If(row.RowIndex Mod 2 = 0, evenRowFormatting, oddRowFormatting))
                ' Set the bottom border for the last data row.
                If isLastRow Then
                    formatting.Border = New XlBorder()
                    formatting.Border.BottomColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.0)
                    formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium
                End If

                ' Create the cell containing the employee's ID. 
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = employee.Id
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the employee's name.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = employee.Name
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the employee's salary.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = employee.Salary
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing information about bonuses.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = employee.Bonus
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the department name.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = employee.Department
                    cell.ApplyFormatting(formatting)
                End Using
            End Using
        End Sub

        Private Sub SetupPageParameters(ByVal sheet As IXlSheet)
            ' Specify the header and footer for the odd-numbered pages.
            sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR("NorthWind Inc.", Nothing, XlHeaderFooter.Date)
            sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR("List of employees", Nothing, XlHeaderFooter.PageNumber & " of " & XlHeaderFooter.PageTotal)
            ' Specify page margins.
            sheet.PageMargins = New XlPageMargins()
            sheet.PageMargins.PageUnits = XlPageUnits.Centimeters
            sheet.PageMargins.Left = 2.0
            sheet.PageMargins.Right = 1.0
            sheet.PageMargins.Top = 1.4
            sheet.PageMargins.Bottom = 1.4
            sheet.PageMargins.Header = 0.7
            sheet.PageMargins.Footer = 0.7
            ' Specify page settings.
            sheet.PageSetup = New XlPageSetup()
            ' Select the paper size.
            sheet.PageSetup.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.A4
            ' Scale the print area to fit to one page wide.
            sheet.PageSetup.FitToPage = True
            sheet.PageSetup.FitToWidth = 1
            sheet.PageSetup.FitToHeight = 0
        End Sub

        Private Sub GenerateDataValidations(ByVal sheet As IXlSheet, ByVal firstDataRowIndex As Integer)
            ' Restrict data entry in the "Employee ID" column using criteria calculated by a worksheet formula (Employee ID must be a 5-digit number).
            Dim validation As XlDataValidation = New XlDataValidation()
            validation.Ranges.Add(XlCellRange.FromLTRB(0, firstDataRowIndex, 0, sheet.CurrentRowIndex - 1))
            validation.Type = XlDataValidationType.Custom
            validation.Criteria1 = String.Format("=AND(ISNUMBER(A{0}),LEN(A{0})=5)", firstDataRowIndex + 1)
            validation.InputPrompt = "Please enter a 5-digit number."
            validation.PromptTitle = "Employee ID"
            sheet.DataValidations.Add(validation)
            ' Restrict data entry in the "Salary" column to a whole number from 600 to 2000.
            validation = New XlDataValidation()
            validation.Ranges.Add(XlCellRange.FromLTRB(2, firstDataRowIndex, 2, sheet.CurrentRowIndex - 1))
            validation.Type = XlDataValidationType.Whole
            validation.Operator = XlDataValidationOperator.Between
            validation.Criteria1 = 600
            validation.Criteria2 = 2000
            ' Specify the error message.
            validation.ErrorMessage = "Salary must be greater than $600 and less than $2000."
            validation.ErrorTitle = "Warning"
            validation.ErrorStyle = XlDataValidationErrorStyle.Warning
            ' Specify the input message.
            validation.InputPrompt = "Please enter a whole number in the range 600...2000."
            validation.PromptTitle = "Salary"
            validation.ShowErrorMessage = True
            validation.ShowInputMessage = True
            sheet.DataValidations.Add(validation)
            ' Restrict data entry in the "Bonus" column to a decimal number within the specified limits (bonus cannot be greater than 10% of the salary.)
            validation = New XlDataValidation()
            validation.Ranges.Add(XlCellRange.FromLTRB(3, firstDataRowIndex, 3, sheet.CurrentRowIndex - 1))
            validation.Type = XlDataValidationType.Whole
            validation.Operator = XlDataValidationOperator.Between
            validation.Criteria1 = 0
            validation.Criteria2 = String.Format("=C{0}*0.1", firstDataRowIndex + 1)
            ' Specify the error message.
            validation.ErrorMessage = "Bonus cannot be greater than 10% of the salary."
            validation.ErrorTitle = "Information"
            validation.ErrorStyle = XlDataValidationErrorStyle.Information
            validation.ShowErrorMessage = True
            sheet.DataValidations.Add(validation)
            ' Restrict data entry in the "Department" column to values in a drop-down list obtained from the cell range in the hidden "Departments" worksheet.
            validation = New XlDataValidation()
            validation.Ranges.Add(XlCellRange.FromLTRB(4, firstDataRowIndex, 4, sheet.CurrentRowIndex - 1))
            validation.Type = XlDataValidationType.List
            Dim sourceRange As XlCellRange = XlCellRange.FromLTRB(0, 0, 0, departments.Count - 1).AsAbsolute()
            sourceRange.SheetName = "Departments"
            validation.Criteria1 = sourceRange
            sheet.DataValidations.Add(validation)
        End Sub
    End Class
End Namespace
