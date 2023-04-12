Namespace XLExportExample

    Partial Class Form1

        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (Me.components IsNot Nothing) Then
                Me.components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

'#Region "Windows Form Designer generated code"
        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(XLExportExample.Form1))
            Me.btnExportToXLSX = New System.Windows.Forms.Button()
            Me.panel1 = New System.Windows.Forms.Panel()
            Me.label1 = New System.Windows.Forms.Label()
            Me.btnExportToXLS = New System.Windows.Forms.Button()
            Me.btnExportToCSV = New System.Windows.Forms.Button()
            Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
            Me.panel1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' btnExportToXLSX
            ' 
            Me.btnExportToXLSX.Image = CType((resources.GetObject("btnExportToXLSX.Image")), System.Drawing.Image)
            Me.btnExportToXLSX.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnExportToXLSX.Location = New System.Drawing.Point(11, 80)
            Me.btnExportToXLSX.Name = "btnExportToXLSX"
            Me.btnExportToXLSX.Size = New System.Drawing.Size(109, 28)
            Me.btnExportToXLSX.TabIndex = 0
            Me.btnExportToXLSX.Text = "Export to XLSX"
            Me.btnExportToXLSX.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnExportToXLSX.UseVisualStyleBackColor = True
            AddHandler Me.btnExportToXLSX.Click, New System.EventHandler(AddressOf Me.btnExportToXLSX_Click)
            ' 
            ' panel1
            ' 
            Me.panel1.BackColor = System.Drawing.SystemColors.Window
            Me.panel1.Controls.Add(Me.label1)
            Me.panel1.Dock = System.Windows.Forms.DockStyle.Top
            Me.panel1.Location = New System.Drawing.Point(0, 0)
            Me.panel1.Name = "panel1"
            Me.panel1.Padding = New System.Windows.Forms.Padding(8)
            Me.panel1.Size = New System.Drawing.Size(365, 63)
            Me.panel1.TabIndex = 1
            ' 
            ' label1
            ' 
            Me.label1.AutoSize = True
            Me.label1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte((204))))
            Me.label1.Location = New System.Drawing.Point(8, 8)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(311, 39)
            Me.label1.TabIndex = 0
            Me.label1.Text = "Generate a list of employees using the XL Export API." & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Click one of the buttons b" & "elow to save the document " & Global.Microsoft.VisualBasic.Constants.vbCrLf & "in the corresponding format."
            ' 
            ' btnExportToXLS
            ' 
            Me.btnExportToXLS.Image = CType((resources.GetObject("btnExportToXLS.Image")), System.Drawing.Image)
            Me.btnExportToXLS.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnExportToXLS.Location = New System.Drawing.Point(126, 80)
            Me.btnExportToXLS.Name = "btnExportToXLS"
            Me.btnExportToXLS.Size = New System.Drawing.Size(109, 28)
            Me.btnExportToXLS.TabIndex = 2
            Me.btnExportToXLS.Text = "Export to XLS"
            Me.btnExportToXLS.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnExportToXLS.UseVisualStyleBackColor = True
            AddHandler Me.btnExportToXLS.Click, New System.EventHandler(AddressOf Me.btnExportToXLS_Click)
            ' 
            ' btnExportToCSV
            ' 
            Me.btnExportToCSV.Image = CType((resources.GetObject("btnExportToCSV.Image")), System.Drawing.Image)
            Me.btnExportToCSV.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.btnExportToCSV.Location = New System.Drawing.Point(241, 80)
            Me.btnExportToCSV.Name = "btnExportToCSV"
            Me.btnExportToCSV.Size = New System.Drawing.Size(109, 28)
            Me.btnExportToCSV.TabIndex = 3
            Me.btnExportToCSV.Text = "Export to CSV"
            Me.btnExportToCSV.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            Me.btnExportToCSV.UseVisualStyleBackColor = True
            AddHandler Me.btnExportToCSV.Click, New System.EventHandler(AddressOf Me.btnExportToCSV_Click)
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(365, 120)
            Me.Controls.Add(Me.btnExportToCSV)
            Me.Controls.Add(Me.btnExportToXLS)
            Me.Controls.Add(Me.panel1)
            Me.Controls.Add(Me.btnExportToXLSX)
            Me.Name = "Form1"
            Me.Text = "XL Export Example"
            Me.panel1.ResumeLayout(False)
            Me.panel1.PerformLayout()
            Me.ResumeLayout(False)
        End Sub

'#End Region
        Private btnExportToXLSX As System.Windows.Forms.Button

        Private panel1 As System.Windows.Forms.Panel

        Private label1 As System.Windows.Forms.Label

        Private btnExportToXLS As System.Windows.Forms.Button

        Private btnExportToCSV As System.Windows.Forms.Button

        Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
    End Class
End Namespace
