<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmbQuarter = New System.Windows.Forms.ComboBox()
        Me.cmdc = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbsec = New System.Windows.Forms.ComboBox()
        Me.cmbtyp = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbdest = New System.Windows.Forms.ComboBox()
        Me.cmbCNm = New System.Windows.Forms.ComboBox()
        Me.cmdexit = New System.Windows.Forms.Button()
        Me.cmdgen = New System.Windows.Forms.Button()
        Me.chkChallan = New System.Windows.Forms.CheckBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        ' Me.DirectorySearcher1 = New System.DirectoryServices.DirectorySearcher()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel1)
        Me.GroupBox1.Controls.Add(Me.cmdexit)
        Me.GroupBox1.Controls.Add(Me.cmdgen)
        Me.GroupBox1.Controls.Add(Me.chkChallan)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox1.Location = New System.Drawing.Point(6, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(324, 262)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Challan Allocation Report"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(6, 218)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(116, 37)
        Me.DataGridView1.TabIndex = 22
        Me.DataGridView1.Visible = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.OutsetDouble
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.0231!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 68.9769!))
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbsec, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbtyp, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbdest, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbCNm, 1, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(6, 21)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 52.54237!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 47.45763!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(306, 167)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbQuarter)
        Me.Panel1.Controls.Add(Me.cmdc)
        Me.Panel1.Location = New System.Drawing.Point(101, 106)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(195, 23)
        Me.Panel1.TabIndex = 3
        '
        'cmbQuarter
        '
        Me.cmbQuarter.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbQuarter.FormattingEnabled = True
        Me.cmbQuarter.Location = New System.Drawing.Point(3, 0)
        Me.cmbQuarter.Name = "cmbQuarter"
        Me.cmbQuarter.Size = New System.Drawing.Size(158, 22)
        Me.cmbQuarter.TabIndex = 0
        '
        'cmdc
        '
        Me.cmdc.BackgroundImage = Global.WizinTDS2018.My.Resources.Resources._37645
        Me.cmdc.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdc.Location = New System.Drawing.Point(167, 0)
        Me.cmdc.Name = "cmdc"
        Me.cmdc.Size = New System.Drawing.Size(28, 23)
        Me.cmdc.TabIndex = 11
        Me.cmdc.TabStop = False
        Me.cmdc.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Location = New System.Drawing.Point(6, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 14)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Company Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.Location = New System.Drawing.Point(6, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 14)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Destination"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label3.Location = New System.Drawing.Point(6, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 14)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Form Type"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label4.Location = New System.Drawing.Point(6, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 14)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Select Quarter"
        '
        'cmbsec
        '
        Me.cmbsec.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbsec.FormattingEnabled = True
        Me.cmbsec.Location = New System.Drawing.Point(101, 138)
        Me.cmbsec.Name = "cmbsec"
        Me.cmbsec.Size = New System.Drawing.Size(195, 22)
        Me.cmbsec.TabIndex = 4
        '
        'cmbtyp
        '
        Me.cmbtyp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbtyp.FormattingEnabled = True
        Me.cmbtyp.Location = New System.Drawing.Point(101, 75)
        Me.cmbtyp.Name = "cmbtyp"
        Me.cmbtyp.Size = New System.Drawing.Size(195, 22)
        Me.cmbtyp.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label5.Location = New System.Drawing.Point(6, 135)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 14)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Section"
        '
        'cmbdest
        '
        Me.cmbdest.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbdest.FormattingEnabled = True
        Me.cmbdest.Location = New System.Drawing.Point(101, 6)
        Me.cmbdest.Name = "cmbdest"
        Me.cmbdest.Size = New System.Drawing.Size(195, 22)
        Me.cmbdest.TabIndex = 0
        '
        'cmbCNm
        '
        Me.cmbCNm.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCNm.FormattingEnabled = True
        Me.cmbCNm.Location = New System.Drawing.Point(101, 42)
        Me.cmbCNm.Name = "cmbCNm"
        Me.cmbCNm.Size = New System.Drawing.Size(195, 22)
        Me.cmbCNm.TabIndex = 1
        '
        'cmdexit
        '
        Me.cmdexit.Image = Global.WizinTDS2018.My.Resources.Resources.RT_EXIT1
        Me.cmdexit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdexit.Location = New System.Drawing.Point(237, 213)
        Me.cmdexit.Name = "cmdexit"
        Me.cmdexit.Size = New System.Drawing.Size(75, 42)
        Me.cmdexit.TabIndex = 21
        Me.cmdexit.TabStop = False
        Me.cmdexit.Text = "Exit"
        Me.cmdexit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdexit.UseVisualStyleBackColor = True
        '
        'cmdgen
        '
        Me.cmdgen.Image = Global.WizinTDS2018.My.Resources.Resources.icon_technology
        Me.cmdgen.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdgen.Location = New System.Drawing.Point(143, 213)
        Me.cmdgen.Name = "cmdgen"
        Me.cmdgen.Size = New System.Drawing.Size(75, 42)
        Me.cmdgen.TabIndex = 2
        Me.cmdgen.Text = "Generate"
        Me.cmdgen.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdgen.UseVisualStyleBackColor = True
        '
        'chkChallan
        '
        Me.chkChallan.AutoSize = True
        Me.chkChallan.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkChallan.ForeColor = System.Drawing.Color.DarkBlue
        Me.chkChallan.Location = New System.Drawing.Point(107, 194)
        Me.chkChallan.Name = "chkChallan"
        Me.chkChallan.Size = New System.Drawing.Size(87, 18)
        Me.chkChallan.TabIndex = 12
        Me.chkChallan.TabStop = False
        Me.chkChallan.Text = "All Deductee"
        Me.chkChallan.UseVisualStyleBackColor = True
        '
        'DirectorySearcher1
        '
        'Me.DirectorySearcher1.ClientTimeout = System.TimeSpan.Parse("-00:00:01")
        'Me.DirectorySearcher1.ServerPageTimeLimit = System.TimeSpan.Parse("-00:00:01")
        'Me.DirectorySearcher1.ServerTimeLimit = System.TimeSpan.Parse("-00:00:01")
        '
        'frmDDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(334, 269)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "frmDDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Deductee List with Challan Detail"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdexit As System.Windows.Forms.Button
    Friend WithEvents cmdgen As System.Windows.Forms.Button
    Friend WithEvents chkChallan As System.Windows.Forms.CheckBox
    Friend WithEvents cmdc As System.Windows.Forms.Button
    Friend WithEvents cmbtyp As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmbsec As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbQuarter As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbdest As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCNm As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    'Friend WithEvents DirectorySearcher1 As System.DirectoryServices.DirectorySearcher
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents DataGridView1 As DataGridView
End Class
