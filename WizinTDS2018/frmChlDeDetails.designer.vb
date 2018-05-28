<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChlDeDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbCNm = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmbQuarter = New System.Windows.Forms.ComboBox()
        Me.cmdc = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbsec = New System.Windows.Forms.ComboBox()
        Me.cmbtyp = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbdest = New System.Windows.Forms.ComboBox()
        Me.cmdexit = New System.Windows.Forms.Button()
        Me.cmdgen = New System.Windows.Forms.Button()
        Me.chkallde = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(6, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 30)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Company Name"
        '
        'cmbCNm
        '
        Me.cmbCNm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCNm.FormattingEnabled = True
        Me.cmbCNm.Location = New System.Drawing.Point(103, 51)
        Me.cmbCNm.Name = "cmbCNm"
        Me.cmbCNm.Size = New System.Drawing.Size(187, 21)
        Me.cmbCNm.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Lavender
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel1)
        Me.GroupBox1.Controls.Add(Me.cmdexit)
        Me.GroupBox1.Controls.Add(Me.cmdgen)
        Me.GroupBox1.Controls.Add(Me.chkallde)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.GroupBox1.Location = New System.Drawing.Point(3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(315, 325)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Deductee Allocation Report"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(6, 264)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(124, 42)
        Me.DataGridView1.TabIndex = 22
        Me.DataGridView1.Visible = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.OutsetDouble
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.84932!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 68.15069!))
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbsec, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbtyp, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbdest, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbCNm, 1, 1)
        Me.TableLayoutPanel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TableLayoutPanel1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(6, 21)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(306, 187)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbQuarter)
        Me.Panel1.Controls.Add(Me.cmdc)
        Me.Panel1.Location = New System.Drawing.Point(103, 126)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(187, 21)
        Me.Panel1.TabIndex = 4
        '
        'cmbQuarter
        '
        Me.cmbQuarter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbQuarter.FormattingEnabled = True
        Me.cmbQuarter.Location = New System.Drawing.Point(0, 0)
        Me.cmbQuarter.Name = "cmbQuarter"
        Me.cmbQuarter.Size = New System.Drawing.Size(153, 21)
        Me.cmbQuarter.TabIndex = 1
        '
        'cmdc
        '
        Me.cmdc.BackgroundImage = Global.WizinTDS2018.My.Resources.Resources._37645
        Me.cmdc.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdc.Location = New System.Drawing.Point(159, -1)
        Me.cmdc.Name = "cmdc"
        Me.cmdc.Size = New System.Drawing.Size(28, 23)
        Me.cmdc.TabIndex = 11
        Me.cmdc.TabStop = False
        Me.cmdc.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label2.Location = New System.Drawing.Point(6, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Destination"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label3.Location = New System.Drawing.Point(6, 93)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 15)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Form Type"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label4.Location = New System.Drawing.Point(6, 123)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 30)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Select Quarter"
        '
        'cmbsec
        '
        Me.cmbsec.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbsec.FormattingEnabled = True
        Me.cmbsec.Location = New System.Drawing.Point(103, 159)
        Me.cmbsec.Name = "cmbsec"
        Me.cmbsec.Size = New System.Drawing.Size(187, 21)
        Me.cmbsec.TabIndex = 5
        '
        'cmbtyp
        '
        Me.cmbtyp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbtyp.FormattingEnabled = True
        Me.cmbtyp.Location = New System.Drawing.Point(103, 96)
        Me.cmbtyp.Name = "cmbtyp"
        Me.cmbtyp.Size = New System.Drawing.Size(187, 21)
        Me.cmbtyp.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label5.Location = New System.Drawing.Point(6, 156)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 15)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Section"
        '
        'cmbdest
        '
        Me.cmbdest.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbdest.FormattingEnabled = True
        Me.cmbdest.Location = New System.Drawing.Point(103, 6)
        Me.cmbdest.Name = "cmbdest"
        Me.cmbdest.Size = New System.Drawing.Size(187, 21)
        Me.cmbdest.TabIndex = 0
        '
        'cmdexit
        '
        Me.cmdexit.Image = Global.WizinTDS2018.My.Resources.Resources.RT_EXIT1
        Me.cmdexit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdexit.Location = New System.Drawing.Point(234, 264)
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
        Me.cmdgen.Location = New System.Drawing.Point(142, 267)
        Me.cmdgen.Name = "cmdgen"
        Me.cmdgen.Size = New System.Drawing.Size(75, 42)
        Me.cmdgen.TabIndex = 0
        Me.cmdgen.Text = "Generate"
        Me.cmdgen.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdgen.UseVisualStyleBackColor = True
        '
        'chkallde
        '
        Me.chkallde.AutoSize = True
        Me.chkallde.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.chkallde.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.chkallde.Location = New System.Drawing.Point(10, 239)
        Me.chkallde.Name = "chkallde"
        Me.chkallde.Size = New System.Drawing.Size(293, 19)
        Me.chkallde.TabIndex = 12
        Me.chkallde.TabStop = False
        Me.chkallde.Text = "Show Challan with Difference in TDS Amt."
        Me.chkallde.UseVisualStyleBackColor = True
        '
        'frmChlDeDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(318, 339)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "frmChlDeDetails"
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbCNm As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbtyp As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmbsec As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbQuarter As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbdest As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkallde As System.Windows.Forms.CheckBox
    Friend WithEvents cmdc As System.Windows.Forms.Button
    Friend WithEvents cmdexit As System.Windows.Forms.Button
    Friend WithEvents cmdgen As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DataGridView1 As DataGridView
End Class
