<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConUtility
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConUtility))
        Me.chkOldReceipt = New System.Windows.Forms.CheckBox()
        Me.CMDCHECK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtOldRRRNo = New System.Windows.Forms.TextBox()
        Me.chkCSIDownload = New System.Windows.Forms.CheckBox()
        Me.optOrginal = New System.Windows.Forms.RadioButton()
        Me.optRevised = New System.Windows.Forms.RadioButton()
        Me.cmdcsi = New System.Windows.Forms.Button()
        Me.cmdprint = New System.Windows.Forms.Button()
        Me.cmdShowStatFile = New System.Windows.Forms.Button()
        Me.cmdShowErrFile = New System.Windows.Forms.Button()
        Me.cmdOpenTxtFile = New System.Windows.Forms.Button()
        Me.cmdFVU = New System.Windows.Forms.Button()
        Me.cmdConvert = New System.Windows.Forms.Button()
        Me.lblQtrDisplay = New System.Windows.Forms.Label()
        Me.cdgconvert = New System.Windows.Forms.OpenFileDialog()
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.cdgOpenCSI = New System.Windows.Forms.OpenFileDialog()
        Me.cdgerrfile = New System.Windows.Forms.OpenFileDialog()
        Me.chkAutoLaunch = New System.Windows.Forms.CheckBox()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkOldReceipt
        '
        Me.chkOldReceipt.AutoSize = True
        Me.chkOldReceipt.Location = New System.Drawing.Point(27, 35)
        Me.chkOldReceipt.Name = "chkOldReceipt"
        Me.chkOldReceipt.Size = New System.Drawing.Size(251, 17)
        Me.chkOldReceipt.TabIndex = 0
        Me.chkOldReceipt.TabStop = False
        Me.chkOldReceipt.Text = "Whether regular statement filed for earlier period"
        Me.chkOldReceipt.UseVisualStyleBackColor = True
        '
        'CMDCHECK
        '
        Me.CMDCHECK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMDCHECK.Location = New System.Drawing.Point(213, 56)
        Me.CMDCHECK.Name = "CMDCHECK"
        Me.CMDCHECK.Size = New System.Drawing.Size(75, 23)
        Me.CMDCHECK.TabIndex = 1
        Me.CMDCHECK.Text = "Select PNR"
        Me.CMDCHECK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Receipt No"
        '
        'txtOldRRRNo
        '
        Me.txtOldRRRNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOldRRRNo.Location = New System.Drawing.Point(107, 55)
        Me.txtOldRRRNo.MaxLength = 15
        Me.txtOldRRRNo.Name = "txtOldRRRNo"
        Me.txtOldRRRNo.Size = New System.Drawing.Size(100, 20)
        Me.txtOldRRRNo.TabIndex = 0
        '
        'chkCSIDownload
        '
        Me.chkCSIDownload.AutoSize = True
        Me.chkCSIDownload.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCSIDownload.Location = New System.Drawing.Point(326, 23)
        Me.chkCSIDownload.Name = "chkCSIDownload"
        Me.chkCSIDownload.Size = New System.Drawing.Size(243, 19)
        Me.chkCSIDownload.TabIndex = 6
        Me.chkCSIDownload.Text = "Auto download CSI file before Validating"
        Me.chkCSIDownload.UseVisualStyleBackColor = True
        Me.chkCSIDownload.Visible = False
        '
        'optOrginal
        '
        Me.optOrginal.AutoSize = True
        Me.optOrginal.Checked = True
        Me.optOrginal.Location = New System.Drawing.Point(638, 55)
        Me.optOrginal.Name = "optOrginal"
        Me.optOrginal.Size = New System.Drawing.Size(14, 13)
        Me.optOrginal.TabIndex = 7
        Me.optOrginal.TabStop = True
        Me.optOrginal.UseVisualStyleBackColor = True
        Me.optOrginal.Visible = False
        '
        'optRevised
        '
        Me.optRevised.AutoSize = True
        Me.optRevised.Location = New System.Drawing.Point(638, 74)
        Me.optRevised.Name = "optRevised"
        Me.optRevised.Size = New System.Drawing.Size(14, 13)
        Me.optRevised.TabIndex = 8
        Me.optRevised.UseVisualStyleBackColor = True
        Me.optRevised.Visible = False
        '
        'cmdcsi
        '
        Me.cmdcsi.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdcsi.Location = New System.Drawing.Point(4, 122)
        Me.cmdcsi.Name = "cmdcsi"
        Me.cmdcsi.Size = New System.Drawing.Size(116, 44)
        Me.cmdcsi.TabIndex = 9
        Me.cmdcsi.Text = "CSI Download"
        Me.cmdcsi.UseVisualStyleBackColor = True
        '
        'cmdprint
        '
        Me.cmdprint.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdprint.Location = New System.Drawing.Point(4, 455)
        Me.cmdprint.Name = "cmdprint"
        Me.cmdprint.Size = New System.Drawing.Size(116, 44)
        Me.cmdprint.TabIndex = 11
        Me.cmdprint.Text = "Open 27A File"
        Me.cmdprint.UseVisualStyleBackColor = True
        '
        'cmdShowStatFile
        '
        Me.cmdShowStatFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowStatFile.Location = New System.Drawing.Point(4, 404)
        Me.cmdShowStatFile.Name = "cmdShowStatFile"
        Me.cmdShowStatFile.Size = New System.Drawing.Size(116, 44)
        Me.cmdShowStatFile.TabIndex = 12
        Me.cmdShowStatFile.Text = "Open TDS Return Statistics Report"
        Me.cmdShowStatFile.UseVisualStyleBackColor = True
        '
        'cmdShowErrFile
        '
        Me.cmdShowErrFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowErrFile.Location = New System.Drawing.Point(4, 347)
        Me.cmdShowErrFile.Name = "cmdShowErrFile"
        Me.cmdShowErrFile.Size = New System.Drawing.Size(116, 44)
        Me.cmdShowErrFile.TabIndex = 13
        Me.cmdShowErrFile.Text = "Open TDS/TCS - Error File"
        Me.cmdShowErrFile.UseVisualStyleBackColor = True
        '
        'cmdOpenTxtFile
        '
        Me.cmdOpenTxtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenTxtFile.Location = New System.Drawing.Point(4, 292)
        Me.cmdOpenTxtFile.Name = "cmdOpenTxtFile"
        Me.cmdOpenTxtFile.Size = New System.Drawing.Size(116, 44)
        Me.cmdOpenTxtFile.TabIndex = 14
        Me.cmdOpenTxtFile.Text = "Open the Converted Text File"
        Me.cmdOpenTxtFile.UseVisualStyleBackColor = True
        '
        'cmdFVU
        '
        Me.cmdFVU.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFVU.Location = New System.Drawing.Point(4, 234)
        Me.cmdFVU.Name = "cmdFVU"
        Me.cmdFVU.Size = New System.Drawing.Size(116, 44)
        Me.cmdFVU.TabIndex = 15
        Me.cmdFVU.Text = "Invoke NSDL's File Validation Utility"
        Me.cmdFVU.UseVisualStyleBackColor = True
        '
        'cmdConvert
        '
        Me.cmdConvert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdConvert.Location = New System.Drawing.Point(4, 178)
        Me.cmdConvert.Name = "cmdConvert"
        Me.cmdConvert.Size = New System.Drawing.Size(116, 44)
        Me.cmdConvert.TabIndex = 16
        Me.cmdConvert.Text = "Convert this Return into Text File"
        Me.cmdConvert.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdConvert.UseVisualStyleBackColor = True
        '
        'lblQtrDisplay
        '
        Me.lblQtrDisplay.AutoSize = True
        Me.lblQtrDisplay.Location = New System.Drawing.Point(396, 9)
        Me.lblQtrDisplay.MaximumSize = New System.Drawing.Size(10, 20)
        Me.lblQtrDisplay.MinimumSize = New System.Drawing.Size(10, 20)
        Me.lblQtrDisplay.Name = "lblQtrDisplay"
        Me.lblQtrDisplay.Size = New System.Drawing.Size(10, 20)
        Me.lblQtrDisplay.TabIndex = 17
        '
        'cdgconvert
        '
        Me.cdgconvert.FileName = "cdgconvert"
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Dock = System.Windows.Forms.DockStyle.None
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem})
        Me.BindingNavigator1.Location = New System.Drawing.Point(4, 4)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(255, 25)
        Me.BindingNavigator1.TabIndex = 18
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 22)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'cdgOpenCSI
        '
        Me.cdgOpenCSI.FileName = "cdgOpenCSI"
        '
        'cdgerrfile
        '
        Me.cdgerrfile.FileName = "cdgconvert"
        '
        'chkAutoLaunch
        '
        Me.chkAutoLaunch.AutoSize = True
        Me.chkAutoLaunch.Checked = True
        Me.chkAutoLaunch.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAutoLaunch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAutoLaunch.Location = New System.Drawing.Point(130, 101)
        Me.chkAutoLaunch.Name = "chkAutoLaunch"
        Me.chkAutoLaunch.Size = New System.Drawing.Size(343, 19)
        Me.chkAutoLaunch.TabIndex = 19
        Me.chkAutoLaunch.Text = "Auto Launch Validation Utility after Conversion to Text File1"
        Me.chkAutoLaunch.UseVisualStyleBackColor = True
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Location = New System.Drawing.Point(126, 94)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(707, 412)
        Me.WebBrowser1.TabIndex = 5
        '
        'frmConUtility
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Lavender
        Me.ClientSize = New System.Drawing.Size(836, 518)
        Me.Controls.Add(Me.chkAutoLaunch)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Controls.Add(Me.lblQtrDisplay)
        Me.Controls.Add(Me.cmdConvert)
        Me.Controls.Add(Me.cmdFVU)
        Me.Controls.Add(Me.cmdOpenTxtFile)
        Me.Controls.Add(Me.cmdShowErrFile)
        Me.Controls.Add(Me.cmdShowStatFile)
        Me.Controls.Add(Me.cmdprint)
        Me.Controls.Add(Me.cmdcsi)
        Me.Controls.Add(Me.optRevised)
        Me.Controls.Add(Me.optOrginal)
        Me.Controls.Add(Me.chkCSIDownload)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.txtOldRRRNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CMDCHECK)
        Me.Controls.Add(Me.chkOldReceipt)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "frmConUtility"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Conversion Utility"
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkOldReceipt As System.Windows.Forms.CheckBox
    Friend WithEvents CMDCHECK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOldRRRNo As System.Windows.Forms.TextBox
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents chkCSIDownload As System.Windows.Forms.CheckBox
    Friend WithEvents optOrginal As System.Windows.Forms.RadioButton
    Friend WithEvents optRevised As System.Windows.Forms.RadioButton
    Friend WithEvents cmdcsi As System.Windows.Forms.Button
    Friend WithEvents cmdprint As System.Windows.Forms.Button
    Friend WithEvents cmdShowStatFile As System.Windows.Forms.Button
    Friend WithEvents cmdShowErrFile As System.Windows.Forms.Button
    Friend WithEvents cmdOpenTxtFile As System.Windows.Forms.Button
    Friend WithEvents cmdFVU As System.Windows.Forms.Button
    Friend WithEvents cmdConvert As System.Windows.Forms.Button
    Friend WithEvents lblQtrDisplay As Label
    Friend WithEvents cdgconvert As OpenFileDialog
    Friend WithEvents BindingNavigator1 As BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As ToolStripSeparator
    Friend WithEvents cdgOpenCSI As OpenFileDialog
    Friend WithEvents cdgerrfile As OpenFileDialog
    Friend WithEvents chkAutoLaunch As CheckBox
End Class
