<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm16A
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
        Me.txtSignByName = New System.Windows.Forms.TextBox()
        Me.txtSignByFatherName = New System.Windows.Forms.TextBox()
        Me.txtSignByCapacity = New System.Windows.Forms.TextBox()
        Me.txtPlace = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cmdgen = New System.Windows.Forms.Button()
        Me.cmdexit = New System.Windows.Forms.Button()
        Me.chkpanvalid = New System.Windows.Forms.CheckBox()
        Me.chkXLOverWrite = New System.Windows.Forms.CheckBox()
        Me.chkOpenXL = New System.Windows.Forms.CheckBox()
        Me.cmbDeductee = New System.Windows.Forms.ComboBox()
        Me.cmdc = New System.Windows.Forms.Button()
        Me.certidt = New System.Windows.Forms.MaskedTextBox()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtSignByName
        '
        Me.txtSignByName.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSignByName.Location = New System.Drawing.Point(240, 47)
        Me.txtSignByName.Name = "txtSignByName"
        Me.txtSignByName.Size = New System.Drawing.Size(203, 23)
        Me.txtSignByName.TabIndex = 0
        '
        'txtSignByFatherName
        '
        Me.txtSignByFatherName.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSignByFatherName.Location = New System.Drawing.Point(240, 74)
        Me.txtSignByFatherName.Name = "txtSignByFatherName"
        Me.txtSignByFatherName.Size = New System.Drawing.Size(203, 23)
        Me.txtSignByFatherName.TabIndex = 1
        '
        'txtSignByCapacity
        '
        Me.txtSignByCapacity.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSignByCapacity.Location = New System.Drawing.Point(240, 106)
        Me.txtSignByCapacity.Name = "txtSignByCapacity"
        Me.txtSignByCapacity.Size = New System.Drawing.Size(203, 23)
        Me.txtSignByCapacity.TabIndex = 2
        '
        'txtPlace
        '
        Me.txtPlace.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPlace.Location = New System.Drawing.Point(240, 135)
        Me.txtPlace.Name = "txtPlace"
        Me.txtPlace.Size = New System.Drawing.Size(203, 23)
        Me.txtPlace.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(172, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Father's Name of Signatory*"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 106)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(177, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Designation of the Signatory*"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(148, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Name of The Signatory*"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(12, 135)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 16)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Place*"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(12, 164)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(104, 16)
        Me.Label11.TabIndex = 8
        Me.Label11.Text = "Certificate Date*"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 195)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(98, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Deductee Name"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(0, 330)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(169, 16)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "All * Fields are Compulsory."
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Red
        Me.Label13.Location = New System.Drawing.Point(243, 330)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(221, 32)
        Me.Label13.TabIndex = 11
        Me.Label13.Text = "Remark:- Tick this for deductees not " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "              having PAN."
        '
        'cmdgen
        '
        Me.cmdgen.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdgen.Location = New System.Drawing.Point(75, 371)
        Me.cmdgen.Name = "cmdgen"
        Me.cmdgen.Size = New System.Drawing.Size(75, 23)
        Me.cmdgen.TabIndex = 12
        Me.cmdgen.Text = "Generate"
        Me.cmdgen.UseVisualStyleBackColor = True
        '
        'cmdexit
        '
        Me.cmdexit.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdexit.Location = New System.Drawing.Point(190, 371)
        Me.cmdexit.Name = "cmdexit"
        Me.cmdexit.Size = New System.Drawing.Size(75, 23)
        Me.cmdexit.TabIndex = 13
        Me.cmdexit.Text = "Exit"
        Me.cmdexit.UseVisualStyleBackColor = True
        '
        'chkpanvalid
        '
        Me.chkpanvalid.AutoSize = True
        Me.chkpanvalid.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkpanvalid.ForeColor = System.Drawing.Color.Red
        Me.chkpanvalid.Location = New System.Drawing.Point(240, 291)
        Me.chkpanvalid.Name = "chkpanvalid"
        Me.chkpanvalid.Size = New System.Drawing.Size(224, 36)
        Me.chkpanvalid.TabIndex = 14
        Me.chkpanvalid.Text = "Tick this to include deductee not " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "having valid PAN in the display list."
        Me.chkpanvalid.UseVisualStyleBackColor = True
        '
        'chkXLOverWrite
        '
        Me.chkXLOverWrite.AutoSize = True
        Me.chkXLOverWrite.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkXLOverWrite.Location = New System.Drawing.Point(131, 256)
        Me.chkXLOverWrite.Name = "chkXLOverWrite"
        Me.chkXLOverWrite.Size = New System.Drawing.Size(275, 20)
        Me.chkXLOverWrite.TabIndex = 15
        Me.chkXLOverWrite.Text = "Check here to overwrite Excel files(if exits)."
        Me.chkXLOverWrite.UseVisualStyleBackColor = True
        '
        'chkOpenXL
        '
        Me.chkOpenXL.AutoSize = True
        Me.chkOpenXL.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOpenXL.Location = New System.Drawing.Point(131, 230)
        Me.chkOpenXL.Name = "chkOpenXL"
        Me.chkOpenXL.Size = New System.Drawing.Size(257, 20)
        Me.chkOpenXL.TabIndex = 16
        Me.chkOpenXL.Text = "Check here to open Excel file after save."
        Me.chkOpenXL.UseVisualStyleBackColor = True
        '
        'cmbDeductee
        '
        Me.cmbDeductee.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDeductee.FormattingEnabled = True
        Me.cmbDeductee.Location = New System.Drawing.Point(240, 188)
        Me.cmbDeductee.Name = "cmbDeductee"
        Me.cmbDeductee.Size = New System.Drawing.Size(203, 24)
        Me.cmbDeductee.TabIndex = 17
        '
        'cmdc
        '
        Me.cmdc.BackgroundImage = Global.WizinTDS2018.My.Resources.Resources._37645
        Me.cmdc.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdc.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdc.Location = New System.Drawing.Point(449, 188)
        Me.cmdc.Name = "cmdc"
        Me.cmdc.Size = New System.Drawing.Size(28, 23)
        Me.cmdc.TabIndex = 18
        Me.cmdc.TabStop = False
        Me.cmdc.UseVisualStyleBackColor = True
        '
        'certidt
        '
        Me.certidt.Location = New System.Drawing.Point(240, 163)
        Me.certidt.Mask = "00/00/0000"
        Me.certidt.Name = "certidt"
        Me.certidt.Size = New System.Drawing.Size(203, 20)
        Me.certidt.TabIndex = 19
        Me.certidt.ValidatingType = GetType(Date)
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Location = New System.Drawing.Point(187, 9)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(0, 13)
        Me.lblMsg.TabIndex = 21
        '
        'frm16A
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(514, 406)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.certidt)
        Me.Controls.Add(Me.cmdc)
        Me.Controls.Add(Me.cmbDeductee)
        Me.Controls.Add(Me.chkOpenXL)
        Me.Controls.Add(Me.chkXLOverWrite)
        Me.Controls.Add(Me.chkpanvalid)
        Me.Controls.Add(Me.cmdexit)
        Me.Controls.Add(Me.cmdgen)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPlace)
        Me.Controls.Add(Me.txtSignByCapacity)
        Me.Controls.Add(Me.txtSignByFatherName)
        Me.Controls.Add(Me.txtSignByName)
        Me.Name = "frm16A"
        Me.Text = "frm16A"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtSignByName As TextBox
    Friend WithEvents txtSignByFatherName As TextBox
    Friend WithEvents txtSignByCapacity As TextBox
    Friend WithEvents txtPlace As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents cmdgen As Button
    Friend WithEvents cmdexit As Button
    Friend WithEvents chkpanvalid As CheckBox
    Friend WithEvents chkXLOverWrite As CheckBox
    Friend WithEvents chkOpenXL As CheckBox
    Friend WithEvents cmbDeductee As ComboBox
    Friend WithEvents cmdc As Button
    Friend WithEvents certidt As MaskedTextBox
    Friend WithEvents lblMsg As Label
End Class
