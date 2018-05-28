<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmdeduteeTDSMST
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboDName = New System.Windows.Forms.ComboBox()
        Me.chkinact = New System.Windows.Forms.CheckBox()
        Me.txtDName = New System.Windows.Forms.TextBox()
        Me.txtRef = New System.Windows.Forms.TextBox()
        Me.txtDPAN = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbocat = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cboCategory = New System.Windows.Forms.ComboBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtDDesgn = New System.Windows.Forms.TextBox()
        Me.cmdcorrection = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.Dedsearch = New System.Windows.Forms.Button()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtDAdd5 = New System.Windows.Forms.TextBox()
        Me.txtDAdd4 = New System.Windows.Forms.TextBox()
        Me.txtDAdd3 = New System.Windows.Forms.TextBox()
        Me.txtDAdd2 = New System.Windows.Forms.TextBox()
        Me.txtDAdd1 = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cboDState = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtDPin = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.optCo = New System.Windows.Forms.RadioButton()
        Me.optOther = New System.Windows.Forms.RadioButton()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CboCollNonRes = New System.Windows.Forms.ComboBox()
        Me.CboPerEstInd = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel8 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtDeEmail = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtDePhone = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel9 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtTIN = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtDid = New System.Windows.Forms.TextBox()
        Me.cmdSearch = New System.Windows.Forms.Button()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.TableLayoutPanel8.SuspendLayout()
        Me.TableLayoutPanel9.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(41, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(438, 28)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Use this form only for modifications in the Deductee Details.  The new        " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "d" &
    "eductees will be  added automatically when you prepare the returns."
        '
        'Label2
        '
        Me.Label2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label2.Location = New System.Drawing.Point(7, 3)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(122, 28)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Selec&t Name"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label5.Location = New System.Drawing.Point(7, 29)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 79)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Deductee's Addre&ss"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label6.Location = New System.Drawing.Point(7, 3)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 23)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Deductee's &Name"
        '
        'cboDName
        '
        Me.cboDName.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cboDName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboDName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboDName.FormattingEnabled = True
        Me.cboDName.Location = New System.Drawing.Point(140, 6)
        Me.cboDName.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cboDName.MaxLength = 100
        Me.cboDName.Name = "cboDName"
        Me.cboDName.Size = New System.Drawing.Size(195, 23)
        Me.cboDName.Sorted = True
        Me.cboDName.TabIndex = 0
        '
        'chkinact
        '
        Me.chkinact.AutoSize = True
        Me.chkinact.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkinact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.chkinact.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.chkinact.Location = New System.Drawing.Point(346, 6)
        Me.chkinact.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.chkinact.Name = "chkinact"
        Me.chkinact.Size = New System.Drawing.Size(119, 22)
        Me.chkinact.TabIndex = 1
        Me.chkinact.Text = "Show Inactive " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Deductee Only"
        Me.chkinact.UseVisualStyleBackColor = True
        '
        'txtDName
        '
        Me.txtDName.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDName.Location = New System.Drawing.Point(138, 6)
        Me.txtDName.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDName.MaxLength = 75
        Me.txtDName.Name = "txtDName"
        Me.txtDName.Size = New System.Drawing.Size(351, 21)
        Me.txtDName.TabIndex = 0
        '
        'txtRef
        '
        Me.txtRef.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRef.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRef.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRef.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRef.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtRef.Location = New System.Drawing.Point(352, 6)
        Me.txtRef.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtRef.MaxLength = 10
        Me.txtRef.Name = "txtRef"
        Me.txtRef.Size = New System.Drawing.Size(136, 20)
        Me.txtRef.TabIndex = 1
        '
        'txtDPAN
        '
        Me.txtDPAN.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDPAN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDPAN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDPAN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDPAN.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDPAN.Location = New System.Drawing.Point(138, 6)
        Me.txtDPAN.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDPAN.MaxLength = 10
        Me.txtDPAN.Name = "txtDPAN"
        Me.txtDPAN.Size = New System.Drawing.Size(130, 20)
        Me.txtDPAN.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label7.Location = New System.Drawing.Point(7, 3)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 26)
        Me.Label7.TabIndex = 1
        Me.Label7.Text = "PAN"
        '
        'cbocat
        '
        Me.cbocat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbocat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbocat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbocat.FormattingEnabled = True
        Me.cbocat.Location = New System.Drawing.Point(139, 6)
        Me.cbocat.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cbocat.Name = "cbocat"
        Me.cbocat.Size = New System.Drawing.Size(350, 21)
        Me.cbocat.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label8.Location = New System.Drawing.Point(7, 3)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(121, 26)
        Me.Label8.TabIndex = 30
        Me.Label8.Text = "Category for PAN"
        '
        'Label11
        '
        Me.Label11.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label11.Location = New System.Drawing.Point(279, 3)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(62, 26)
        Me.Label11.TabIndex = 35
        Me.Label11.Text = "REF"
        '
        'cboCategory
        '
        Me.cboCategory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCategory.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboCategory.FormattingEnabled = True
        Me.cboCategory.Location = New System.Drawing.Point(225, 6)
        Me.cboCategory.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(263, 21)
        Me.cboCategory.TabIndex = 7
        '
        'Label14
        '
        Me.Label14.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label14.Location = New System.Drawing.Point(7, 3)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(207, 26)
        Me.Label14.TabIndex = 37
        Me.Label14.Text = "Category"
        '
        'Label15
        '
        Me.Label15.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label15.Location = New System.Drawing.Point(7, 32)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(207, 29)
        Me.Label15.TabIndex = 36
        Me.Label15.Text = "Designation (if any)"
        '
        'txtDDesgn
        '
        Me.txtDDesgn.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDDesgn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDDesgn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDDesgn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDDesgn.Location = New System.Drawing.Point(225, 35)
        Me.txtDDesgn.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDDesgn.MaxLength = 25
        Me.txtDDesgn.Name = "txtDDesgn"
        Me.txtDDesgn.Size = New System.Drawing.Size(263, 20)
        Me.txtDDesgn.TabIndex = 8
        '
        'cmdcorrection
        '
        Me.cmdcorrection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdcorrection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdcorrection.Image = Global.WizinTDS2018.My.Resources.Resources.Knob_Search_icon
        Me.cmdcorrection.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdcorrection.Location = New System.Drawing.Point(29, 505)
        Me.cmdcorrection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdcorrection.Name = "cmdcorrection"
        Me.cmdcorrection.Size = New System.Drawing.Size(73, 48)
        Me.cmdcorrection.TabIndex = 10
        Me.cmdcorrection.TabStop = False
        Me.cmdcorrection.Text = "Correction"
        Me.cmdcorrection.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdcorrection.UseVisualStyleBackColor = True
        Me.cmdcorrection.Visible = False
        '
        'cmdSave
        '
        Me.cmdSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Image = Global.WizinTDS2018.My.Resources.Resources.RT_SAVE1
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdSave.Location = New System.Drawing.Point(123, 539)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(83, 48)
        Me.cmdSave.TabIndex = 11
        Me.cmdSave.Text = "Save"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Image = Global.WizinTDS2018.My.Resources.Resources.Knob_Cancel_icon
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCancel.Location = New System.Drawing.Point(313, 539)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(84, 48)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.TabStop = False
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Image = Global.WizinTDS2018.My.Resources.Resources.Knob_Refresh_icon
        Me.cmdDelete.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDelete.Location = New System.Drawing.Point(218, 539)
        Me.cmdDelete.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(87, 48)
        Me.cmdDelete.TabIndex = 12
        Me.cmdDelete.TabStop = False
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'Dedsearch
        '
        Me.Dedsearch.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Dedsearch.Location = New System.Drawing.Point(469, 521)
        Me.Dedsearch.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Dedsearch.Name = "Dedsearch"
        Me.Dedsearch.Size = New System.Drawing.Size(10, 40)
        Me.Dedsearch.TabIndex = 48
        Me.Dedsearch.UseVisualStyleBackColor = True
        Me.Dedsearch.Visible = False
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel3.ColumnCount = 3
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.86567!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 42.00426!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.13006!))
        Me.TableLayoutPanel3.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.cboDName, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.chkinact, 2, 0)
        Me.TableLayoutPanel3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(0, 46)
        Me.TableLayoutPanel3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(497, 34)
        Me.TableLayoutPanel3.TabIndex = 0
        Me.TableLayoutPanel3.TabStop = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.36917!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.63083!))
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDName, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 1)
        Me.TableLayoutPanel1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(1, 80)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 23.52941!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 76.47059!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(496, 111)
        Me.TableLayoutPanel1.TabIndex = 1
        Me.TableLayoutPanel1.TabStop = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.txtDAdd5)
        Me.Panel1.Controls.Add(Me.txtDAdd4)
        Me.Panel1.Controls.Add(Me.txtDAdd3)
        Me.Panel1.Controls.Add(Me.txtDAdd2)
        Me.Panel1.Controls.Add(Me.txtDAdd1)
        Me.Panel1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Panel1.Location = New System.Drawing.Point(138, 32)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(351, 73)
        Me.Panel1.TabIndex = 10
        '
        'txtDAdd5
        '
        Me.txtDAdd5.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd5.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDAdd5.Location = New System.Drawing.Point(0, 60)
        Me.txtDAdd5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd5.MaxLength = 25
        Me.txtDAdd5.Name = "txtDAdd5"
        Me.txtDAdd5.Size = New System.Drawing.Size(358, 14)
        Me.txtDAdd5.TabIndex = 5
        '
        'txtDAdd4
        '
        Me.txtDAdd4.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDAdd4.Location = New System.Drawing.Point(0, 45)
        Me.txtDAdd4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd4.MaxLength = 25
        Me.txtDAdd4.Name = "txtDAdd4"
        Me.txtDAdd4.Size = New System.Drawing.Size(358, 14)
        Me.txtDAdd4.TabIndex = 4
        '
        'txtDAdd3
        '
        Me.txtDAdd3.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDAdd3.Location = New System.Drawing.Point(0, 30)
        Me.txtDAdd3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd3.MaxLength = 25
        Me.txtDAdd3.Name = "txtDAdd3"
        Me.txtDAdd3.Size = New System.Drawing.Size(358, 14)
        Me.txtDAdd3.TabIndex = 3
        '
        'txtDAdd2
        '
        Me.txtDAdd2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDAdd2.Location = New System.Drawing.Point(0, 15)
        Me.txtDAdd2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd2.MaxLength = 25
        Me.txtDAdd2.Name = "txtDAdd2"
        Me.txtDAdd2.Size = New System.Drawing.Size(358, 14)
        Me.txtDAdd2.TabIndex = 2
        '
        'txtDAdd1
        '
        Me.txtDAdd1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDAdd1.Location = New System.Drawing.Point(0, 0)
        Me.txtDAdd1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd1.MaxLength = 25
        Me.txtDAdd1.Name = "txtDAdd1"
        Me.txtDAdd1.Size = New System.Drawing.Size(353, 14)
        Me.txtDAdd1.TabIndex = 1
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.77485!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.80325!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.21298!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.6004!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label9, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.cboDState, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label10, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.txtDPin, 3, 0)
        Me.TableLayoutPanel2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(1, 192)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(496, 39)
        Me.TableLayoutPanel2.TabIndex = 2
        Me.TableLayoutPanel2.TabStop = True
        '
        'Label9
        '
        Me.Label9.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label9.Location = New System.Drawing.Point(7, 3)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(121, 33)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "State"
        '
        'cboDState
        '
        Me.cboDState.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDState.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cboDState.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboDState.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDState.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboDState.FormattingEnabled = True
        Me.cboDState.Location = New System.Drawing.Point(139, 6)
        Me.cboDState.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cboDState.Name = "cboDState"
        Me.cboDState.Size = New System.Drawing.Size(131, 21)
        Me.cboDState.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label10.Location = New System.Drawing.Point(281, 3)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 33)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Pin Code"
        '
        'txtDPin
        '
        Me.txtDPin.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDPin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDPin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDPin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDPin.Location = New System.Drawing.Point(357, 6)
        Me.txtDPin.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDPin.MaxLength = 6
        Me.txtDPin.Name = "txtDPin"
        Me.txtDPin.Size = New System.Drawing.Size(132, 20)
        Me.txtDPin.TabIndex = 1
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.57201!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.42799!))
        Me.TableLayoutPanel4.Controls.Add(Me.Label8, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.cbocat, 1, 0)
        Me.TableLayoutPanel4.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(1, 231)
        Me.TableLayoutPanel4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(496, 32)
        Me.TableLayoutPanel4.TabIndex = 3
        Me.TableLayoutPanel4.TabStop = True
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.optCo)
        Me.Panel2.Controls.Add(Me.optOther)
        Me.Panel2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Panel2.Location = New System.Drawing.Point(1, 289)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(496, 38)
        Me.Panel2.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label3.Location = New System.Drawing.Point(5, 11)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Deductee Type"
        '
        'optCo
        '
        Me.optCo.AutoSize = True
        Me.optCo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optCo.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.optCo.Location = New System.Drawing.Point(275, 2)
        Me.optCo.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.optCo.Name = "optCo"
        Me.optCo.Size = New System.Drawing.Size(86, 30)
        Me.optCo.TabIndex = 1
        Me.optCo.Text = "Company/" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Corporates"
        Me.optCo.UseVisualStyleBackColor = True
        '
        'optOther
        '
        Me.optOther.AutoSize = True
        Me.optOther.Checked = True
        Me.optOther.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optOther.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.optOther.Location = New System.Drawing.Point(140, 3)
        Me.optOther.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.optOther.Name = "optOther"
        Me.optOther.Size = New System.Drawing.Size(85, 30)
        Me.optOther.TabIndex = 0
        Me.optOther.TabStop = True
        Me.optOther.Text = "Other than" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Company"
        Me.optOther.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel5.ColumnCount = 2
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 44.30894!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 55.69106!))
        Me.TableLayoutPanel5.Controls.Add(Me.Label14, 0, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.Label15, 0, 1)
        Me.TableLayoutPanel5.Controls.Add(Me.txtDDesgn, 1, 1)
        Me.TableLayoutPanel5.Controls.Add(Me.cboCategory, 1, 0)
        Me.TableLayoutPanel5.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(2, 359)
        Me.TableLayoutPanel5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 2
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 48.4375!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 51.5625!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(495, 64)
        Me.TableLayoutPanel5.TabIndex = 7
        Me.TableLayoutPanel5.TabStop = True
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel6.ColumnCount = 4
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.62602!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.65854!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.63415!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 29.47154!))
        Me.TableLayoutPanel6.Controls.Add(Me.Label11, 2, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.txtDPAN, 1, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.Label7, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.txtRef, 3, 0)
        Me.TableLayoutPanel6.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(2, 263)
        Me.TableLayoutPanel6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 1
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(495, 32)
        Me.TableLayoutPanel6.TabIndex = 4
        Me.TableLayoutPanel6.TabStop = True
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel7.ColumnCount = 2
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 79.10751!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.89249!))
        Me.TableLayoutPanel7.Controls.Add(Me.Label12, 0, 1)
        Me.TableLayoutPanel7.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.CboCollNonRes, 1, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.CboPerEstInd, 1, 1)
        Me.TableLayoutPanel7.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(1, 455)
        Me.TableLayoutPanel7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 2
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.87719!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.12281!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(496, 60)
        Me.TableLayoutPanel7.TabIndex = 9
        Me.TableLayoutPanel7.TabStop = True
        '
        'Label12
        '
        Me.Label12.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label12.Location = New System.Drawing.Point(7, 31)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(377, 26)
        Me.Label12.TabIndex = 39
        Me.Label12.Text = "Deductee is having Permanent Establishment in India"
        Me.Label12.Visible = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label4.Location = New System.Drawing.Point(7, 3)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(377, 25)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "Collectee is Non Resident"
        '
        'CboCollNonRes
        '
        Me.CboCollNonRes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCollNonRes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboCollNonRes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CboCollNonRes.FormattingEnabled = True
        Me.CboCollNonRes.Location = New System.Drawing.Point(395, 6)
        Me.CboCollNonRes.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CboCollNonRes.Name = "CboCollNonRes"
        Me.CboCollNonRes.Size = New System.Drawing.Size(94, 21)
        Me.CboCollNonRes.TabIndex = 0
        '
        'CboPerEstInd
        '
        Me.CboPerEstInd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboPerEstInd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboPerEstInd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CboPerEstInd.FormattingEnabled = True
        Me.CboPerEstInd.Location = New System.Drawing.Point(395, 34)
        Me.CboPerEstInd.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CboPerEstInd.Name = "CboPerEstInd"
        Me.CboPerEstInd.Size = New System.Drawing.Size(94, 21)
        Me.CboPerEstInd.TabIndex = 1
        Me.CboPerEstInd.Visible = False
        '
        'TableLayoutPanel8
        '
        Me.TableLayoutPanel8.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel8.ColumnCount = 4
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.42276!))
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.84553!))
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.8374!))
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.4878!))
        Me.TableLayoutPanel8.Controls.Add(Me.Label13, 2, 0)
        Me.TableLayoutPanel8.Controls.Add(Me.txtDeEmail, 1, 0)
        Me.TableLayoutPanel8.Controls.Add(Me.Label16, 0, 0)
        Me.TableLayoutPanel8.Controls.Add(Me.txtDePhone, 3, 0)
        Me.TableLayoutPanel8.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel8.Location = New System.Drawing.Point(2, 327)
        Me.TableLayoutPanel8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel8.Name = "TableLayoutPanel8"
        Me.TableLayoutPanel8.RowCount = 1
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel8.Size = New System.Drawing.Size(495, 31)
        Me.TableLayoutPanel8.TabIndex = 6
        Me.TableLayoutPanel8.TabStop = True
        '
        'Label13
        '
        Me.Label13.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label13.Location = New System.Drawing.Point(274, 3)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(63, 25)
        Me.Label13.TabIndex = 35
        Me.Label13.Text = "Phone No."
        '
        'txtDeEmail
        '
        Me.txtDeEmail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDeEmail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDeEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeEmail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDeEmail.Location = New System.Drawing.Point(137, 6)
        Me.txtDeEmail.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDeEmail.MaxLength = 50
        Me.txtDeEmail.Name = "txtDeEmail"
        Me.txtDeEmail.Size = New System.Drawing.Size(126, 20)
        Me.txtDeEmail.TabIndex = 0
        '
        'Label16
        '
        Me.Label16.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label16.Location = New System.Drawing.Point(7, 3)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(119, 25)
        Me.Label16.TabIndex = 1
        Me.Label16.Text = "E-MAIL"
        '
        'txtDePhone
        '
        Me.txtDePhone.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDePhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDePhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDePhone.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDePhone.Location = New System.Drawing.Point(348, 6)
        Me.txtDePhone.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDePhone.MaxLength = 11
        Me.txtDePhone.Name = "txtDePhone"
        Me.txtDePhone.Size = New System.Drawing.Size(140, 20)
        Me.txtDePhone.TabIndex = 1
        '
        'TableLayoutPanel9
        '
        Me.TableLayoutPanel9.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel9.ColumnCount = 2
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 44.30894!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 55.69106!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 87.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 153.0!))
        Me.TableLayoutPanel9.Controls.Add(Me.txtTIN, 1, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.Label18, 0, 0)
        Me.TableLayoutPanel9.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel9.Location = New System.Drawing.Point(2, 422)
        Me.TableLayoutPanel9.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel9.Name = "TableLayoutPanel9"
        Me.TableLayoutPanel9.RowCount = 1
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel9.Size = New System.Drawing.Size(495, 32)
        Me.TableLayoutPanel9.TabIndex = 8
        Me.TableLayoutPanel9.TabStop = True
        '
        'txtTIN
        '
        Me.txtTIN.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTIN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTIN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTIN.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtTIN.Location = New System.Drawing.Point(225, 6)
        Me.txtTIN.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtTIN.MaxLength = 11
        Me.txtTIN.Name = "txtTIN"
        Me.txtTIN.Size = New System.Drawing.Size(263, 20)
        Me.txtTIN.TabIndex = 0
        '
        'Label18
        '
        Me.Label18.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label18.Location = New System.Drawing.Point(7, 3)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(207, 26)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "Deductee's TIN/UIN"
        '
        'txtDid
        '
        Me.txtDid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDid.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.txtDid.Location = New System.Drawing.Point(487, 518)
        Me.txtDid.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDid.MaxLength = 50
        Me.txtDid.Name = "txtDid"
        Me.txtDid.Size = New System.Drawing.Size(13, 21)
        Me.txtDid.TabIndex = 52
        Me.txtDid.Visible = False
        '
        'cmdSearch
        '
        Me.cmdSearch.BackgroundImage = Global.WizinTDS2018.My.Resources.Resources.Knob_Search_icon
        Me.cmdSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.cmdSearch.Location = New System.Drawing.Point(41, 539)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(75, 48)
        Me.cmdSearch.TabIndex = 53
        Me.cmdSearch.UseVisualStyleBackColor = True
        '
        'frmdeduteeTDSMST
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Lavender
        Me.ClientSize = New System.Drawing.Size(505, 609)
        Me.Controls.Add(Me.cmdSearch)
        Me.Controls.Add(Me.txtDid)
        Me.Controls.Add(Me.TableLayoutPanel9)
        Me.Controls.Add(Me.TableLayoutPanel8)
        Me.Controls.Add(Me.TableLayoutPanel7)
        Me.Controls.Add(Me.TableLayoutPanel6)
        Me.Controls.Add(Me.TableLayoutPanel5)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.TableLayoutPanel4)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.TableLayoutPanel3)
        Me.Controls.Add(Me.Dedsearch)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdcorrection)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "frmdeduteeTDSMST"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Modify Deductee Details"
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.TableLayoutPanel6.PerformLayout()
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.TableLayoutPanel7.PerformLayout()
        Me.TableLayoutPanel8.ResumeLayout(False)
        Me.TableLayoutPanel8.PerformLayout()
        Me.TableLayoutPanel9.ResumeLayout(False)
        Me.TableLayoutPanel9.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboDName As System.Windows.Forms.ComboBox
    Friend WithEvents chkinact As System.Windows.Forms.CheckBox
    Friend WithEvents txtDName As System.Windows.Forms.TextBox
    Friend WithEvents txtRef As System.Windows.Forms.TextBox
    Friend WithEvents txtDPAN As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbocat As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboCategory As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtDDesgn As System.Windows.Forms.TextBox
    Friend WithEvents cmdcorrection As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents Dedsearch As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel4 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents optCo As System.Windows.Forms.RadioButton
    Friend WithEvents optOther As System.Windows.Forms.RadioButton
    Friend WithEvents TableLayoutPanel5 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtDAdd5 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd4 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd3 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd2 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd1 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboDState As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtDPin As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel6 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel7 As TableLayoutPanel
    Friend WithEvents Label4 As Label
    Friend WithEvents CboCollNonRes As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents CboPerEstInd As ComboBox
    Friend WithEvents TableLayoutPanel8 As TableLayoutPanel
    Friend WithEvents Label13 As Label
    Friend WithEvents txtDeEmail As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents txtDePhone As TextBox
    Friend WithEvents TableLayoutPanel9 As TableLayoutPanel
    Friend WithEvents txtTIN As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txtDid As TextBox
    Friend WithEvents cmdSearch As Button
End Class
