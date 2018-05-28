<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeducteeTDS
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDName = New System.Windows.Forms.TextBox()
        Me.txtDAdd1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtDAdd5 = New System.Windows.Forms.TextBox()
        Me.txtDAdd4 = New System.Windows.Forms.TextBox()
        Me.txtDAdd3 = New System.Windows.Forms.TextBox()
        Me.txtDAdd2 = New System.Windows.Forms.TextBox()
        Me.cmdlastadd = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboDState = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtDPin = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbocat = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtDPAN = New System.Windows.Forms.TextBox()
        Me.txtref = New System.Windows.Forms.TextBox()
        Me.optOther = New System.Windows.Forms.RadioButton()
        Me.optCo = New System.Windows.Forms.RadioButton()
        Me.cboCategory = New System.Windows.Forms.ComboBox()
        Me.txtDDesgn = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.TableLayoutPanel9 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtTIN = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtDeEmail = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtDePhone = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.CboCollNonRes = New System.Windows.Forms.ComboBox()
        Me.CboPerEstInd = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel9.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 3)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 29)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Deductee's &Name"
        '
        'txtDName
        '
        Me.txtDName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDName.Enabled = False
        Me.txtDName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDName.Location = New System.Drawing.Point(186, 6)
        Me.txtDName.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDName.MaxLength = 75
        Me.txtDName.Name = "txtDName"
        Me.txtDName.Size = New System.Drawing.Size(277, 20)
        Me.txtDName.TabIndex = 0
        Me.txtDName.TabStop = False
        '
        'txtDAdd1
        '
        Me.txtDAdd1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd1.Location = New System.Drawing.Point(0, 1)
        Me.txtDAdd1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd1.MaxLength = 25
        Me.txtDAdd1.Name = "txtDAdd1"
        Me.txtDAdd1.Size = New System.Drawing.Size(277, 13)
        Me.txtDAdd1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 35)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 86)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Deductee's Addre&ss"
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
        Me.Panel1.Location = New System.Drawing.Point(186, 38)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(277, 80)
        Me.Panel1.TabIndex = 4
        '
        'txtDAdd5
        '
        Me.txtDAdd5.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd5.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd5.Location = New System.Drawing.Point(0, 61)
        Me.txtDAdd5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd5.MaxLength = 25
        Me.txtDAdd5.Name = "txtDAdd5"
        Me.txtDAdd5.Size = New System.Drawing.Size(283, 13)
        Me.txtDAdd5.TabIndex = 5
        '
        'txtDAdd4
        '
        Me.txtDAdd4.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd4.Location = New System.Drawing.Point(0, 46)
        Me.txtDAdd4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd4.MaxLength = 25
        Me.txtDAdd4.Name = "txtDAdd4"
        Me.txtDAdd4.Size = New System.Drawing.Size(282, 13)
        Me.txtDAdd4.TabIndex = 4
        '
        'txtDAdd3
        '
        Me.txtDAdd3.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd3.Location = New System.Drawing.Point(0, 31)
        Me.txtDAdd3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd3.MaxLength = 25
        Me.txtDAdd3.Name = "txtDAdd3"
        Me.txtDAdd3.Size = New System.Drawing.Size(282, 13)
        Me.txtDAdd3.TabIndex = 3
        '
        'txtDAdd2
        '
        Me.txtDAdd2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDAdd2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDAdd2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDAdd2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDAdd2.Location = New System.Drawing.Point(0, 16)
        Me.txtDAdd2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDAdd2.MaxLength = 25
        Me.txtDAdd2.Name = "txtDAdd2"
        Me.txtDAdd2.Size = New System.Drawing.Size(281, 13)
        Me.txtDAdd2.TabIndex = 2
        '
        'cmdlastadd
        '
        Me.cmdlastadd.BackColor = System.Drawing.Color.Lavender
        Me.cmdlastadd.FlatAppearance.BorderColor = System.Drawing.SystemColors.HotTrack
        Me.cmdlastadd.FlatAppearance.BorderSize = 7
        Me.cmdlastadd.Location = New System.Drawing.Point(16, 76)
        Me.cmdlastadd.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdlastadd.Name = "cmdlastadd"
        Me.cmdlastadd.Size = New System.Drawing.Size(113, 45)
        Me.cmdlastadd.TabIndex = 5
        Me.cmdlastadd.TabStop = False
        Me.cmdlastadd.Text = "To &Fetch Last Add Click Me"
        Me.cmdlastadd.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdlastadd.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 3)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(54, 29)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "State"
        '
        'cboDState
        '
        Me.cboDState.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDState.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cboDState.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboDState.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDState.FormattingEnabled = True
        Me.cboDState.Location = New System.Drawing.Point(72, 6)
        Me.cboDState.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cboDState.Name = "cboDState"
        Me.cboDState.Size = New System.Drawing.Size(113, 21)
        Me.cboDState.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(196, 3)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 29)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Pin Code"
        '
        'txtDPin
        '
        Me.txtDPin.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDPin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDPin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDPin.Location = New System.Drawing.Point(282, 6)
        Me.txtDPin.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDPin.MaxLength = 6
        Me.txtDPin.Name = "txtDPin"
        Me.txtDPin.Size = New System.Drawing.Size(180, 20)
        Me.txtDPin.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 3)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(188, 30)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Category for PAN"
        '
        'cbocat
        '
        Me.cbocat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbocat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbocat.FormattingEnabled = True
        Me.cbocat.Items.AddRange(New Object() {"VALID PAN", "PANAPPLIED", "PANINVALID", "PANNOTAVBL"})
        Me.cbocat.Location = New System.Drawing.Point(206, 6)
        Me.cbocat.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cbocat.Name = "cbocat"
        Me.cbocat.Size = New System.Drawing.Size(257, 21)
        Me.cbocat.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 36)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(169, 24)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Designation (if any)"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 3)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 27)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "PAN"
        '
        'Label8
        '
        Me.Label8.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 3)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(169, 30)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Category"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 15)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(102, 15)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "Deductee Type"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(201, 3)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 27)
        Me.Label10.TabIndex = 23
        Me.Label10.Text = "REF"
        '
        'txtDPAN
        '
        Me.txtDPAN.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDPAN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDPAN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDPAN.Location = New System.Drawing.Point(63, 6)
        Me.txtDPAN.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDPAN.MaxLength = 10
        Me.txtDPAN.Name = "txtDPAN"
        Me.txtDPAN.Size = New System.Drawing.Size(127, 20)
        Me.txtDPAN.TabIndex = 0
        '
        'txtref
        '
        Me.txtref.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtref.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtref.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtref.Location = New System.Drawing.Point(277, 6)
        Me.txtref.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtref.MaxLength = 10
        Me.txtref.Name = "txtref"
        Me.txtref.Size = New System.Drawing.Size(186, 20)
        Me.txtref.TabIndex = 1
        '
        'optOther
        '
        Me.optOther.AutoSize = True
        Me.optOther.Checked = True
        Me.optOther.Location = New System.Drawing.Point(148, 5)
        Me.optOther.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.optOther.Name = "optOther"
        Me.optOther.Size = New System.Drawing.Size(92, 34)
        Me.optOther.TabIndex = 0
        Me.optOther.TabStop = True
        Me.optOther.Text = "Other than" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Company"
        Me.optOther.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.optOther.UseVisualStyleBackColor = True
        '
        'optCo
        '
        Me.optCo.AutoSize = True
        Me.optCo.Location = New System.Drawing.Point(264, 5)
        Me.optCo.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.optCo.Name = "optCo"
        Me.optCo.Size = New System.Drawing.Size(95, 34)
        Me.optCo.TabIndex = 1
        Me.optCo.Text = "Company/" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Corporates"
        Me.optCo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.optCo.UseVisualStyleBackColor = True
        '
        'cboCategory
        '
        Me.cboCategory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCategory.FormattingEnabled = True
        Me.cboCategory.ItemHeight = 13
        Me.cboCategory.Items.AddRange(New Object() {"G - General/Other", "W - Woman Assessee", "S - Senior Citizen", "O - Super Senior Citizen"})
        Me.cboCategory.Location = New System.Drawing.Point(187, 6)
        Me.cboCategory.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(276, 21)
        Me.cboCategory.TabIndex = 0
        '
        'txtDDesgn
        '
        Me.txtDDesgn.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDDesgn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDDesgn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDDesgn.Location = New System.Drawing.Point(187, 39)
        Me.txtDDesgn.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDDesgn.MaxLength = 25
        Me.txtDDesgn.Name = "txtDDesgn"
        Me.txtDDesgn.Size = New System.Drawing.Size(276, 20)
        Me.txtDDesgn.TabIndex = 1
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38.2653!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 61.7347!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDName, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 2)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.39683!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 74.60317!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(470, 124)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 34.06593!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65.93407!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 83.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 187.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.cboDState, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label5, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.txtDPin, 3, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 127)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(469, 35)
        Me.TableLayoutPanel2.TabIndex = 1
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel3.ColumnCount = 2
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38.52041!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 61.47959!))
        Me.TableLayoutPanel3.Controls.Add(Me.Label8, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Label6, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.cboCategory, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.txtDDesgn, 1, 1)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 286)
        Me.TableLayoutPanel3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 2
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 57.14286!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 42.85714!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(470, 63)
        Me.TableLayoutPanel3.TabIndex = 5
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 42.60204!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 57.39796!))
        Me.TableLayoutPanel4.Controls.Add(Me.Label3, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.cbocat, 1, 0)
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(3, 162)
        Me.TableLayoutPanel4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(470, 36)
        Me.TableLayoutPanel4.TabIndex = 2
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel5.ColumnCount = 4
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.07882!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 71.92118!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 73.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 193.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.Label7, 0, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.txtDPAN, 1, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.Label10, 2, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.txtref, 3, 0)
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(3, 201)
        Me.TableLayoutPanel5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 1
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(470, 33)
        Me.TableLayoutPanel5.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.optCo)
        Me.Panel2.Controls.Add(Me.optOther)
        Me.Panel2.Location = New System.Drawing.Point(2, 240)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(471, 42)
        Me.Panel2.TabIndex = 4
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdSave
        '
        Me.cmdSave.Image = Global.WizinTDS2018.My.Resources.Resources.RT_SAVE1
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdSave.Location = New System.Drawing.Point(114, 544)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(84, 46)
        Me.cmdSave.TabIndex = 9
        Me.cmdSave.Text = "Save"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Image = Global.WizinTDS2018.My.Resources.Resources.Knob_Cancel_icon
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCancel.Location = New System.Drawing.Point(233, 544)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(84, 46)
        Me.cmdCancel.TabIndex = 10
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel9
        '
        Me.TableLayoutPanel9.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel9.ColumnCount = 2
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38.52041!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 61.47959!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 87.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 153.0!))
        Me.TableLayoutPanel9.Controls.Add(Me.txtTIN, 1, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.Label18, 0, 0)
        Me.TableLayoutPanel9.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel9.Location = New System.Drawing.Point(2, 413)
        Me.TableLayoutPanel9.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel9.Name = "TableLayoutPanel9"
        Me.TableLayoutPanel9.RowCount = 1
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel9.Size = New System.Drawing.Size(470, 32)
        Me.TableLayoutPanel9.TabIndex = 7
        '
        'txtTIN
        '
        Me.txtTIN.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTIN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTIN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTIN.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtTIN.Location = New System.Drawing.Point(187, 6)
        Me.txtTIN.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtTIN.MaxLength = 11
        Me.txtTIN.Name = "txtTIN"
        Me.txtTIN.Size = New System.Drawing.Size(276, 20)
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
        Me.Label18.Size = New System.Drawing.Size(169, 26)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "Deductee's TIN/UIN"
        '
        'Label13
        '
        Me.Label13.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label13.Location = New System.Drawing.Point(7, 31)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(168, 30)
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
        Me.txtDeEmail.Location = New System.Drawing.Point(186, 6)
        Me.txtDeEmail.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDeEmail.MaxLength = 50
        Me.txtDeEmail.Name = "txtDeEmail"
        Me.txtDeEmail.Size = New System.Drawing.Size(277, 20)
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
        Me.Label16.Size = New System.Drawing.Size(168, 25)
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
        Me.txtDePhone.Location = New System.Drawing.Point(186, 34)
        Me.txtDePhone.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDePhone.MaxLength = 11
        Me.txtDePhone.Name = "txtDePhone"
        Me.txtDePhone.Size = New System.Drawing.Size(277, 20)
        Me.txtDePhone.TabIndex = 1
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel7.ColumnCount = 2
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 82.6972!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.3028!))
        Me.TableLayoutPanel7.Controls.Add(Me.Label12, 0, 1)
        Me.TableLayoutPanel7.Controls.Add(Me.Label11, 0, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.CboCollNonRes, 1, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.CboPerEstInd, 1, 1)
        Me.TableLayoutPanel7.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(2, 445)
        Me.TableLayoutPanel7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 2
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.87719!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.12281!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(469, 60)
        Me.TableLayoutPanel7.TabIndex = 8
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
        Me.Label12.Size = New System.Drawing.Size(372, 26)
        Me.Label12.TabIndex = 39
        Me.Label12.Text = "Deductee is having Permanent Establishment in India"
        Me.Label12.Visible = False
        '
        'Label11
        '
        Me.Label11.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label11.Location = New System.Drawing.Point(7, 3)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(372, 25)
        Me.Label11.TabIndex = 37
        Me.Label11.Text = "Collectee is Non Resident"
        '
        'CboCollNonRes
        '
        Me.CboCollNonRes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCollNonRes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboCollNonRes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CboCollNonRes.FormattingEnabled = True
        Me.CboCollNonRes.Location = New System.Drawing.Point(390, 6)
        Me.CboCollNonRes.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CboCollNonRes.Name = "CboCollNonRes"
        Me.CboCollNonRes.Size = New System.Drawing.Size(72, 21)
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
        Me.CboPerEstInd.Location = New System.Drawing.Point(390, 34)
        Me.CboPerEstInd.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CboPerEstInd.Name = "CboPerEstInd"
        Me.CboPerEstInd.Size = New System.Drawing.Size(72, 21)
        Me.CboPerEstInd.TabIndex = 2
        Me.CboPerEstInd.Visible = False
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel6.ColumnCount = 2
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38.2653!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 61.7347!))
        Me.TableLayoutPanel6.Controls.Add(Me.txtDePhone, 1, 1)
        Me.TableLayoutPanel6.Controls.Add(Me.Label13, 0, 1)
        Me.TableLayoutPanel6.Controls.Add(Me.txtDeEmail, 1, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.Label16, 0, 0)
        Me.TableLayoutPanel6.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(2, 349)
        Me.TableLayoutPanel6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 2
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45.90164!))
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 54.09836!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(470, 64)
        Me.TableLayoutPanel6.TabIndex = 6
        '
        'frmDeducteeTDS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Lavender
        Me.ClientSize = New System.Drawing.Size(486, 599)
        Me.Controls.Add(Me.TableLayoutPanel9)
        Me.Controls.Add(Me.TableLayoutPanel7)
        Me.Controls.Add(Me.TableLayoutPanel6)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdlastadd)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.TableLayoutPanel5)
        Me.Controls.Add(Me.TableLayoutPanel4)
        Me.Controls.Add(Me.TableLayoutPanel3)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.cmdSave)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "frmDeducteeTDS"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Deductee Details"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel9.ResumeLayout(False)
        Me.TableLayoutPanel9.PerformLayout()
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.TableLayoutPanel7.PerformLayout()
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.TableLayoutPanel6.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDName As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtDAdd5 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd4 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd3 As System.Windows.Forms.TextBox
    Friend WithEvents txtDAdd2 As System.Windows.Forms.TextBox
    Friend WithEvents cmdlastadd As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboDState As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDPin As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbocat As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtDPAN As System.Windows.Forms.TextBox
    Friend WithEvents txtref As System.Windows.Forms.TextBox
    Friend WithEvents optOther As System.Windows.Forms.RadioButton
    Friend WithEvents optCo As System.Windows.Forms.RadioButton
    Friend WithEvents cboCategory As System.Windows.Forms.ComboBox
    Friend WithEvents txtDDesgn As System.Windows.Forms.TextBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel4 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel5 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdCancel As Button
    Friend WithEvents TableLayoutPanel9 As TableLayoutPanel
    Friend WithEvents txtTIN As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents TableLayoutPanel7 As TableLayoutPanel
    Friend WithEvents Label12 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents CboCollNonRes As ComboBox
    Friend WithEvents CboPerEstInd As ComboBox
    Friend WithEvents TableLayoutPanel6 As TableLayoutPanel
    Friend WithEvents txtDePhone As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txtDeEmail As TextBox
    Friend WithEvents Label16 As Label
End Class
