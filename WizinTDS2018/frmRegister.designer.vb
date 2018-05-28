<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmRegister
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRegister))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCDKey = New System.Windows.Forms.TextBox()
        Me.txtSiteKey = New System.Windows.Forms.TextBox()
        Me.txtRegKey = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Mylock = New AxJAKLock.AxActiveLock()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtReadSite = New System.Windows.Forms.TextBox()
        Me.txtReadCD = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmdRegisterLater = New System.Windows.Forms.Button()
        Me.cmdregister = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Mylock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(6, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(115, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Your Site Key is:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label2.Location = New System.Drawing.Point(6, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Enter CD &Key"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label3.Location = New System.Drawing.Point(6, 62)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(155, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Enter &Registration Key"
        '
        'txtCDKey
        '
        Me.txtCDKey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCDKey.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCDKey.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCDKey.Location = New System.Drawing.Point(179, 6)
        Me.txtCDKey.MaxLength = 16
        Me.txtCDKey.Name = "txtCDKey"
        Me.txtCDKey.Size = New System.Drawing.Size(178, 20)
        Me.txtCDKey.TabIndex = 0
        '
        'txtSiteKey
        '
        Me.txtSiteKey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSiteKey.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSiteKey.Enabled = False
        Me.txtSiteKey.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSiteKey.Location = New System.Drawing.Point(179, 36)
        Me.txtSiteKey.Name = "txtSiteKey"
        Me.txtSiteKey.Size = New System.Drawing.Size(178, 20)
        Me.txtSiteKey.TabIndex = 1
        '
        'txtRegKey
        '
        Me.txtRegKey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRegKey.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRegKey.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegKey.Location = New System.Drawing.Point(179, 65)
        Me.txtRegKey.MaxLength = 16
        Me.txtRegKey.Name = "txtRegKey"
        Me.txtRegKey.Size = New System.Drawing.Size(178, 20)
        Me.txtRegKey.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label4.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label4.Location = New System.Drawing.Point(4, 239)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(611, 41)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = resources.GetString("Label4.Text")
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Lavender
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Controls.Add(Me.Mylock)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtReadSite)
        Me.GroupBox1.Controls.Add(Me.txtReadCD)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel1)
        Me.GroupBox1.Controls.Add(Me.cmdRegisterLater)
        Me.GroupBox1.Controls.Add(Me.cmdregister)
        Me.GroupBox1.Font = New System.Drawing.Font("Algerian", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.GroupBox1.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(611, 233)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.WizinTDS2018.My.Resources.Resources.Aluminum_Background_8_625x419
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(140, 25)
        Me.PictureBox1.TabIndex = 99
        Me.PictureBox1.TabStop = False
        '
        'Mylock
        '
        Me.Mylock.Enabled = True
        Me.Mylock.Location = New System.Drawing.Point(8, 152)
        Me.Mylock.Name = "Mylock"
        Me.Mylock.OcxState = CType(resources.GetObject("Mylock.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Mylock.Size = New System.Drawing.Size(32, 31)
        Me.Mylock.TabIndex = 15
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Button1.Location = New System.Drawing.Point(273, 171)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 28)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "&Cancel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(393, 133)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(161, 16)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Read your Site Key as -"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(393, 37)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(153, 16)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Read your CD Key as -"
        '
        'txtReadSite
        '
        Me.txtReadSite.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtReadSite.Enabled = False
        Me.txtReadSite.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReadSite.Location = New System.Drawing.Point(396, 152)
        Me.txtReadSite.MaxLength = 16
        Me.txtReadSite.Multiline = True
        Me.txtReadSite.Name = "txtReadSite"
        Me.txtReadSite.Size = New System.Drawing.Size(178, 74)
        Me.txtReadSite.TabIndex = 11
        Me.txtReadSite.TabStop = False
        '
        'txtReadCD
        '
        Me.txtReadCD.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtReadCD.Enabled = False
        Me.txtReadCD.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReadCD.Location = New System.Drawing.Point(396, 56)
        Me.txtReadCD.MaxLength = 16
        Me.txtReadCD.Multiline = True
        Me.txtReadCD.Name = "txtReadCD"
        Me.txtReadCD.Size = New System.Drawing.Size(178, 74)
        Me.txtReadCD.TabIndex = 10
        Me.txtReadCD.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Algerian", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(168, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(264, 21)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Software Registration"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.05556!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.94444!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtRegKey, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSiteKey, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCDKey, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(8, 53)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.84746!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.15254!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(363, 91)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'cmdRegisterLater
        '
        Me.cmdRegisterLater.Font = New System.Drawing.Font("Tahoma", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRegisterLater.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.cmdRegisterLater.Location = New System.Drawing.Point(163, 171)
        Me.cmdRegisterLater.Name = "cmdRegisterLater"
        Me.cmdRegisterLater.Size = New System.Drawing.Size(104, 27)
        Me.cmdRegisterLater.TabIndex = 3
        Me.cmdRegisterLater.TabStop = False
        Me.cmdRegisterLater.Text = "Register &Later"
        Me.cmdRegisterLater.UseVisualStyleBackColor = True
        '
        'cmdregister
        '
        Me.cmdregister.Font = New System.Drawing.Font("Tahoma", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdregister.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.cmdregister.Location = New System.Drawing.Point(82, 171)
        Me.cmdregister.Name = "cmdregister"
        Me.cmdregister.Size = New System.Drawing.Size(75, 28)
        Me.cmdregister.TabIndex = 2
        Me.cmdregister.Text = "&Register"
        Me.cmdregister.UseVisualStyleBackColor = True
        '
        'frmRegister
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(615, 282)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "frmRegister"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Registration of Wizin-TDS"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Mylock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCDKey As System.Windows.Forms.TextBox
    Friend WithEvents txtSiteKey As System.Windows.Forms.TextBox
    Friend WithEvents txtRegKey As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRegisterLater As System.Windows.Forms.Button
    Friend WithEvents cmdregister As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    'Friend WithEvents MyLock As JAKLock.ActiveLock
    Friend WithEvents Label5 As Label
    Friend WithEvents txtReadCD As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents txtReadSite As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Mylock As AxJAKLock.AxActiveLock
    Friend WithEvents PictureBox1 As PictureBox
    'Friend WithEvents MyLock1 As AxJAKLock.AxActiveLock
    'Friend WithEvents MYLOCK As AxJAKLock.AxActiveLock
End Class
