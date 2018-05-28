<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOnlineTanLogin
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
        Me.brwWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optQtr4 = New System.Windows.Forms.RadioButton()
        Me.optQtr3 = New System.Windows.Forms.RadioButton()
        Me.optQtr2 = New System.Windows.Forms.RadioButton()
        Me.optQtr1 = New System.Windows.Forms.RadioButton()
        Me.cboFormNo = New System.Windows.Forms.ComboBox()
        Me.timTimer = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'brwWebBrowser
        '
        Me.brwWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.brwWebBrowser.Location = New System.Drawing.Point(0, 0)
        Me.brwWebBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.brwWebBrowser.Name = "brwWebBrowser"
        Me.brwWebBrowser.Size = New System.Drawing.Size(747, 537)
        Me.brwWebBrowser.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optQtr4)
        Me.GroupBox1.Controls.Add(Me.optQtr3)
        Me.GroupBox1.Controls.Add(Me.optQtr2)
        Me.GroupBox1.Controls.Add(Me.optQtr1)
        Me.GroupBox1.Location = New System.Drawing.Point(291, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 48)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Quarter"
        Me.GroupBox1.Visible = False
        '
        'optQtr4
        '
        Me.optQtr4.AutoSize = True
        Me.optQtr4.Location = New System.Drawing.Point(237, 19)
        Me.optQtr4.Name = "optQtr4"
        Me.optQtr4.Size = New System.Drawing.Size(69, 17)
        Me.optQtr4.TabIndex = 3
        Me.optQtr4.TabStop = True
        Me.optQtr4.Text = "Quarter 4"
        Me.optQtr4.UseVisualStyleBackColor = True
        '
        'optQtr3
        '
        Me.optQtr3.AutoSize = True
        Me.optQtr3.Location = New System.Drawing.Point(166, 19)
        Me.optQtr3.Name = "optQtr3"
        Me.optQtr3.Size = New System.Drawing.Size(69, 17)
        Me.optQtr3.TabIndex = 2
        Me.optQtr3.TabStop = True
        Me.optQtr3.Text = "Quarter 3"
        Me.optQtr3.UseVisualStyleBackColor = True
        '
        'optQtr2
        '
        Me.optQtr2.AutoSize = True
        Me.optQtr2.Location = New System.Drawing.Point(91, 19)
        Me.optQtr2.Name = "optQtr2"
        Me.optQtr2.Size = New System.Drawing.Size(69, 17)
        Me.optQtr2.TabIndex = 1
        Me.optQtr2.TabStop = True
        Me.optQtr2.Text = "Quarter 2"
        Me.optQtr2.UseVisualStyleBackColor = True
        '
        'optQtr1
        '
        Me.optQtr1.AutoSize = True
        Me.optQtr1.Location = New System.Drawing.Point(16, 20)
        Me.optQtr1.Name = "optQtr1"
        Me.optQtr1.Size = New System.Drawing.Size(69, 17)
        Me.optQtr1.TabIndex = 0
        Me.optQtr1.TabStop = True
        Me.optQtr1.Text = "Quarter 1"
        Me.optQtr1.UseVisualStyleBackColor = True
        '
        'cboFormNo
        '
        Me.cboFormNo.BackColor = System.Drawing.SystemColors.Window
        Me.cboFormNo.FormattingEnabled = True
        Me.cboFormNo.Location = New System.Drawing.Point(12, 27)
        Me.cboFormNo.Name = "cboFormNo"
        Me.cboFormNo.Size = New System.Drawing.Size(273, 21)
        Me.cboFormNo.TabIndex = 2
        Me.cboFormNo.Visible = False
        '
        'timTimer
        '
        Me.timTimer.Interval = 5
        '
        'frmOnlineTanLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(747, 537)
        Me.Controls.Add(Me.cboFormNo)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.brwWebBrowser)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmOnlineTanLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmOnlineTanLogin"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents brwWebBrowser As System.Windows.Forms.WebBrowser
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optQtr4 As System.Windows.Forms.RadioButton
    Friend WithEvents optQtr3 As System.Windows.Forms.RadioButton
    Friend WithEvents optQtr2 As System.Windows.Forms.RadioButton
    Friend WithEvents optQtr1 As System.Windows.Forms.RadioButton
    Friend WithEvents cboFormNo As System.Windows.Forms.ComboBox
    Friend WithEvents timTimer As Timer
End Class
