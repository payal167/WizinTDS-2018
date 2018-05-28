<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmform16Parametre
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lvwAllowances = New System.Windows.Forms.ListView()
        Me.txtAllowances = New System.Windows.Forms.TextBox()
        Me.cmdAddAllowance = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.lvw16otherIncome = New System.Windows.Forms.ListView()
        Me.txtOthInc = New System.Windows.Forms.TextBox()
        Me.cmdAddOthInc = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.lvw1680c = New System.Windows.Forms.ListView()
        Me.txt80CCE = New System.Windows.Forms.TextBox()
        Me.cmdAdd80CCE = New System.Windows.Forms.Button()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.lvw1680CCF = New System.Windows.Forms.ListView()
        Me.txt80CCF = New System.Windows.Forms.TextBox()
        Me.cmdAdd80CCF = New System.Windows.Forms.Button()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.lvw1680CCG = New System.Windows.Forms.ListView()
        Me.txt80CCG = New System.Windows.Forms.TextBox()
        Me.cmdAdd80CCG = New System.Windows.Forms.Button()
        Me.tab6 = New System.Windows.Forms.TabPage()
        Me.lvw16OtherIVA = New System.Windows.Forms.ListView()
        Me.txtChap6a = New System.Windows.Forms.TextBox()
        Me.cmdChp6aAdd = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.tab6.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TabControl1)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(516, 226)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Items to be shown in Respective Table in Form 16"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.tab6)
        Me.TabControl1.Location = New System.Drawing.Point(6, 19)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(504, 201)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage1.Controls.Add(Me.lvwAllowances)
        Me.TabPage1.Controls.Add(Me.txtAllowances)
        Me.TabPage1.Controls.Add(Me.cmdAddAllowance)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(496, 175)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Allowances"
        '
        'lvwAllowances
        '
        Me.lvwAllowances.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwAllowances.AutoArrange = False
        Me.lvwAllowances.CheckBoxes = True
        Me.lvwAllowances.Location = New System.Drawing.Point(7, 37)
        Me.lvwAllowances.Name = "lvwAllowances"
        Me.lvwAllowances.Size = New System.Drawing.Size(317, 132)
        Me.lvwAllowances.TabIndex = 3
        Me.lvwAllowances.UseCompatibleStateImageBehavior = False
        '
        'txtAllowances
        '
        Me.txtAllowances.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAllowances.Location = New System.Drawing.Point(7, 11)
        Me.txtAllowances.Name = "txtAllowances"
        Me.txtAllowances.Size = New System.Drawing.Size(237, 20)
        Me.txtAllowances.TabIndex = 0
        '
        'cmdAddAllowance
        '
        Me.cmdAddAllowance.Location = New System.Drawing.Point(249, 8)
        Me.cmdAddAllowance.Name = "cmdAddAllowance"
        Me.cmdAddAllowance.Size = New System.Drawing.Size(75, 23)
        Me.cmdAddAllowance.TabIndex = 1
        Me.cmdAddAllowance.Text = "Add"
        Me.cmdAddAllowance.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage2.Controls.Add(Me.lvw16otherIncome)
        Me.TabPage2.Controls.Add(Me.txtOthInc)
        Me.TabPage2.Controls.Add(Me.cmdAddOthInc)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(496, 175)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Other Income"
        '
        'lvw16otherIncome
        '
        Me.lvw16otherIncome.CheckBoxes = True
        Me.lvw16otherIncome.Location = New System.Drawing.Point(6, 37)
        Me.lvw16otherIncome.Name = "lvw16otherIncome"
        Me.lvw16otherIncome.Size = New System.Drawing.Size(318, 132)
        Me.lvw16otherIncome.TabIndex = 2
        Me.lvw16otherIncome.UseCompatibleStateImageBehavior = False
        '
        'txtOthInc
        '
        Me.txtOthInc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOthInc.Location = New System.Drawing.Point(6, 11)
        Me.txtOthInc.Name = "txtOthInc"
        Me.txtOthInc.Size = New System.Drawing.Size(237, 20)
        Me.txtOthInc.TabIndex = 0
        '
        'cmdAddOthInc
        '
        Me.cmdAddOthInc.Location = New System.Drawing.Point(249, 8)
        Me.cmdAddOthInc.Name = "cmdAddOthInc"
        Me.cmdAddOthInc.Size = New System.Drawing.Size(75, 23)
        Me.cmdAddOthInc.TabIndex = 1
        Me.cmdAddOthInc.Text = "Add"
        Me.cmdAddOthInc.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage3.Controls.Add(Me.lvw1680c)
        Me.TabPage3.Controls.Add(Me.txt80CCE)
        Me.TabPage3.Controls.Add(Me.cmdAdd80CCE)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(496, 175)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Section 80CCE"
        '
        'lvw1680c
        '
        Me.lvw1680c.CheckBoxes = True
        Me.lvw1680c.Location = New System.Drawing.Point(6, 37)
        Me.lvw1680c.Name = "lvw1680c"
        Me.lvw1680c.Size = New System.Drawing.Size(318, 132)
        Me.lvw1680c.TabIndex = 2
        Me.lvw1680c.UseCompatibleStateImageBehavior = False
        '
        'txt80CCE
        '
        Me.txt80CCE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt80CCE.Location = New System.Drawing.Point(6, 11)
        Me.txt80CCE.Name = "txt80CCE"
        Me.txt80CCE.Size = New System.Drawing.Size(237, 20)
        Me.txt80CCE.TabIndex = 0
        '
        'cmdAdd80CCE
        '
        Me.cmdAdd80CCE.Location = New System.Drawing.Point(249, 8)
        Me.cmdAdd80CCE.Name = "cmdAdd80CCE"
        Me.cmdAdd80CCE.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd80CCE.TabIndex = 1
        Me.cmdAdd80CCE.Text = "Add"
        Me.cmdAdd80CCE.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage4.Controls.Add(Me.lvw1680CCF)
        Me.TabPage4.Controls.Add(Me.txt80CCF)
        Me.TabPage4.Controls.Add(Me.cmdAdd80CCF)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(496, 175)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Section 80CCF"
        '
        'lvw1680CCF
        '
        Me.lvw1680CCF.CheckBoxes = True
        Me.lvw1680CCF.Location = New System.Drawing.Point(6, 37)
        Me.lvw1680CCF.Name = "lvw1680CCF"
        Me.lvw1680CCF.Size = New System.Drawing.Size(318, 132)
        Me.lvw1680CCF.TabIndex = 2
        Me.lvw1680CCF.UseCompatibleStateImageBehavior = False
        '
        'txt80CCF
        '
        Me.txt80CCF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt80CCF.Location = New System.Drawing.Point(6, 11)
        Me.txt80CCF.Name = "txt80CCF"
        Me.txt80CCF.Size = New System.Drawing.Size(237, 20)
        Me.txt80CCF.TabIndex = 0
        '
        'cmdAdd80CCF
        '
        Me.cmdAdd80CCF.Location = New System.Drawing.Point(249, 8)
        Me.cmdAdd80CCF.Name = "cmdAdd80CCF"
        Me.cmdAdd80CCF.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd80CCF.TabIndex = 1
        Me.cmdAdd80CCF.Text = "Add"
        Me.cmdAdd80CCF.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage5.Controls.Add(Me.lvw1680CCG)
        Me.TabPage5.Controls.Add(Me.txt80CCG)
        Me.TabPage5.Controls.Add(Me.cmdAdd80CCG)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(496, 175)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Section 80CCG"
        '
        'lvw1680CCG
        '
        Me.lvw1680CCG.CheckBoxes = True
        Me.lvw1680CCG.Location = New System.Drawing.Point(6, 37)
        Me.lvw1680CCG.Name = "lvw1680CCG"
        Me.lvw1680CCG.Size = New System.Drawing.Size(318, 132)
        Me.lvw1680CCG.TabIndex = 2
        Me.lvw1680CCG.UseCompatibleStateImageBehavior = False
        '
        'txt80CCG
        '
        Me.txt80CCG.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt80CCG.Location = New System.Drawing.Point(6, 11)
        Me.txt80CCG.Name = "txt80CCG"
        Me.txt80CCG.Size = New System.Drawing.Size(237, 20)
        Me.txt80CCG.TabIndex = 0
        '
        'cmdAdd80CCG
        '
        Me.cmdAdd80CCG.Location = New System.Drawing.Point(249, 8)
        Me.cmdAdd80CCG.Name = "cmdAdd80CCG"
        Me.cmdAdd80CCG.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd80CCG.TabIndex = 1
        Me.cmdAdd80CCG.Text = "Add"
        Me.cmdAdd80CCG.UseVisualStyleBackColor = True
        '
        'tab6
        '
        Me.tab6.BackColor = System.Drawing.Color.Gainsboro
        Me.tab6.Controls.Add(Me.lvw16OtherIVA)
        Me.tab6.Controls.Add(Me.txtChap6a)
        Me.tab6.Controls.Add(Me.cmdChp6aAdd)
        Me.tab6.Location = New System.Drawing.Point(4, 22)
        Me.tab6.Name = "tab6"
        Me.tab6.Padding = New System.Windows.Forms.Padding(3)
        Me.tab6.Size = New System.Drawing.Size(496, 175)
        Me.tab6.TabIndex = 5
        Me.tab6.Text = "Other Chapter VI-A"
        '
        'lvw16OtherIVA
        '
        Me.lvw16OtherIVA.CheckBoxes = True
        Me.lvw16OtherIVA.Location = New System.Drawing.Point(6, 32)
        Me.lvw16OtherIVA.Name = "lvw16OtherIVA"
        Me.lvw16OtherIVA.Size = New System.Drawing.Size(465, 137)
        Me.lvw16OtherIVA.TabIndex = 2
        Me.lvw16OtherIVA.UseCompatibleStateImageBehavior = False
        '
        'txtChap6a
        '
        Me.txtChap6a.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtChap6a.Location = New System.Drawing.Point(6, 6)
        Me.txtChap6a.Name = "txtChap6a"
        Me.txtChap6a.Size = New System.Drawing.Size(237, 20)
        Me.txtChap6a.TabIndex = 0
        '
        'cmdChp6aAdd
        '
        Me.cmdChp6aAdd.Location = New System.Drawing.Point(249, 3)
        Me.cmdChp6aAdd.Name = "cmdChp6aAdd"
        Me.cmdChp6aAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdChp6aAdd.TabIndex = 1
        Me.cmdChp6aAdd.Text = "Add"
        Me.cmdChp6aAdd.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(193, 228)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 1
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmform16Parametre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(528, 252)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Name = "frmform16Parametre"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form 16 Parameter Form"
        Me.GroupBox1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.tab6.ResumeLayout(False)
        Me.tab6.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents txtAllowances As System.Windows.Forms.TextBox
    Friend WithEvents cmdAddAllowance As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents txtOthInc As System.Windows.Forms.TextBox
    Friend WithEvents cmdAddOthInc As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents txt80CCE As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd80CCE As System.Windows.Forms.Button
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents txt80CCF As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd80CCF As System.Windows.Forms.Button
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents txt80CCG As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd80CCG As System.Windows.Forms.Button
    Friend WithEvents tab6 As System.Windows.Forms.TabPage
    Friend WithEvents txtChap6a As System.Windows.Forms.TextBox
    Friend WithEvents cmdChp6aAdd As System.Windows.Forms.Button
    Friend WithEvents lvwAllowances As ListView
    Friend WithEvents lvw16otherIncome As ListView
    Friend WithEvents lvw1680c As ListView
    Friend WithEvents lvw1680CCF As ListView
    Friend WithEvents lvw1680CCG As ListView
    Friend WithEvents lvw16OtherIVA As ListView
    Friend WithEvents cmdClose As Button
End Class
