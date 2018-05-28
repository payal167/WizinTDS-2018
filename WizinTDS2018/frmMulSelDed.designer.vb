<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMulSelDed
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
        Me.cmbcl = New System.Windows.Forms.Button()
        Me.cmbsl = New System.Windows.Forms.Button()
        Me.cmbcancel = New System.Windows.Forms.Button()
        Me.cmdback = New System.Windows.Forms.Button()
        Me.grdbtb = New System.Windows.Forms.DataGridView()
        CType(Me.grdbtb, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbcl
        '
        Me.cmbcl.Location = New System.Drawing.Point(93, 247)
        Me.cmbcl.Name = "cmbcl"
        Me.cmbcl.Size = New System.Drawing.Size(75, 23)
        Me.cmbcl.TabIndex = 1
        Me.cmbcl.Text = "Clear All"
        Me.cmbcl.UseVisualStyleBackColor = True
        '
        'cmbsl
        '
        Me.cmbsl.Location = New System.Drawing.Point(174, 247)
        Me.cmbsl.Name = "cmbsl"
        Me.cmbsl.Size = New System.Drawing.Size(75, 23)
        Me.cmbsl.TabIndex = 2
        Me.cmbsl.Text = "Select All"
        Me.cmbsl.UseVisualStyleBackColor = True
        '
        'cmbcancel
        '
        Me.cmbcancel.Location = New System.Drawing.Point(255, 247)
        Me.cmbcancel.Name = "cmbcancel"
        Me.cmbcancel.Size = New System.Drawing.Size(75, 23)
        Me.cmbcancel.TabIndex = 3
        Me.cmbcancel.TabStop = False
        Me.cmbcancel.Text = "Cancel"
        Me.cmbcancel.UseVisualStyleBackColor = True
        '
        'cmdback
        '
        Me.cmdback.Location = New System.Drawing.Point(12, 247)
        Me.cmdback.Name = "cmdback"
        Me.cmdback.Size = New System.Drawing.Size(75, 23)
        Me.cmdback.TabIndex = 0
        Me.cmdback.Text = "Back"
        Me.cmdback.UseVisualStyleBackColor = True
        '
        'grdbtb
        '
        Me.grdbtb.AllowUserToAddRows = False
        Me.grdbtb.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdbtb.Location = New System.Drawing.Point(1, 1)
        Me.grdbtb.Name = "grdbtb"
        Me.grdbtb.Size = New System.Drawing.Size(338, 240)
        Me.grdbtb.TabIndex = 5
        '
        'frmMulSelDed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(339, 277)
        Me.Controls.Add(Me.cmbcl)
        Me.Controls.Add(Me.cmbsl)
        Me.Controls.Add(Me.cmbcancel)
        Me.Controls.Add(Me.cmdback)
        Me.Controls.Add(Me.grdbtb)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "frmMulSelDed"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select"
        CType(Me.grdbtb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmbcl As System.Windows.Forms.Button
    Friend WithEvents cmbsl As System.Windows.Forms.Button
    Friend WithEvents cmbcancel As System.Windows.Forms.Button
    Friend WithEvents cmdback As System.Windows.Forms.Button
    Friend WithEvents grdbtb As System.Windows.Forms.DataGridView
End Class
