<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImport
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.cmdok = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lvwImport = New System.Windows.Forms.ListView()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 339)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(430, 18)
        Me.ProgressBar1.TabIndex = 1
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 374)
        Me.Splitter1.TabIndex = 2
        Me.Splitter1.TabStop = False
        '
        'cmdok
        '
        Me.cmdok.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdok.ForeColor = System.Drawing.Color.Black
        Me.cmdok.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdok.Location = New System.Drawing.Point(123, 304)
        Me.cmdok.Name = "cmdok"
        Me.cmdok.Size = New System.Drawing.Size(82, 29)
        Me.cmdok.TabIndex = 3
        Me.cmdok.Text = "Import"
        Me.cmdok.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdok.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Black
        Me.cmdExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdExit.Location = New System.Drawing.Point(211, 304)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(82, 29)
        Me.cmdExit.TabIndex = 5
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'lvwImport
        '
        Me.lvwImport.CheckBoxes = True
        Me.lvwImport.Location = New System.Drawing.Point(0, 0)
        Me.lvwImport.Name = "lvwImport"
        Me.lvwImport.Size = New System.Drawing.Size(428, 269)
        Me.lvwImport.TabIndex = 6
        Me.lvwImport.UseCompatibleStateImageBehavior = False
        '
        'frmImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 374)
        Me.Controls.Add(Me.lvwImport)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdok)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmImport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmImport"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents cmdok As System.Windows.Forms.Button
    Friend WithEvents cmdExit As Button
    Friend WithEvents lvwImport As ListView
End Class
