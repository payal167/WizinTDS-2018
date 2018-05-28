<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOnLineChallan
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
        Me.brwWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.SuspendLayout()
        '
        'brwWebBrowser
        '
        Me.brwWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.brwWebBrowser.Location = New System.Drawing.Point(0, 0)
        Me.brwWebBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.brwWebBrowser.Name = "brwWebBrowser"
        Me.brwWebBrowser.Size = New System.Drawing.Size(578, 408)
        Me.brwWebBrowser.TabIndex = 0
        '
        'frmOnLineChallan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(578, 408)
        Me.Controls.Add(Me.brwWebBrowser)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmOnLineChallan"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WizinTDS - Online Challan Utility"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents brwWebBrowser As System.Windows.Forms.WebBrowser
End Class
