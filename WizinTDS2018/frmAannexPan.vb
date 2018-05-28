Public Class frmAannexPan
    Private Sub frmAannexPan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        getrspan()
    End Sub
    Public Sub getrspan()
        'Report.Database.SetDataSource rspan
        'CRViewer1.ReportSource = Report
        Select Case PrintRep
            Case False
                'Me.WindowState = vbMaximized
                'Me.WindowState = vbMaximized
                'CRViewer1.ViewReport
                '        While CRViewer1.IsBusy
                '            DoEvents
                '    Wend
                '    CRViewer1.Zoom 1

                'Case True
                '        Report.PrintOut
                'Report.Export
                '          If frm16A.cmbdest = "Export" Then
                '            Report.Export
                '            End If
        End Select
    End Sub
    Private Sub exportrep()
        'Dim exp As ExportOptions
        'exp = Report.ExportOptions
        ''exp.DestinationType = crEDTDiskFile
        'exp.DestinationType = crEDTDiskFile
        'exp.FormatType = crEFTExcel50
        'exp.DiskFileName = App.Path & "\default1.xls"
        'Report.Export False
    End Sub

    Private Sub Form_Resize()
        'CRViewer1.Left = 0
        '      CRViewer1.Width = ScaleWidth
        '      If ScaleHeight > frm16A.Height Then
        '        Me.CRViewer1.Height = ScaleHeight - frm16A.cmdback.Height
        '          CRViewer1.Top = frm16A.Height
        '      Else
        '        Me.CRViewer1.Height = ScaleHeight
        '        CRViewer1.Top = 0
        '      End If
        'CRViewer1.Width = ScaleWidth

        'CRViewer1.Height = ScaleHeight
        'If FrmDedList.cmbdest = "Export" Then
        '    'Me.cmdexp.Visible = True
        '    CRViewer1.Top = 450
        'Else
        '    CRViewer1.Top = 0
        'End If

    End Sub


End Class