Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmCorrection
    Dim savdata As Boolean
    Dim Pdata As String 'Previous Data in the cell
    Public cname As String
    Public Row, Col As Integer

    Private Sub gridcorrection_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridcorrection.CellContentClick

    End Sub

    Private Sub frmCorrection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Main()
        Dim fds As DataSet
        'selectedcoid = 4
        Dim R As Integer
        'gridhead()
        ' fds = FetchDataSet("SELECT DeductMst.DName As DeducteeName, DeductMst.DPan as PAN,  DeductMst.DPANRef as RefNo From DeductMst where coid= " & selectedcoid)
        fds = FetchDataSet("SELECT DeductMst.DName As DeducteeName, DeductMst.DPan as PAN,  DeductMst.DPANRef as RefNo From DeductMst where coid= " & selectedcoid)
        gridcorrection.DataSource = fds.Tables(0)

        Dim cmb, cmb1, cmb2 As New DataGridViewComboBoxColumn()
        cmb1.HeaderText = "PAN Cat."
        cmb1.MaxDropDownItems = 4
        cmb1.Items.AddRange("VALID PAN", "PANAPPLIED", "PANINVALID", "PANNOTAVBL")
        gridcorrection.Columns.Insert(1, cmb1)

        cmb2.HeaderText = "Category"
        cmb2.MaxDropDownItems = 3
        cmb2.Items.AddRange("G", "W", "S")
        gridcorrection.Columns.Insert(4, cmb2)

        cmb.HeaderText = "Type"
        ' cmb.Name = "cmb"
        cmb.MaxDropDownItems = 2
        cmb.Items.AddRange("O", "C")
        gridcorrection.Columns.Insert(5, cmb)
        gridcorrection.AutoGenerateColumns = False
        'Set Columns Count
        gridcorrection.ColumnCount = 6
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Close()
    End Sub

    Private Sub cmdexport_Click(sender As Object, e As EventArgs) Handles cmdexport.Click
        Dim i As Integer, m As Integer
        'Dim N As Single, j As Single
        'On Error GoTo excelerr
        Dim xlapp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        'Dim rs As New ADODB.Recordset
        ' Dim R As Long, c As Long
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")
        xlapp.Visible = True
        xlSheet = xlBook.Sheets("Sheet1")
        xlSheet = xlBook.ActiveSheet
        xlSheet.Name = "Export sheet"
        com() 'used for company Name

        xlSheet.Cells(1, 5) = cname & " (FY  " & FY & ") "
        xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 6)).Merge()
        xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 6)).Font.Bold = True
        xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 6)).HorizontalAlignment = HorizontalAlignment.Center

        xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 1)).Value = "Deductee Name"
        xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 2)).Value = "PAN Cat"
        xlSheet.Range(xlSheet.Cells(2, 3), xlSheet.Cells(2, 3)).Value = "PAN"
        xlSheet.Range(xlSheet.Cells(2, 4), xlSheet.Cells(2, 4)).Value = "Ref. No."
        xlSheet.Range(xlSheet.Cells(2, 5), xlSheet.Cells(2, 5)).Value = "Category"
        xlSheet.Range(xlSheet.Cells(2, 6), xlSheet.Cells(2, 6)).Value = "Type"

        For i = 0 To gridcorrection.Columns.Count - 2
            For m = 0 To gridcorrection.Rows.Count - 2
                xlSheet.Cells(m + 3, i + 1) = gridcorrection.Rows(m).Cells(i).Value.ToString()
                ' xlSheet.Range(xlSheet.Cells(m + 3, i + 1), xlSheet.Cells(m + 3, i + 1)).BorderAround()
                'N = N + 1
            Next
            'j = j + 1
            '  N = 0
        Next
        For m = 2 To xlSheet.UsedRange.Rows.Count
            For i = 1 To xlSheet.UsedRange.Columns.Count
                xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 6)).Font.Bold = True
                xlSheet.Range(xlSheet.Cells(m, i), xlSheet.Cells(m, i)).BorderAround(10)
            Next
        Next
        xlSheet.UsedRange.Cells.Columns.AutoFit()
        'xlapp.Application.Visible = True
        Exit Sub
        'excelerr:

        '        MsgBox("Cannot open Excel", vbCritical)
    End Sub

    'Private Sub gridcorrection_AfterEdit(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles gridcorrection.EditingControlShowing
    '    'ByVal Row, Col As Integer
    '    'Dim Row, Col As New Integer
    '    savdata = True
    '    If savdata = True Then
    '        UpdData(Col, Row)
    '    End If
    '    Pdata = ""
    'End Sub

    'Private Sub gridcorrection_beforeEdit(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles gridcorrection.EditingControlShowing
    '    gridcorrection.CurrentCell.Style.BackColor = Color.Yellow
    '    'If gridcorrection.Columns.Count = 0 Then gridcorrection.MaximumSize = 70
    '    'If gridcorrection.Columns.Count = 2 Then gridcorrection.EditMaxLength = 10
    '    'If gridcorrection.Columns.Count = 3 Then gridcorrection.EditMaxLength = 10
    'End Sub

    Private Sub gridcorrection_LeaveCell()

        '  gridcorrection.Cell(flexcpBackColor, gridcorrection.Rows, gridcorrection.Columns) = vbWhite
    End Sub

    'used for the update data in deductmst table



    Private Sub gridcorrection_KeyPressEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles gridcorrection.KeyPress
        'ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer
        Dim Row, Col, KeyAscii As New Integer
        savdata = True
        KeyAscii = CtrlKeyPress(gridcorrection, KeyAscii, MyKeypressEnum.KeypressUpperCase)
        If KeyAscii = Asc("/") Or KeyAscii = Asc("\") Or KeyAscii = Asc("'") Or KeyAscii = Asc(",") Or KeyAscii = Asc(".") Or KeyAscii = Asc(";") Or KeyAscii = Asc(":") Or KeyAscii = Asc("<") Or KeyAscii = Asc(">") Or KeyAscii = Asc("=") Or KeyAscii = Asc("-") Or KeyAscii = Asc("[") Or KeyAscii = Asc("]") Or KeyAscii = Asc("*") Or KeyAscii = Asc("+") Or KeyAscii = Asc("!") Or KeyAscii = Asc("~") Or KeyAscii = Asc("`") Or KeyAscii = Asc("@") Or KeyAscii = Asc("#") Or KeyAscii = Asc("$") Or KeyAscii = Asc("%") Or KeyAscii = Asc("^") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = Asc("|") Or KeyAscii = Asc("{") Or KeyAscii = Asc("}") Or KeyAscii = Asc("?") Or KeyAscii = Asc("_") Or KeyAscii = 32 Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub gridcorrection_KeyUpEdit()
        Dim Row, Col, KeyCode, Shift As Integer
        If KeyCode = 46 Then
            savdata = True
        End If
    End Sub

    Public Sub UpdData(clno As Integer, rno As Integer)
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        Dim sql As String
        Dim da As New OleDbDataAdapter
        Dim fds As New DataSet
        Dim dt As New DataTable
        'Dim i As Integer
        Dim PAN As String
        With gridcorrection
            ' On Error GoTo v
            'cn.Open()
            cmd.Connection = cn
            transaction = cn.BeginTransaction()
            cmd.Transaction = transaction
            ' cmd.CommandText = sql
            If clno = 0 Then
                fds.Tables.Add(dt)
                sql = "update DeductMst set DName='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString() & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value.ToString() & ""
                cmd.CommandText = sql
            ElseIf clno = 1 Then
                If gridcorrection.Rows(rno).Cells(1).Value = "0" Then
                    sql = "update DeductMst set DPan='" & gridcorrection.Rows(rno).Cells(clno + 1).Value & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanCat='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                ElseIf gridcorrection.Rows(rno).Cells(1).Value = "1" Then
                    PAN = "PANAPPLIED"
                    gridcorrection.Rows(rno).Cells(2).Value = ""
                    sql = "update DeductMst set DPan='" & PAN & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanref='" & IIf(gridcorrection.Rows(rno).Cells(clno + 2).Value.ToString = "", " ", gridcorrection.Rows(rno).Cells(clno + 2).Value.ToString) & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanCat='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                ElseIf gridcorrection.Rows(rno).Cells(1).Value = "2" Then
                    PAN = "PANINVALID"
                    gridcorrection.Rows(rno).Cells(2).Value = ""
                    sql = "update DeductMst set DPan='" & PAN & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanref='" & IIf(gridcorrection.Rows(rno).Cells(clno + 2).Value.ToString = "", " ", gridcorrection.Rows(rno).Cells(clno + 2).Value.ToString) & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanCat='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                ElseIf gridcorrection.Rows(rno).Cells(1).Value = "3" Then
                    PAN = "PANNOTAVBL"
                    gridcorrection.Rows(rno).Cells(2).Value = ""
                    sql = "update DeductMst set DPan='" & PAN & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanref='" & IIf(gridcorrection.Rows(rno).Cells(clno + 2).Value.ToString = "", " ", gridcorrection.Rows(rno).Cells(clno + 2).Value.ToString) & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    sql = "update DeductMst set DPanCat='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                End If
            ElseIf clno = 2 Then
                If gridcorrection.Rows(rno).Cells(1).Value = "0" Or gridcorrection.Rows(rno).Cells(1).Value = "VALID PAN" Then
                    If gridcorrection.Rows(rno).Cells(clno).Value = "" Then
                        MsgBox("Enter the PAN No.")
                    End If
                    sql = "update DeductMst set DPan='" & gridcorrection.Rows(rno).Cells(2).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    ' gridcorrection.Col = 4
                    cmd.CommandText = sql
                Else
                    gridcorrection.Rows(rno).Cells(clno).Value = ""
                    If gridcorrection.Rows(rno).Cells(clno - 1).Value.ToString = "1" Or gridcorrection.Rows(rno).Cells(clno - 1).Value.ToString = "PANAPPLIED" Then
                        PAN = "PANAPPLIED"
                    ElseIf gridcorrection.Rows(rno).Cells(clno - 1).Value.ToString = "2" Or gridcorrection.Rows(rno).Cells(clno - 1).Value.ToString = "PANINVALID" Then
                        PAN = "PANINVALID"
                    ElseIf gridcorrection.Rows(rno).Cells(clno - 1).Value.ToString = "3" Or gridcorrection.Rows(rno).Cells(clno - 1).Value.ToString = "PANNOTAVBL" Then
                        PAN = "PANNOTAVBL"
                    End If
                    sql = "update DeductMst set DPan='" & PAN & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                    MsgBox("Deductee not having VALID PAN !")

                End If
            ElseIf clno = 3 Then
                If gridcorrection.Rows(rno).Cells(1).Value <> "0" And gridcorrection.Rows(rno).Cells(1).Value <> "VALID PAN" Then
                    If gridcorrection.Rows(rno).Cells(clno).Value = "" Then
                        MsgBox("Please Enter the Ref.No!")
                        '.Col = 3
                    End If
                    sql = "update DeductMst set DPanref='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    cmd.CommandText = sql
                Else
                    sql = "update DeductMst set DPanref='" & gridcorrection.Rows(rno).Cells(clno).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                    ' gridcorrection.Col = 4
                    cmd.CommandText = sql

                    If gridcorrection.Rows(rno).Cells(clno).Value.ToString <> "" Then
                        MsgBox("Deductee having VALID PAN !")
                    End If
                    'cn.Close()
                End If
            ElseIf clno = 4 Then
                sql = "update DeductMst set Category='" & gridcorrection.Rows(rno).Cells(4).Value.ToString & "' where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                cmd.CommandText = sql
            ElseIf clno = 5 Then
                sql = "update DeductMst set DType='" & gridcorrection.Rows(rno).Cells(4).Value.ToString & " 'where DId=" & gridcorrection.Rows(rno).Cells(6).Value & ""
                cmd.CommandText = sql
                ' cn.Close()
            End If
            Try
                cmd.ExecuteNonQuery()
                transaction.Commit()
            Catch ex As Exception
                transaction.Rollback()
                MessageBox.Show(ex.Message) 'Error MEssage
            End Try
        End With
    End Sub

    'Private Sub gridcorrection_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles gridcorrection.CellValidating
    '    'code added by nitin to check duplicate pan nos. on 06-06-2007
    '    Dim Row, Col As New Integer
    '    Dim Cancel As Boolean
    '    Dim rs As New DataSet
    '    rs = FetchDataSet("Select DPAN from Deductmst WHERE DID <> " & gridcorrection.Rows(Row).Cells(6).Value & " and CoID=" & selectedcoid & " and DPAN not in ('PANNOTAVBL','PANINVALID','PANAPPLIED')")
    '    ' While Not (rs.BOF Or rs.EOF)
    '    If gridcorrection.EditingControl.Text = rs.Tables(0).ToString Then
    '        If MsgBox("This PAN is already mentioned for another deductee" & vbCrLf &
    '                   "Do you want to still save with this PAN No", vbCritical + vbYesNo,
    '                   "WARNING!! SAVE WITH DUPLICATE PAN") = vbYes Then
    '            Cancel = False
    '        Else
    '            Cancel = True
    '            Exit Sub
    '        End If
    '    End If
    '    'End While
    '    ' end of code added on 06-06-07
    '    Pdata = gridcorrection.Rows(Row).Cells(Col).Value
    'End Sub

    Private Sub com()

        Dim fds As New DataSet
        fds = FetchDataSet("select coName from CoMst where CoId=" & selectedcoid)
        ' rs.Open sql, Cnn
        If fds.Tables(0).Rows.Count > 0 Then
            cname = fds.ToString
        End If
    End Sub

    Private Sub gridcorrection_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles gridcorrection.CellEndEdit
        'Dim Cancel As Boolean
        'Dim rs As New DataSet
        'rs = FetchDataSet("Select DPAN from Deductmst WHERE DID <> " & gridcorrection.Rows(e.RowIndex).Cells(6).Value & " and CoID=" & selectedcoid & " and DPAN not in ('PANNOTAVBL','PANINVALID','PANAPPLIED')")
        '' While Not (rs.BOF Or rs.EOF)
        'For Each dTable As DataTable In rs.Tables
        '    For Each dRow As DataRow In dTable.Rows
        '        If gridcorrection.EditingControl.Text = rs.Tables.ToString Then
        '            If MsgBox("This PAN is already mentioned for another deductee" & vbCrLf &
        '                       "Do you want to still save with this PAN No", vbCritical + vbYesNo,
        '                       "WARNING!! SAVE WITH DUPLICATE PAN") = vbYes Then
        '                Cancel = False
        '            Else
        '                Cancel = True
        '                Exit Sub
        '            End If
        '        End If
        '        'End While
        '        ' end of code added on 06-06-07
        '        Pdata = gridcorrection.Rows(Row).Cells(Col).Value
        '    Next
        'Next
        savdata = True
        If savdata = True Then
            'cn.Close()
            UpdData(e.ColumnIndex, e.RowIndex)
        End If
        Pdata = ""
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub gridcorrection_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles gridcorrection.CellBeginEdit
        ' Dim a As 
        ' gridcorrection.CurrentCell = (a, s )
        'UpdData(e.ColumnIndex, e.RowIndex)

    End Sub
End Class