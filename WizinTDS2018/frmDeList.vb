Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class frmDeList
    'Dim Report As New CrpdeductList
    Public Typ As String
    'Dim frmd As New frmmultselded
    Dim sqld As String
    Dim frm As New frmAannexPan
    Dim row, col As Integer

    'Dim crp As CrpdeductListsamepan
    Private Sub frmDeList_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        'frmCoMst.EnterTab(e)
    End Sub

    Private Sub cmbDesti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdest.Enter
        cmbdest.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbDesti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbdest.Leave
        cmbdest.BackColor = Color.White
    End Sub

    Private Sub cmbCompName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCNm.Enter
        cmbCNm.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbCompName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCNm.Leave
        cmbCNm.BackColor = Color.White
    End Sub

    Private Sub cmbFrmTyp_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbtyp.Leave
        cmbtyp.BackColor = Color.White
    End Sub

    Private Sub cmbFrmTyp_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbtyp.Enter
        cmbtyp.BackColor = Color.LightYellow
    End Sub


    Private Sub cmbQuar_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbQuarter.Enter
        cmbQuarter.BackColor = Color.LightYellow
    End Sub

    Private Sub cmbQuar_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbQuarter.Leave
        cmbQuarter.BackColor = Color.White
    End Sub

    Private Sub frmDeList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'cmbdest.Clear
        'Main()
        'cmbdest.Items.Add("Screen")
        ' cmbdest.Items.Add("Printer")
        cmbdest.Items.Add("Export")
        Me.cmbdest.SelectedIndex = 0
        'cmbtyp.Clear()
        cmbtyp.Items.Add("24Q")
        cmbtyp.Items.Add("26Q")
        cmbtyp.Items.Add("27EQ")
        cmbtyp.Items.Add("27Q")
        cmbtyp.SelectedIndex = 0
        FillCName()
        'If cmbCNm.Text <> 0 Then
        '    For i = 0 To cmbCNm.Text - 1
        '        If cmbCNm.ItemData(i) = selectedcoid Then
        '            cmbCNm.SelectedIndex = i
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub
    Private Sub FillCName()
        Dim ds As DataSet
        ds = FetchDataSet("Select CoId,CoName From CoMst Order By CoName")
        cmbCNm.DataSource = ds.Tables(0)
        cmbCNm.ValueMember = "CoId"
        cmbCNm.DisplayMember = "CoName"
        ds.Dispose()
        '        rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        '    If rs.RecordCount > 0 Then
        '            cmbCNm.Clear
        '            While Not rs.EOF
        '                cmbCNm.AddItem rs!CoName & ""
        '            cmbCNm.ItemData(cmbCNm.NewIndex) = rs!coid
        '                rs.MoveNext
        '        Wend
        '        cmbCNm.ListIndex = 0
        '        End If
        '        If rs.State = adStateOpen Then rs.Close
        'Set rs = Nothing
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdexit.Click
        Me.Close()
    End Sub

    Private Sub chknilrec_CheckedChanged(sender As Object, e As EventArgs) Handles chknilrec.CheckedChanged
        If chknilrec.Checked = 1 Then
            chkpan.Enabled = False
            cmbtyp.Enabled = False
            cmbQuarter.Enabled = False
            cmdc.Enabled = False
        Else
            chkpan.Enabled = True
            cmbtyp.Enabled = True
            cmbQuarter.Enabled = True
            cmdc.Enabled = True
        End If
    End Sub

    Private Sub chkpan_CheckedChanged(sender As Object, e As EventArgs) Handles chkpan.CheckedChanged
        If chkpan.Checked = 1 Then
            chknilrec.Enabled = False
        Else
            chknilrec.Enabled = True
        End If
    End Sub

    Private Sub cmbtyp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbtyp.SelectedIndexChanged
        'cn.Close()
        ' Main()
        Qtr()
    End Sub

    Private Sub cmbQuarter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbQuarter.SelectedIndexChanged
        If cmbQuarter.Text = "24Q4" Then
            chksalary.Visible = True
        Else
            chksalary.Visible = False
        End If
    End Sub

    Private Sub cmdc_Click()
        'frmd.grdbtb = ""
        Dim rs As New DataSet
        Dim sql As String

        rs = FetchDataSet("SELECT distinct RetnMst.FrmType FROM RetnMst") ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex))
        If cmbtyp.SelectedIndex = 0 Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"
        ElseIf cmbtyp.SelectedIndex = 1 Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cmbtyp.SelectedIndex = 2 Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        End If
        ' rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        cmbQuarter.SelectedIndex = -1
        'frmd.sql = sql
        'frmd.strflg = True
        'frmd.Show 'vbModal
        'sqld = frmd.sql
        'If sqld = "" Then
        '    cmbQuarter.Enabled = True
        'Else
        '    cmbQuarter.Enabled = False
        'End If
        'frmd.Hide

    End Sub
    Private Sub cmdgen_Click(sender As Object, e As EventArgs) Handles cmdgen.Click

        'xlapp.Visible = True
        'xlSheet = xlBook.Sheets("Sheet1")
        'xlSheet = xlBook.ActiveSheet
        'xlSheet.Name = "Export sheet"
        Dim cmd As New OleDbCommand
        Dim sql, sqlpan As String
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter
        'Report.DiscardSavedData
        frm = Nothing
        'If cmbQuarter.Text = "" And chknilrec.Checked <> True Then
        '    MsgBox("Please Select Quarter...")
        '    Exit Sub
        'End If

        If chknilrec.Checked = True Then
            sql = " SELECT DISTINCT CoMst.CoID, CoMst.CoName, DeductMst.DId, DeductMst.DName, DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5,StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, DeductMst.DPANRef, '' as frmtype,DeductMst.Category" _
                    & " FROM ((CoMst INNER JOIN DeductMst ON CoMst.CoID = DeductMst.CoID) INNER JOIN StateMst ON DeductMst.DState = StateMst.StateID) INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID " _
                    & " WHERE (((DeductMst.DId) Not In (select did from Deductee24Q ) And (DeductMst.DId) Not In (select did from Deductee26Q ) And (DeductMst.DId) Not In (select did from Deductee27EQ ) And (DeductMst.DId) Not In (select did from SalaryDetail24Q )))"
            cmd = New OleDbCommand(sql, cn)
            da.SelectCommand = cmd
            ds = New DataSet()
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Excel1()
            DataGridView1.Dispose()
        Else
            'If cmbQuarter.Text = "" Then
            '    MsgBox("Please Select Quarter...")
            '    Exit Sub
            'Else
            Select Case cmbtyp.SelectedIndex
                    Case 0            ' for form 24Q
                        'for only form 24 of qouter 4(Annual Salary) 
                        If cmbQuarter.Text = "24Q4" Then
                            If chksalary.Checked = 0 Then
                                sql = "SELECT DISTINCT CoMst.CoID, CoMst.CoName, SalaryDetail24Q.DId, DeductMst.DName, " _
                                 & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                                 & "StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                                 & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((SalaryDetail24Q INNER JOIN DeductMst ON SalaryDetail24Q.DId = DeductMst.DId) " _
                                 & "INNER JOIN RetnMst ON SalaryDetail24Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                                 & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                                If cmbQuarter.SelectedIndex >= 0 Then
                                    If cmbQuarter.Text = "24Q4" Then  'for Q4 of form 24Q(Annual salary)
                                        sql = sql & " where RetnMst.FrmType='24Q4'"
                                    Else
                                        sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                                    End If
                                    If cmbCNm.SelectedIndex > -1 Then
                                        sql = sql & " and CoMst.COId = " & cmbCNm.SelectedIndex + 1
                                        If chkpan.Checked = vbTrue Then
                                            sql = sql & " AND DeductMst.DPANcat in(1,2,3)"
                                        End If
                                    Else
                                        If chkpan.Checked = vbTrue Then
                                            sql = sql & " and DeductMst.DPANcat in(1,2,3)"
                                        End If
                                    End If

                                ElseIf sqld <> vbNullString Then
                                    If cmbtyp.Text = "24Q" Then   'for all qouter of form24Q with Annual Salary 
                                        ' sql = sql & "and RetnMst.FrmType in " & sqld & ""
                                        sql = sql & " where RetnMst.FrmType in " & "(" & sqld & ")"
                                    Else
                                        ' sql = sql & " where RetnMst.FrmType in " & sqld & ""
                                        sql = sql & " where RetnMst.FrmType in " & "(" & sqld & ")"
                                    End If
                                    If cmbCNm.SelectedIndex > -1 Then
                                        sql = sql & " and CoMst.COId = " & cmbCNm.SelectedIndex
                                        If chkpan.Checked = vbTrue Then
                                            sql = sql & " AND DeductMst.DPANcat In (1,2,3)"
                                        End If
                                    Else
                                        If chkpan.Checked = vbTrue Then
                                            sql = sql & " and DeductMst.DPANcat In (1,2,3)"
                                        End If
                                    End If
                                End If
                                'for all qouter of 24Q form (with Annual Salary) add by jayshree on 20/07/06

                                '            sql = " SELECT DISTINCT CoMst.CoID, CoMst.CoName, DeductMst.DId, DeductMst.DName, DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, StateMst.StateName,DeductMst.DPin, DeductMst.DState, DeductMst.DPan, DeductMst.DType, DeductMst.DPANRef, '' as frmtype,DeductMst.Category FROM ((CoMst INNER JOIN DeductMst ON CoMst.CoID = DeductMst.CoID) INNER JOIN StateMst ON DeductMst.DState = StateMst.StateID) INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID  WHERE (((DeductMst.DId)  In (select did from Deductee24Q )  or (DeductMst.DId) In (select did from SalaryDetail24Q )))"
                                '        ' for Q1,Q2,Q4 of form 24Q
                                '        Else
                                sql = sql & " union all SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee24Q.DId, DeductMst.DName, " _
                                & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                                & "StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                                & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((Deductee24Q INNER JOIN DeductMst ON Deductee24Q.DId = DeductMst.DId) " _
                                & "INNER JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                                & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                            Else
                                sql = sql & " SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee24Q.DId, DeductMst.DName, " _
                                 & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                                 & "StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                                 & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((Deductee24Q INNER JOIN DeductMst ON Deductee24Q.DId = DeductMst.DId) " _
                                 & "INNER JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                                 & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                            End If
                        ElseIf sqld <> vbNullString Then
                            sql = "SELECT DISTINCT CoMst.CoID, CoMst.CoName, SalaryDetail24Q.DId, DeductMst.DName, " _
                           & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                           & "StateMst.StateName ,deductmst.dpin ,  DeductMst.DState , DeductMst.DPan, DeductMst.DType, " _
                           & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((SalaryDetail24Q INNER JOIN DeductMst ON SalaryDetail24Q.DId = DeductMst.DId) " _
                           & "INNER JOIN RetnMst ON SalaryDetail24Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                           & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                            If cmbQuarter.SelectedIndex >= 0 Then
                                If cmbQuarter.Text = "24Q4" Then  'for Q4 of form 24Q(Annual salary)add by jayshree on 20/07/06
                                    sql = sql & " where RetnMst.FrmType='24Q4'"
                                Else
                                    sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                                End If
                                If cmbCNm.SelectedIndex > -1 Then
                                    sql = sql & " and CoMst.COId = " & cmbCNm.SelectedIndex
                                    If chkpan.Checked = vbTrue Then
                                        sql = sql & " AND DeductMst.DPANcat in(1,2,3)"
                                    End If
                                Else
                                    If chkpan.Checked = vbTrue Then
                                        sql = sql & " and DeductMst.DPANcat in(1,2,3)"
                                    End If
                                End If

                            ElseIf sqld <> vbNullString Then
                                If cmbtyp.Text = "24Q" Then   'for all qouter of form24Q with Annual Salary add by jayshree on 20/07/06
                                    ' sql = sql & " where RetnMst.FrmType in " & sqld & " "
                                    sql = sql & " where RetnMst.FrmType in " & sqld
                                Else
                                    ' sql = sql & " where RetnMst.FrmType in " & (sqld) & ""
                                    sql = sql & " where RetnMst.FrmType in " & "(" & sqld & ")"
                                    ' "(" & sql1 & ")"

                                End If
                                If cmbCNm.SelectedIndex > -1 Then
                                    sql = sql & " And CoMst.COId = " & cmbCNm.SelectedValue
                                    If chkpan.Checked = vbTrue Then
                                        sql = sql & " And DeductMst.DPANcat In (1,2,3)"
                                    End If
                                Else
                                    If chkpan.Checked = vbTrue Then
                                        sql = sql & " And DeductMst.DPANcat In (1,2,3)"
                                    End If
                                End If
                            End If
                            'for all qouter of 24Q form (with Annual Salary) add by jayshree on 20/07/06

                            '            sql = " Select DISTINCT CoMst.CoID, CoMst.CoName, DeductMst.DId, DeductMst.DName, DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, StateMst.StateName,DeductMst.DPin, DeductMst.DState, DeductMst.DPan, DeductMst.DType, DeductMst.DPANRef, '' as frmtype,DeductMst.Category FROM ((CoMst INNER JOIN DeductMst ON CoMst.CoID = DeductMst.CoID) INNER JOIN StateMst ON DeductMst.DState = StateMst.StateID) INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID  WHERE (((DeductMst.DId)  In (select did from Deductee24Q )  or (DeductMst.DId) In (select did from SalaryDetail24Q )))"
                            '        ' for Q1,Q2,Q4 of form 24Q
                            '        Else
                            sql = sql & " union all SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee24Q.DId, DeductMst.DName, " _
                            & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                            & "StateMst.StateName ,deductmst.dpin, DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                            & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((Deductee24Q INNER JOIN DeductMst ON Deductee24Q.DId = DeductMst.DId) " _
                            & "INNER JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                            & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"

                        Else
                            sql = sql & "  SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee24Q.DId, DeductMst.DName," _
                                & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5," _
                                & "StateMst.StateName ,deductmst.dpin, DeductMst.DState, DeductMst.DPan, DeductMst.DType," _
                                & "DeductMst.DPANRef, RetnMst.FrmType, DeductMst.Category " _
                                & "FROM (((Deductee24Q INNER JOIN DeductMst ON Deductee24Q.DId = DeductMst.DId) " _
                                & "INNER JOIN RetnMst ON Deductee24Q.RetnID = RetnMst.RetnID)" _
                                & "INNER JOIN StateMst ON DeductMst.DState = StateMst.StateID)" _
                                & "INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID "
                        End If
                    Case 1  ' for form 26Q
                        sql = "SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee26Q.DId, DeductMst.DName, " _
                            & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                            & "StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                            & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((Deductee26Q INNER JOIN DeductMst ON Deductee26Q.DId = DeductMst.DId) " _
                            & "INNER JOIN RetnMst ON Deductee26Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                            & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                    Case 2 ' for form 27EQ
                        sql = "SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee27EQ.DId, DeductMst.DName, " _
                          & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                          & "StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                          & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((Deductee27EQ INNER JOIN DeductMst ON Deductee27EQ.DId = DeductMst.DId) " _
                          & "INNER JOIN RetnMst ON Deductee27EQ.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                          & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                    Case 3 ' for form 27Q
                        sql = "SELECT DISTINCT CoMst.CoID, CoMst.CoName, Deductee27Q.DId, DeductMst.DName, " _
                          & "DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, " _
                          & "StateMst.StateName ,deductmst.dpin , DeductMst.DState, DeductMst.DPan, DeductMst.DType, " _
                          & "DeductMst.DPANRef, RetnMst.FrmType,DeductMst.Category FROM (((Deductee27Q INNER JOIN DeductMst ON Deductee27Q.DId = DeductMst.DId) " _
                          & "INNER JOIN RetnMst ON Deductee27Q.RetnID = RetnMst.RetnID) INNER JOIN StateMst ON " _
                          & "DeductMst.DState = StateMst.StateID) INNER JOIN CoMst ON DeductMst.CoID = CoMst.CoID"
                End Select
                If cmbQuarter.SelectedIndex >= 0 Then
                    If cmbQuarter.Text = "24Q4" Then  'for Q4 of form 24Q(Annual salary)add by jayshree on 20/07/06
                        sql = sql & " where RetnMst.FrmType='24Q4'"
                    Else
                        sql = sql & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                    End If
                    If cmbCNm.SelectedIndex > -1 Then
                        sql = sql & " and CoMst.COId = " & cmbCNm.SelectedValue
                        If chkpan.Checked = vbTrue Then
                            sql = sql & " AND DeductMst.DPANcat in(1,2,3)"
                        End If
                    Else
                        If chkpan.Checked = vbTrue Then
                            sql = sql & " and DeductMst.DPANcat in(1,2,3)"
                        End If
                    End If

                ElseIf sqld <> vbNullString Then
                    If cmbtyp.Text = "24Q" Then   'for all qouter of form24Q with Annual Salary add by jayshree on 20/07/06
                        'sql = sql & " where RetnMst.FrmType in " & sqld & ""
                        sql = sql & " where RetnMst.FrmType in " & sqld
                    Else
                        ' sql = sql & " where RetnMst.FrmType in " & sqld & ""
                        sql = sql & " where RetnMst.FrmType in " & "(" & sqld & ")"
                    End If
                    If cmbCNm.SelectedIndex > -1 Then
                        sql = sql & " and CoMst.COId = " & cmbCNm.SelectedValue
                        If chkpan.Checked = vbTrue Then
                            sql = sql & " AND DeductMst.DPANcat In (1,2,3)"
                        End If
                    Else
                        If chkpan.Checked = vbTrue Then
                            sql = sql & " and DeductMst.DPANcat In (1,2,3)"
                        End If
                    End If
                End If
                sql = sql & " Order By DeductMst.Dname"


                ' add for same pan/ref of deductee 
                PrintRep = False
                sqlpan = " SELECT DISTINCT DeductMst.DPin, DeductMst.DName, DeductMst.DPan, DeductMst.DPANRef, DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, StateMst.StateName,  DeductMst.Category, RetnMst.FrmType, CoMst.CoName, CoMst.CoID " _
                    & " FROM (CoMst INNER JOIN (DeductMst INNER JOIN StateMst ON DeductMst.DState = StateMst.StateID) ON CoMst.CoID = DeductMst.CoID) INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID"
                If cmbQuarter.SelectedIndex >= 0 Then
                    sqlpan = sqlpan & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                ElseIf sqld <> vbNullString Then
                    sqlpan = sqlpan & " where RetnMst.FrmType in " & sqld & ""
                End If
                sqlpan = sqlpan & " and  DeductMst.CoID= " & cmbCNm.SelectedValue & " and DeductMst.DPan In (select DeductMst.DPANRef from deductmst)"

                sqlpan = sqlpan & " Union All"
                sqlpan = sqlpan & " SELECT DISTINCT DeductMst.DName, DeductMst.DPan, DeductMst.DPANRef, DeductMst.DAdd1, DeductMst.DAdd2, DeductMst.DAdd3, DeductMst.DAdd4, DeductMst.DAdd5, StateMst.StateName, DeductMst.DPin, DeductMst.Category, CoMst.CoName, CoMst.CoID, RetnMst.FrmType" _
                     & " FROM (CoMst INNER JOIN (DeductMst INNER JOIN StateMst ON DeductMst.DState = StateMst.StateID) ON CoMst.CoID = DeductMst.CoID) INNER JOIN RetnMst ON CoMst.CoID = RetnMst.CoID "
                If cmbQuarter.SelectedIndex >= 0 Then
                    sqlpan = sqlpan & " WHERE (((RetnMst.FrmType)='" & cmbtyp.Text & cmbQuarter.SelectedIndex + 1 & "'))"
                ElseIf sqld <> vbNullString Then
                    sqlpan = sqlpan & " where RetnMst.FrmType in " & sqld & ""
                End If
                sqlpan = sqlpan & " and  DeductMst.CoID= " & cmbCNm.SelectedValue & " and  DeductMst.DPANRef In (select DeductMst.DPan from deductmst)"
                rspan = FetchDataSet(sqlpan)
                If rspan.Tables(0).Rows.Count > 0 Then   ' for Anexure same pan with ref. no
                    ' frm.Show()
                    cmd = New OleDbCommand(sqlpan, cn)
                    da.SelectCommand = cmd
                    ds = New DataSet()
                    da.Fill(ds)
                    DataGridView1.DataSource = ds.Tables(0)
                    Excel2()
                    DataGridView1.Dispose()

                Else
                    cmd = New OleDbCommand(sql, cn)
                    da.SelectCommand = cmd
                    ds = New DataSet()
                    da.Fill(ds)
                    DataGridView1.DataSource = ds.Tables(0)
                    Excel1()
                    DataGridView1.Dispose()
                End If
            End If
        'End If
    End Sub

    Private Sub Excel1()
        Dim xlapp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")
        If chkpan.Checked = True Then
            xlSheet.Cells(2, 7) = "PAN Missing Deductee's Detail List Of " & cmbtyp.Text
        Else
            If cmbQuarter.Text = "" Then
                xlSheet.Cells(2, 7) = "Deductee's Detail List Of " & cmbtyp.Text
            Else
                xlSheet.Cells(2, 7) = "Deductee's Detail List Of " & cmbQuarter.Text
            End If
        End If
        If chknilrec.Checked = True Then
            xlSheet.Cells(2, 7) = "Deductee's Detail List Of No Entry in Transaction"
            xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 6)).Merge()
            'Report.Text1.SetText()
        End If

        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 3)).Value = "Deductee Name"
        xlSheet.Range(xlSheet.Cells(3, 4), xlSheet.Cells(3, 4)).Value = "Address1"
        xlSheet.Range(xlSheet.Cells(3, 5), xlSheet.Cells(3, 5)).Value = "Address2"
        xlSheet.Range(xlSheet.Cells(3, 6), xlSheet.Cells(3, 6)).Value = "Address3"
        xlSheet.Range(xlSheet.Cells(3, 7), xlSheet.Cells(3, 7)).Value = "Address4"
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).Value = "Address5"
        xlSheet.Range(xlSheet.Cells(3, 9), xlSheet.Cells(3, 9)).Value = "State Name"
        xlSheet.Range(xlSheet.Cells(3, 10), xlSheet.Cells(3, 10)).Value = "Pin Code"
        xlSheet.Range(xlSheet.Cells(3, 12), xlSheet.Cells(3, 12)).Value = "PAN No."
        xlSheet.Range(xlSheet.Cells(3, 13), xlSheet.Cells(3, 13)).Value = "Type"
        xlSheet.Range(xlSheet.Cells(3, 16), xlSheet.Cells(3, 16)).Value = "Category"
        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 16)).Font.Bold = True
        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 16)).BorderAround()

        Dim r As String
        r = (xlSheet.UsedRange.Rows.Count) + 1
        For m = 0 To DataGridView1.Rows.Count - 2
            For i = 0 To DataGridView1.Columns.Count - 2
                xlSheet.Cells(m + 4, i + 1) = DataGridView1.Rows(m).Cells(i + 1).Value.ToString()
                xlSheet.Range(xlSheet.Cells(r, i + 1), xlSheet.Cells(r, i + 1)).BorderAround(10)
                xlSheet.Range(xlSheet.Cells(m + 4, i + 1), xlSheet.Cells(m + 4, i + 1)).BorderAround(10)
            Next
        Next
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 9), xlSheet.Cells(3, 9)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 11), xlSheet.Cells(3, 11)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 11), xlSheet.Cells(3, 11)).EntireColumn.Delete()
        'xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 3)).EntireColumn.Delete()
        xlSheet.UsedRange.Cells.Columns.AutoFit()
        'For row = 4 To xlSheet.UsedRange.Rows.Count + 1

        '    ' xlSheet.Range(xlSheet.Cells(row, 4), xlSheet.Cells(row, 8)).Merge()
        '    Dim a As String = Chr(34) & "," & Chr(34)
        '    Dim i As String

        '    i = "=CONCATENATE(C" & row & "," & a & ",E" & row & ")"
        '    xlSheet.Cells(row, 4) = i
        '    ' "=sum(c4:c" & i - 1 & ")"
        '    '    'row = (xlSheet.UsedRange.Rows.Count)
        '    '    "=concatenate(d3:d)"
        '    '    Dim tmp As New List(Of String)
        '    '    tmp.AddRange((d3: d).ToString().Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries))
        '    '    Dim values() As String = xlSheet.Cells(5).ToString().Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        '    '    For i As Integer = 0 To values.Length - 1
        '    '        If i <= tmp.Count - 1 Then
        '    '            tmp(i) = tmp(i) & "," & values(i)
        '    '        Else
        '    '            tmp.Add("," & values(i))
        '    '        End If
        '    '    Next
        '    '    xlSheet.Cells(row, 6) = String.Join(vbCrLf, tmp.ToArray)
        'Next

        'xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 3)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(4, 11), xlSheet.Cells(4, 11)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(4, 14), xlSheet.Cells(4, 14)).EntireColumn.Hidden = True
        'xlSheet.Range(xlSheet.Cells(4, 15), xlSheet.Cells(4, 15)).EntireColumn.Hidden = True
        xlSheet.UsedRange.Cells.Columns.Font.Size = 8

        xlSheet.UsedRange.Cells.Columns.BorderAround(10)

        xlSheet.UsedRange.Cells.Columns.AutoFit()
        xlapp.Visible = True

    End Sub
    Private Sub Excel2()
        Dim xlapp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        xlapp = New Excel.Application
        xlBook = xlapp.Workbooks.Add
        xlSheet = xlBook.Worksheets("Sheet1")

        xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 1)).Value = "Deductee Name Having Same PAN with Reference No. of Other Deductee's Name"
        xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 11)).Merge()
        xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 11)).Font.Bold = True

        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 1)).Value = "Deductee Name"
        xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 2)).Value = "PAN No."
        xlSheet.Range(xlSheet.Cells(3, 3), xlSheet.Cells(3, 3)).Value = "PAN Ref. No."
        xlSheet.Range(xlSheet.Cells(3, 4), xlSheet.Cells(3, 4)).Value = "Address1"
        xlSheet.Range(xlSheet.Cells(3, 5), xlSheet.Cells(3, 5)).Value = "Address2"
        xlSheet.Range(xlSheet.Cells(3, 6), xlSheet.Cells(3, 6)).Value = "Address3"
        xlSheet.Range(xlSheet.Cells(3, 7), xlSheet.Cells(3, 7)).Value = "Address4"
        xlSheet.Range(xlSheet.Cells(3, 8), xlSheet.Cells(3, 8)).Value = "Address5"
        xlSheet.Range(xlSheet.Cells(3, 9), xlSheet.Cells(3, 9)).Value = "State Name"
        xlSheet.Range(xlSheet.Cells(3, 10), xlSheet.Cells(3, 10)).Value = "Category"
        xlSheet.Range(xlSheet.Cells(3, 11), xlSheet.Cells(3, 11)).Value = "Type"
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).Font.Bold = True
        xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).BorderAround()

        Dim r As String
        r = (xlSheet.UsedRange.Rows.Count) + 1
        For m = 0 To DataGridView1.Rows.Count - 2
            For i = 0 To DataGridView1.Columns.Count - 2
                xlSheet.Cells(m + 4, i + 1) = DataGridView1.Rows(m).Cells(i + 1).Value.ToString()
                xlSheet.Range(xlSheet.Cells(r, i + 1), xlSheet.Cells(r, i + 1)).BorderAround(10)
                xlSheet.Range(xlSheet.Cells(m + 4, i + 1), xlSheet.Cells(m + 4, i + 1)).BorderAround(10)
            Next
        Next
        xlSheet.Range(xlSheet.Cells(3, 12), xlSheet.Cells(3, 12)).EntireColumn.Delete()
        xlSheet.Range(xlSheet.Cells(3, 12), xlSheet.Cells(3, 12)).EntireColumn.Delete()
        xlSheet.UsedRange.Cells.Columns.AutoFit()
        xlSheet.UsedRange.Cells.Columns.Font.Size = 8
        xlSheet.UsedRange.Cells.Columns.BorderAround(10)
        xlapp.Visible = True

    End Sub

    Private Sub cmbdest_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbdest.SelectedIndexChanged

    End Sub

    Private Sub cmdc_Click(sender As Object, e As EventArgs) Handles cmdc.Click
        Dim rs As New DataSet
        Dim sql As String
        Dim frmd As New frmMulSelDed
        sql = "SELECT distinct RetnMst.FrmType FROM RetnMst" ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex)
        If cmbtyp.SelectedIndex = 0 Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"
        ElseIf cmbtyp.SelectedIndex = 1 Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cmbtyp.SelectedIndex = 2 Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        End If
        ' rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        cmbQuarter.SelectedIndex = -1
        frmd.sql1 = sql
        frmd.strflg = True
        frmd.ShowDialog()
        sqld = frmd.sql1
        If sqld = "" Then
            cmbQuarter.Enabled = True
        Else
            cmbQuarter.Enabled = False
        End If
        'frmd.Hide
    End Sub

    Private Sub Qtr()
        'Main()
        Dim rs As New DataSet
        'Dim rs1 As New ADODB.Recordset
        Dim sql As String
        Dim t1 As New DataTable()
        Dim da As New OleDbDataAdapter
        Dim cmd As New OleDbCommand
        Dim transaction As OleDbTransaction
        ' Dim sql1 As String
        cmbQuarter.Items.Clear()
        ' cmbQuarter.SelectedIndex = -1
        sql = "SELECT distinct RetnMst.FrmType FROM RetnMst" ' where COId = " & cmbCNm.ItemData(cmbCNm.ListIndex)
        cmbQuarter.Items.Clear()
        If cmbtyp.SelectedIndex = 0 Then
            sql = sql & " where RetnMst.frmtype in ('24Q1','24Q2','24Q3','24Q4')"
        ElseIf cmbtyp.SelectedIndex = 1 Then
            sql = sql & " where RetnMst.frmtype in ('26Q1','26Q2','26Q3','26Q4')"
        ElseIf cmbtyp.SelectedIndex = 2 Then
            sql = sql & " where RetnMst.frmtype in ('27EQ1','27EQ2','27EQ3','27EQ4')"
        ElseIf cmbtyp.SelectedIndex = 3 Then
            sql = sql & " where RetnMst.frmtype in ('27Q1','27Q2','27Q3','27Q4')"
        End If
        't1.Columns.Add("FrmType")
        't1.Rows.Count()
        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction
        cmd.CommandText = sql
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            cmbQuarter.Items.Add(dr.Item(0))
        End While
        dr.Close()
        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage
        End Try
        cmd.Dispose()
        transaction.Dispose()

        'da = New OleDbDataAdapter(sql, cn)
        'da.Fill(rs, "RetnMst")

        'For i = 0 To t1.Rows.Count - 1
        '    cmbQuarter.DataSource = t1
        '    cmbQuarter.DisplayMember = "FrmType"
        '    'cmbQuarter.Items.Add(t1.Rows(i).Item(0).ToString)
        'Next

        'If rs.Tables.Count > 0 Then
        'cmbQuarter.Items.Clear()
        'While Not rs.EOF
        '        cmbQuarter.Items.Add(rs!FrmType & "")
        '        cmbQuarter.Items(cmbQuarter.SelectedIndex) = rs!RetnId
        '        rs.MoveNext
        '    End While
        '    cmbQuarter.SelectedIndex = 0
        'End If
        'If rs.State = adStateOpen Then rs.Close 
        'Set rs = Nothing
    End Sub

    Private Sub cmbCNm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCNm.SelectedIndexChanged

    End Sub

    Private Sub cmbdest_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cmbdest.KeyPress

    End Sub

    Private Sub cmbQuarter_Click(sender As Object, e As EventArgs) Handles cmbQuarter.Click
        If cmbQuarter.Text = "24Q4" Then
            chksalary.Visible = True
        Else
            chksalary.Visible = False
        End If
    End Sub

    Private Sub chkpan_Click(sender As Object, e As EventArgs) Handles chkpan.Click
        If chkpan.Checked = True Then
            chknilrec.Enabled = False
        Else
            chknilrec.Enabled = True
        End If
    End Sub

    Private Sub cmbtyp_Click(sender As Object, e As EventArgs) Handles cmbtyp.Click
        Qtr()
    End Sub

    Private Sub chknilrec_Click(sender As Object, e As EventArgs) Handles chknilrec.Click
        If chknilrec.Checked = True Then
            chkpan.Enabled = False
            cmbtyp.Enabled = False
            cmbQuarter.Enabled = False
            cmdc.Enabled = False
        Else
            chkpan.Enabled = True
            cmbtyp.Enabled = True
            cmbQuarter.Enabled = True
            cmdc.Enabled = True
        End If
    End Sub

End Class
