Public Class ClsChallan27Qobj

    'local variable(s) to hold property value(s)
    Private mvarId As Long 'local copy
    Private mvarRetnID As Long 'local copy
    Private mvarSection As String 'local copy
    Private mvarAmt As Double 'local copy
    Private mvarDtOfVoucher As Date 'local copy
    Private mvarBankBrCode As Long 'local copy
    Public Event PrepareDataForSave(ByVal Cancel As Boolean)
    Public Event BeforeSave(ByVal Cancel As Boolean)
    Public Event AfterSave()
    Public Event BeforeDelete(ByVal Cancel As Boolean)
    Public Event AfterDelete()
    Private mvarSurcharge As Double 'local copy
    Private mvarECess As Double 'local copy
    Private mvarInterest As Double 'local copy
    Private mvarOthers As Double 'local copy
    Private mvarChqNo As Double 'local copy
    Private mvarIsBookEntry As Boolean 'local copy
    Private mvarCollCode As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarTotalTax As Double 'local copy
    Private mvarAInterest As Double 'local copy
    Private mvarAECess As Double 'local copy
    Private mvarTranVouNo As Long 'local copy
    Private mvarRemark As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarBankChallanNo As Long 'local copy
    Private mvarAFees As Double 'local copy
    Private mvarMinorHead As String 'local copy

    Public Property MinorHead As String
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.MinorHead = 5
            mvarMinorHead = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            MinorHead = mvarMinorHead
        End Get
    End Property

    Public Property AFees As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.AFees = 5
            mvarAFees = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AFees
            AFees = mvarAFees
        End Get
    End Property

    Public Property BankChallanNo() As Long
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.BankChallanNo = 5
            mvarBankChallanNo = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BankChallanNo
            BankChallanNo = mvarBankChallanNo
        End Get
    End Property

    Public Function LinkDed27Q(ByVal ID As Long) As Boolean
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select ChallanID  From Deductee27Q Where ChallanID = " & ID
        nds = FetchDataSet(sql)

        If nds.Tables(0).Rows.Count > 0 Then
            LinkDed27Q = True
        Else
            LinkDed27Q = False
        End If
        nds.Dispose()
    End Function

    Public Property Remark As String
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Remark = 5
            mvarRemark = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Remark
            Remark = mvarRemark
        End Get
    End Property

    Public Property TranVouNo As Long
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TranVouNo = 5
            mvarTranVouNo = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TranVouNo
            TranVouNo = mvarTranVouNo
        End Get
    End Property

    Public Property AOthers As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.AECess = 5
            mvarAECess = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AECess
            AOthers = mvarAECess
        End Get
    End Property

    Public Property AInterest As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.AInterest = 5
            mvarAInterest = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AInterest
            AInterest = mvarAInterest
        End Get
    End Property

    Public Property TotalTax As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TotalTax = 5
            mvarTotalTax = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TotalTax
            TotalTax = mvarTotalTax
        End Get
    End Property

    Public Property CollCode As String
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.CollCode = 5
            mvarCollCode = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CollCode
            CollCode = mvarCollCode
        End Get
    End Property

    Public Property IsBookEntry As Boolean
        Set(ByVal vData As Boolean)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.IsBookEntry = 5
            mvarIsBookEntry = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.IsBookEntry
            IsBookEntry = mvarIsBookEntry
        End Get
    End Property

    Public Property ChqDDNo As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.ChqNo = 5
            mvarChqNo = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ChqNo
            ChqDDNo = mvarChqNo
        End Get
    End Property

    Public Property Others As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Others = 5
            mvarOthers = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Others
            Others = mvarOthers
        End Get
    End Property

    Public Property Interest As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Interest = 5
            mvarInterest = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Interest
            Interest = mvarInterest
        End Get
    End Property

    Public Property ECess As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.ECess = 5
            mvarECess = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ECess
            ECess = mvarECess
        End Get
    End Property

    Public Property Surcharge As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Surcharge = 5
            mvarSurcharge = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Surcharge
            Surcharge = mvarSurcharge
        End Get
    End Property

    Public Property BankBrCode As Long
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.BankBrCode = 5
            mvarBankBrCode = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BankBrCode
            BankBrCode = mvarBankBrCode
        End Get
    End Property

    Public Property DtOfChallan As Date
        Set(ByVal vData As Date)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DtOfVoucher = 5
            mvarDtOfVoucher = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfVoucher
            DtOfChallan = mvarDtOfVoucher
        End Get
    End Property

    Public Property TaxAmt As Double
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Amt = 5
            mvarAmt = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Amt
            TaxAmt = mvarAmt
        End Get
    End Property

    Public Property Sec As String
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Section = 5
            mvarSection = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Section
            Sec = mvarSection
        End Get
    End Property

    Public Property RetnID As Long
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.RetnId = 5
            mvarRetnID = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RetnId
            RetnID = mvarRetnID
        End Get
    End Property

    Public Property ChallanID As Long
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Id = 5
            mvarId = vData
        End Set
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Id
            ChallanID = mvarId
        End Get
    End Property

    Public Function Fetch(ByVal ID As Long) As ClsChallan27Qobj
        Dim nds As New DataSet, Chln As New ClsChallan27Qobj
        Dim dt As Date
        nds = FetchDataSet("SELECT * FROM Challan27Q WHERE ChallanID=" & ID)
        If nds.Tables(0).Rows.Count > 0 Then
            With Chln
                .ChallanID = nds.Tables(0).Rows(0)("ChallanID").Value & ""
                .RetnID = nds.Tables(0).Rows(0)("RetnID").Value & ""
                .Sec = nds.Tables(0).Rows(0)("Sec").Value & ""
                .TaxAmt = nds.Tables(0).Rows(0)("TaxAmt").Value & ""
                .BankChallanNo = nds.Tables(0).Rows(0)("BankChallanNo").Value & ""
                dt = nds.Tables(0).Rows(0)("DtOfChallan")
                .DtOfChallan = dt.ToString("dd/MMM/yyyy") 'Format(rst("DtOfChallan").Value, "dd/MMM/yyyy") & ""
                .BankBrCode = nds.Tables(0).Rows(0)("BankBrCode").Value & ""
                .AFees = nds.Tables(0).Rows(0)("AFees").Value & ""
                .MinorHead = nds.Tables(0).Rows(0)("MinorHead").Value & ""
            End With
            Fetch = Chln
        Else
            Fetch = Nothing
        End If
        nds.Dispose()
        Chln = Nothing
    End Function

    Public Function Delete(ByVal ID As Long) As Boolean
        Dim cnl As Boolean
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction
        'On Error GoTo DelErr
        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then
            Delete = False
            Exit Function
        End If
        Dim sql As String
        sql = "Delete * From Challan27Q Where ChallanID = " & ID
        cmd.Connection = cn
        cmd.CommandText = sql
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction

        Try

            cmd.ExecuteNonQuery()
            transaction.Commit()
            RaiseEvent AfterDelete()
            Delete = True
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message)
            Delete = False
        End Try

        cmd.Dispose()
        transaction.Dispose()
        Exit Function


DelErr:
        MsgBox(err.Description, , err.Number)
    End Function

    Public Function Update(ByVal Challan As ClsChallan27Qobj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
        'On Error GoTo UpErr
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then
            Update = False
            Exit Function
        End If
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then
            Update = False

            Exit Function
        End If
        With Challan
            sql = "Update Challan27Q Set RetnId = " & IIf(.RetnID = 0, 0, .RetnID) & "," _
                & " Sec = " & IIf(.Sec = vbNullString, "Null", "'" & .Sec & "'") & "," _
                & " taxAmt = " & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
                & " Surcharge = " & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
                & " ecess = " & IIf(.ECess = 0, 0, .ECess) & "," _
                & " interest = " & IIf(.Interest = 0, 0, .Interest) & "," _
                & " others = " & IIf(.Others = 0, 0, .Others) & "," _
                & " TotalTax = " & IIf(.TotalTax = 0, 0, .TotalTax) & "," _
                & " AInterest = " & IIf(.AInterest = 0, 0, .AInterest) & "," _
                & " AOthers = " & IIf(.AOthers = 0, 0, .AOthers) & "," _
                & " chqddno = " & IIf(.ChqDDNo = -1, "Null", .ChqDDNo) & "," _
                & " isbookentry = " & .IsBookEntry & "," _
                & " BankChallanNo = " & IIf(.BankChallanNo = -1, "Null", .BankChallanNo) & "," _
                & " TranVouNo = " & IIf(.TranVouNo = -1, "Null", .TranVouNo) & "," _
                & " DtOfChallan = #" & Format(.DtOfChallan, "dd/MMM/yyyy") & "#," _
                & " BankBrCode  = " & IIf(.BankBrCode = -1, "Null", .BankBrCode) & "," _
                & " Remark = " & IIf(.Remark = vbNullString, "Null", "'" & .Remark & "'") & "," _
                & " AFees = " & IIf(.AFees = 0, 0, .AFees) & "," _
                & " MinorHead = " & IIf(.MinorHead = vbNullString, "Null", "'" & .MinorHead & "'") _
                & " Where ChallanId = " & .ChallanID
            cmd.Connection = cn
            cmd.CommandText = sql
            transaction = cn.BeginTransaction()
            cmd.Transaction = transaction

            Try
                cmd.ExecuteNonQuery()
                transaction.Commit()
                Update = True
                RaiseEvent AfterSave()

            Catch ex As Exception
                transaction.Rollback()
                MessageBox.Show(ex.Message) 'Error MEssage
                Update = False
                Exit Function
            End Try


            cmd.Dispose()
            transaction.Dispose()


        End With


    End Function
    'Private Sub oChln1_PrepareDataForSave(Cancel As Boolean)
    '   With oChln1
    '      .ChallanID = IIf(Val(cboChallanSection.Tag) = 0, 0, Val(cboChallanSection.Tag))
    '      .RetnId = Me.Tag
    '      .Sec = cboChallanSection.Text
    '      .TaxAmt = IIf(Len(Trim(txtAmtDeducted.Text)) = 0, 0, txtAmtDeducted)
    '      .Surcharge = IIf(Len(Trim(txtSurcharge.Text)) = 0, 0, txtSurcharge.Text)
    '      .ECess = IIf(Len(Trim(txtECess.Text)) = 0, 0, txtECess.Text)
    '      .Interest = IIf(Len(Trim(txtIntt.Text)) = 0, 0, txtIntt.Text)
    '      .Others = IIf(Len(Trim(txtOthers.Text)) = 0, 0, txtOthers.Text)
    ''      .AInterest = vbNullString
    ''      .AOthers = vbNullString
    '      .TotalTax = IIf(Len(Trim(txtTotalTDS.Text)) = 0, 0, txtTotalTDS.Text)
    '      .BankChallanNo = IIf(IsNull(txtChallanNo) Or txtChallanNo = "", -1, txtChallanNo.Text)
    '     '' .BankChallanNo = IIf(Len(Trim(txtChallanNo)) = 0, "Null", txtChallanNo.Text)
    '      .DtOfChallan = Format(dtpChallanDate.Text, "dd/MMM/yyyy")
    '      .BankBrCode = IIf(Len(Trim(cboBankBrCode.Text)) = 0, -1, cboBankBrCode.Text)
    '      .TranVouNo = IIf(Len(Trim(txtTranVouNo.Text)) = 0, -1, txtTranVouNo.Text)
    '      .IsBookEntry = IIf(chkBookEntry = vbChecked, True, False)
    '      .ChqDDNo = IIf(Len(Trim(txtChqNo.Text)) = 0, -1, txtChqNo.Text)
    '      .Remark = txtChallanRemark.Text
    '   End With
    'End Sub


    Public Function Insert(ByVal Challan As ClsChallan27Qobj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction

        'On Error GoTo InErr
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then
            Insert = False
            Exit Function
        End If
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then
            Insert = False
            Exit Function
        End If

        With Challan
            .ChallanID = MaxID + 1

            sql = " Insert Into Challan27Q (ChallanId,RetnId,Sec,TaxAmt, Surcharge, Ecess, Interest, Others, "
            sql = sql & "TotalTax,AInterest,AOthers,ChqDDNo, BankChallanNo,TranVouNo, DtOfChallan, BankBrCode, IsBookEntry,Remark,"
            sql = sql & "Afees,MinorHead) Values("
            sql = sql & .ChallanID & "," & IIf(.RetnID = 0, 0, .RetnID) & ","
            sql = sql & IIf(.Sec = vbNullString, "Null", "'" & .Sec & "'") & ","
            sql = sql & IIf(.TaxAmt = 0, 0, .TaxAmt) & ","
            sql = sql & IIf(.Surcharge = 0, 0, .Surcharge) & ","
            sql = sql & IIf(.ECess = 0, 0, .ECess) & ","
            sql = sql & IIf(.Interest = 0, 0, .Interest) & ","
            sql = sql & IIf(.Others = 0, 0, .Others) & ","
            sql = sql & IIf(.TotalTax = 0, 0, .TotalTax) & ","
            sql = sql & IIf(.AInterest = 0, 0, .AInterest) & ","
            sql = sql & IIf(.AOthers = 0, 0, .AOthers) & ","
            sql = sql & IIf(.ChqDDNo = -1, "Null", .ChqDDNo) & ","
            sql = sql & IIf(.BankChallanNo = -1, "Null", .BankChallanNo) & ","
            sql = sql & IIf(.TranVouNo = -1, "Null", .TranVouNo) & ","
            sql = sql & "#" & Format(.DtOfChallan, "dd/MMM/yyyy") & "#,"
            sql = sql & IIf(.BankBrCode = -1, "Null", .BankBrCode) & ","
            sql = sql & .IsBookEntry & ","
            sql = sql & IIf(.Remark = vbNullString, "Null", "'" & .Remark & "'") & ","
            sql = sql & IIf(.AFees = 0, 0, .AFees) & ","
            sql = sql & IIf(.MinorHead = vbNullString, "Null", "'" & .MinorHead & "'") & ")"
        End With

        cmd.CommandText = sql
        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction

        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
            RaiseEvent AfterSave()
            Insert = True
        Catch ex As Exception

            transaction.Rollback()
            MessageBox.Show(ex.Message)
            Insert = False
        End Try

        cmd.Dispose()
        transaction.Dispose()
        Exit Function


    End Function

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(ChallanId) as ID From Challan27Q"
        nds = FetchDataSet(sql)


        If nds.Tables(0).Rows.Count > 0 Then
            If nds.Tables(0).Rows(0)(0).ToString = "" Then
                MaxID = 0
            Else
                MaxID = nds.Tables(0).Rows(0)(0)
            End If
        Else
            MaxID = 0
        End If

        nds.Dispose()
    End Function


End Class
