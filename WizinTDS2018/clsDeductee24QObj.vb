Public Class clsDeductee24QObj
    'local variable(s) to hold property value(s)
    Private mvarID24Q As Long 'local copy
    Private mvarRetnID As Long 'local copy
    Private mvarDId As Long 'local copy
    Private mvarDCode As String 'local copy
    Private mvarAmtOfPayment As Double 'local copy
    Private mvarDtOfPayment As Date 'local copy
    'Private mvarIsBookEntry As Boolean 'local copy
    ' Private mvarRateOfTDS As Double 'local copy
    Private mvarTaxAmt As Double 'local copy
    Private mvarSurcharge As Double 'local copy
    Private mvarECess As Double 'local copy
    Private mvarTotalTaxDeposited As Double 'local copy
    Private mvarTotalTaxDeducted As Double 'local copy
    Private mvarDtOfDeduction As Date 'local copy
    Private mvarRemark As String 'local copy
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent AfterDelete[(arg1, arg2, ... , argn)]
    Public Event AfterDelete()
    Public Event AfterSave()
    Public Event BeforeDelete(ByVal Cancel As Boolean)
    Public Event BeforeSave(ByVal Cancel As Boolean)
    Public Event Fetch(ByVal ID As Long)
    Public Event PrepareDataForSave(ByVal Cancel As Boolean)
    'local variable(s) to hold property value(s)
    Private mvarChallanId As Long 'local copy
    'local variable(s) to hold property value(s)
    Private mvarSec As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarCertNo As String 'local copy

    Public Property CertNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CertNo
            CertNo = mvarCertNo
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.CertNo = 5
            mvarCertNo = vData
        End Set
    End Property

    Public Property Sec As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Sec
            Sec = mvarSec
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Sec = 5
            mvarSec = vData
        End Set
    End Property
    Public Property ChallanID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ChallanNo
            ChallanID = mvarChallanId
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.ChallanNo = 5
            mvarChallanId = vData
        End Set
    End Property




    Public Function Update(ByVal Deductee As clsDeductee24QObj) As Boolean
        Dim sql As String, cnl As Boolean
        'On Error GoTo UpErr
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then
            Update = False
            Exit Function

        End If
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then
            Update = False
            Exit Function
        End If
        With Deductee
            sql = "Update Deductee24Q Set RetnId = " & IIf(.RetnID = 0, 0, .RetnID) & "," _
                & " DId = " & IIf(.did = 0, 0, .did) & "," _
                & " DCode = " & IIf(.DCode = vbNullString, "Null", "'" & .DCode & "'") & "," _
                & " Sec = " & IIf(.Sec = vbNullString, "Null", "'" & .Sec & "'") & "," _
                & " AmtOfPayment = " & IIf(.AmtOfPayment = 0, 0, .AmtOfPayment) & "," _
                & " DtOfPayment = #" & Format(.DtOfPayment, "dd/MMM/yyyy") & "#," _
                & " TaxAmt = " & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
                & " Surcharge = " & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
                & " ECess = " & IIf(.ECess = 0, 0, .ECess) & "," _
                & " TotalTaxDeducted  = " & IIf(.TotalTaxDeducted = 0, 0, .TotalTaxDeducted) & "," _
                & " DtOfDeduction  =" & IIf(.TaxAmt = 0, "Null", "#" & Format(.DtOfDeduction, "dd/MMM/yyyy") & "#") & "," _
                & " TotalTaxDeposited  = " & IIf(.TotalTaxDeposited = 0, 0, .TotalTaxDeposited) & "," _
                & " ChallanID  = " & IIf(.ChallanID = 0, 0, .ChallanID) & "," _
                & " Remark = " & IIf(.Remark = vbNullString, "Null", "'" & .Remark & "'") & "," _
                & " CertNo = " & "'" & .CertNo & "" & "'" _
                & " Where Id24q = " & .ID24Q
        End With
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


    End Function

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(Id24Q) as Id From Deductee24Q"
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            If nds.Tables(0).Rows(0)(0).ToString() = "" Then
                MaxID = 0
            Else
                MaxID = nds.Tables(0).Rows(0)(0)
            End If

        Else
            MaxID = 0
        End If
        nds.Dispose()
    End Function

    Public Function Insert(ByVal Deductee As clsDeductee24QObj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction

        Dim Err As ErrObject
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
        With Deductee

            .ID24Q = MaxID() + 1
            sql = "Insert Into Deductee24Q (Id24Q,RetnId,DId,DCode,Sec,AmtOfPayment,DtOfPayment,TaxAmt,Surcharge,ECess,TotalTaxDeposited,TotalTaxDeducted,DtOfDeduction,Remark,ChallanID,CertNo) Values(" _
                & .ID24Q & ", " & IIf(.RetnID = 0, 0, .RetnID) & ", " _
                & IIf(.did = 0, 0, .did) & ", " _
                & IIf(.DCode = vbNullString, "Null", "'" & .DCode & "'") & "," _
                & IIf(.Sec = vbNullString, "Null", "'" & .Sec & "'") & "," _
                & IIf(.AmtOfPayment = 0, 0, .AmtOfPayment) & ",#" _
                & Format(.DtOfPayment, "dd/MMM/yyyy") & "#," _
            & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
            & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
            & IIf(.ECess = 0, 0, .ECess) & "," _
            & IIf(.TotalTaxDeposited = 0, 0, .TotalTaxDeposited) & "," _
            & IIf(.TotalTaxDeducted = 0, 0, .TotalTaxDeducted) & "," _
            & IIf(.TaxAmt = 0, "Null", "#" & Format(.DtOfDeduction, "dd/MMM/yyyy") & "#") & "," _
            & IIf(.Remark = vbNullString, "Null", "'" & .Remark & "'") & "," _
            & IIf(.ChallanID = 0, 0, .ChallanID) & "," _
            & IIf(.CertNo = vbNullString, "Null", "'" & .CertNo & "'") & ")"
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
Err:
        MsgBox(Err.Description, , Err.Number)
    End Function

    Public Function Delete(ByVal ID As Long) As Boolean
        Dim cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand

        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then
            Delete = False
            Exit Function
        End If
        Dim sql As String

        sql = "Delete * From Deductee24Q Where Id24Q = " & ID
        cmd.Connection = cn
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction

        Try
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
            Delete = True
            transaction.Commit()

        Catch ex As Exception
            MessageBox.Show("Message:", ex.Message)

            transaction.Rollback()

            Delete = False
        End Try

        cmd.Dispose()
        transaction.Dispose()
    End Function


    Public Property Remark As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Remark
            Remark = mvarRemark
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Remark = 5
            mvarRemark = vData
        End Set
    End Property

    Public Property DtOfDeduction As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfDeduction
            DtOfDeduction = mvarDtOfDeduction
        End Get
        Set(ByVal vData As Date)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DtOfDeduction = 5
            mvarDtOfDeduction = vData
        End Set
    End Property

    Public Property TotalTaxDeducted As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TotalTaxDeducted
            TotalTaxDeducted = mvarTotalTaxDeducted
        End Get
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TotalTaxDeducted = 5
            mvarTotalTaxDeducted = vData
        End Set
    End Property

    Public Property TotalTaxDeposited As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TotalTaxDeposited
            TotalTaxDeposited = mvarTotalTaxDeposited
        End Get
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TotalTaxDeposited = 5
            mvarTotalTaxDeposited = vData
        End Set
    End Property

    Public Property ECess As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ECess
            ECess = mvarECess
        End Get
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.ECess = 5
            mvarECess = vData
        End Set
    End Property

    Public Property Surcharge As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Surcharge
            Surcharge = mvarSurcharge
        End Get
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Surcharge = 5
            mvarSurcharge = vData
        End Set
    End Property
    Public Property TaxAmt As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TaxAmt
            TaxAmt = mvarTaxAmt
        End Get
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TaxAmt = 5
            mvarTaxAmt = vData
        End Set
    End Property

    'Public Property RateOfTDS As Double
    '    Get
    '        'used when retrieving value of a property, on the right side of an assignment.
    '        'Syntax: Debug.Print X.RateOfTDS
    '        RateOfTDS = mvarRateOfTDS
    '    End Get
    '    Set(ByVal vData As Double)
    '        'used when assigning a value to the property, on the left side of an assignment.
    '        'Syntax: X.RateOfTDS = 5
    '        mvarRateOfTDS = vData
    '    End Set
    'End Property

    'Public Property IsBookEntry As Boolean
    '    Get
    '        'used when retrieving value of a property, on the right side of an assignment.
    '        'Syntax: Debug.Print X.IsBookEntry
    '        IsBookEntry = mvarIsBookEntry
    '    End Get
    '    Set(ByVal vData As Boolean)
    '        'used when assigning a value to the property, on the left side of an assignment.
    '        'Syntax: X.IsBookEntry = 5
    '        mvarIsBookEntry = vData
    '    End Set
    'End Property

    Public Property DtOfPayment As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfPayment
            DtOfPayment = mvarDtOfPayment
        End Get
        Set(ByVal vData As Date)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DtOfPayment = 5
            mvarDtOfPayment = vData
        End Set
    End Property

    Public Property AmtOfPayment As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AmtOfPayment
            AmtOfPayment = mvarAmtOfPayment
        End Get
        Set(ByVal vData As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.AmtOfPayment = 5
            mvarAmtOfPayment = vData
        End Set
    End Property

    Public Property DCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DCode
            DCode = mvarDCode
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DCode = 5
            mvarDCode = vData
        End Set
    End Property

    Public Property did As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            did = mvarDId
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DId = 5
            mvarDId = vData
        End Set
    End Property

    Public Property RetnID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RetnID
            RetnID = mvarRetnID
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.RetnID = 5
            mvarRetnID = vData
        End Set
    End Property

    Public Property ID24Q As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ID24Q
            ID24Q = mvarID24Q
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.ID24Q = 5
            mvarID24Q = vData
        End Set
    End Property


    Public Function getChallanNo(ByVal ID As Long) As String
        Dim nds As New DataSet, dt As Date
        nds = FetchDataSet("select BankChallanNo,DtOfChallan, TranVouNo from Challan24Q WHERE challanID=" & ID)
        If nds.Tables(0).Rows.Count > 0 Then
            'If String.IsNullOrEmpty(rs(0).Value) Then
            dt = nds.Tables(0).Rows(0)(1).ToString()

            'dt.ToString("dd/MM/yy")
            If nds.Tables(0).Rows(0)(0).ToString() = "" Then
                getChallanNo = SetFormat("00000", nds.Tables(0).Rows(0)(2).ToString()) & " - " & Format(dt, "dd/MM/yy")

            Else

                getChallanNo = SetFormat("00000", nds.Tables(0).Rows(0)(0).ToString()) & " - " & Format(dt, "dd/MM/yy")
            End If
        Else
            getChallanNo = vbNullString
        End If
        nds.Dispose()
    End Function
    Public Function GetChallanID(ByVal BChlNo As Long, ChlDt As Date) As Long

        Dim nds As New DataSet

        nds = FetchDataSet("SELECT challanid FROM Challan24Q WHERE BankChallanNo=" & BChlNo & " and DtOfChallan=#" & Format(ChlDt, "MM/dd/yyyy") & "#")
        If nds.Tables(0).Rows.Count > 0 Then
            GetChallanID = nds.Tables(0).Rows(0)(0)
        Else
            GetChallanID = 0
        End If
        nds.Dispose()

    End Function
End Class
