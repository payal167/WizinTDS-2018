Public Class clsDeductee27QObj
    'local variable(s) to hold property value(s)
    Private mvarID27Q As Long 'local copy
    Private mvarRetnID As Long 'local copy
    Private mvarDId As Long 'local copy
    Private mvarDCode As String 'local copy
    Private mvarAmtOfPayment As Double 'local copy
    Private mvarDtOfPayment As Date 'local copy
    Private mvarIsBookEntry As Boolean 'local copy
    Private mvarRateOfTDS As Double 'local copy
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
    Private mvarDTAA As String 'local copy
    Private mvarRemitID As String 'local copy
    Private mvarUniqueAck As String 'local copy
    Private mvarCountryID As String 'local copy

    'Public Property Let CountryID(ByVal vData As String)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.CountryID = 5
    '    mvarCountryID = vData
    'End Property


    Public Property CountryID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CountryID
            CountryID = mvarCountryID
        End Get
        Set(ByVal vData As String)
            mvarCountryID = vData
        End Set
    End Property



    'Public Property Let UniqueAck(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.UniqueAck = 5
    '    mvarUniqueAck = vData
    'End Property


    Public Property UniqueAck As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.UniqueAck
            UniqueAck = mvarUniqueAck
        End Get
        Set(ByVal vData As String)
            mvarUniqueAck = vData
        End Set
    End Property



    'Public Property Let RemitID(ByVal vData As String)  'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.RemitID = 5
    '    mvarRemitID = vData
    'End Property


    Public Property RemitID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RemitID
            RemitID = mvarRemitID
        End Get
        Set(ByVal vData As String)
            mvarRemitID = vData
        End Set
    End Property



    'Public Property Let DTAA(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DTAA = 5
    '    mvarDTAA = vData
    'End Property


    Public Property DTAA As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DTAA
            DTAA = mvarDTAA
        End Get
        Set(ByVal vData As String)
            mvarDTAA = vData
        End Set
    End Property



    'Public Property Let CertNo(ByVal vData As String)    'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.CertNo = 5
    '    mvarCertNo = vData
    'End Property


    Public Property CertNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CertNo
            CertNo = mvarCertNo
        End Get
        Set(ByVal vData As String)
            mvarCertNo = vData
        End Set
    End Property



    'Public Property Let Sec(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Sec = 5
    '    mvarSec = vData
    'End Property


    Public Property Sec As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Sec
            Sec = mvarSec
        End Get
        Set(ByVal vData As String)
            mvarSec = vData
        End Set
    End Property



    'Public Property Let ChallanID(ByVal vData As Long)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.ChallanNo = 5
    '    mvarChallanId = vData
    'End Property


    Public Property ChallanID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ChallanNo
            ChallanID = mvarChallanId
        End Get
        Set(ByVal vData As Long)
            mvarChallanId = vData
        End Set
    End Property




    Public Function Update(ByVal Deductee As clsDeductee27QObj) As Boolean
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

            sql = "Update Deductee27Q Set RetnId = " & IIf(.RetnID = 0, 0, .RetnID) & "," _
                & " DId = " & IIf(.did = 0, 0, .did) & "," _
                & " DCode = " & IIf(.DCode = vbNullString, "Null", "'" & .DCode & "'") & "," _
                & " Sec = " & IIf(.Sec = vbNullString, "Null", "'" & .Sec & "'") & "," _
                & " AmtOfPayment = " & IIf(.AmtOfPayment = 0, 0, .AmtOfPayment) & "," _
                & " DtOfPayment = #" & Format(.DtOfPayment, "dd/MMM/yyyy") & "#," _
                & " RateOfTDS = " & IIf(.RateOfTDS = 0, 0, .RateOfTDS) & "," _
                & " TaxAmt = " & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
                & " Surcharge = " & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
                & " ECess = " & IIf(.ECess = 0, 0, .ECess) & "," _
                & " TotalTaxDeducted  = " & IIf(.TotalTaxDeducted = 0, 0, .TotalTaxDeducted) & "," _
                & " DtOfDeduction  =" & IIf(.TaxAmt = 0, "Null", "#" & Format(.DtOfDeduction, "dd/MMM/yyyy") & "#") & "," _
                & " TotalTaxDeposited  = " & IIf(.TotalTaxDeposited = 0, 0, .TotalTaxDeposited) & "," _
                & " ChallanID  = " & IIf(.ChallanID = 0, 0, .ChallanID) & "," _
                & " Remark = " & IIf(.Remark = vbNullString, "Null", "'" & .Remark & "'") & "," _
                & " IsBookEntry = " & .IsBookEntry & "," _
                & " CertNo = " & IIf(.CertNo = vbNullString, "Null", "'" & .CertNo & "'") & "," _
                & " DTAA = " & IIf(.DTAA = vbNullString, "Null", "'" & .DTAA & "'") & "," _
                & " Remitid = " & IIf(.RemitID = vbNullString, "Null", "'" & .RemitID & "'") & "," _
                & " UniqueAck = " & IIf(.UniqueAck = vbNullString, "Null", "'" & .UniqueAck & "'") & "," _
                & " CountryID = " & IIf(.CountryID = vbNullString, "Null", "'" & .CountryID & "'") & " " _
                & " Where ID27Q = " & .ID27Q
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
        sql = "Select Max(ID27Q) as Id From Deductee27Q"
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

    Public Function Insert(ByVal Deductee As clsDeductee27QObj) As Boolean
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
            .ID27Q = MaxID() + 1


            sql = "Insert Into Deductee27Q (ID27Q,RetnId,DId,DCode,Sec,AmtOfPayment,DtOfPayment,RateOfTDS, " _
                & "TaxAmt,Surcharge,ECess,TotalTaxDeposited,TotalTaxDeducted,DtOfDeduction,Remark,ChallanID,IsBookEntry,CertNo," _
                & "DTAA,RemitID,UniqueAck,CountryID) Values (" _
                & .ID27Q & "," & IIf(.RetnID = 0, 0, .RetnID) & "," _
                & IIf(.did = 0, 0, .did) & "," _
                & IIf(.DCode = vbNullString, "Null", "'" & .DCode & "'") & "," _
                & IIf(.Sec = vbNullString, "Null", "'" & .Sec & "'") & "," _
                & IIf(.AmtOfPayment = 0, 0, .AmtOfPayment) & ",#" _
                & Format(.DtOfPayment, "dd/MMM/yyyy") & "#," _
                & IIf(.RateOfTDS = 0, 0, .RateOfTDS) & "," _
                & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
                & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
                & IIf(.ECess = 0, 0, .ECess) & "," _
                & IIf(.TotalTaxDeposited = 0, 0, .TotalTaxDeposited) & "," _
                & IIf(.TotalTaxDeducted = 0, 0, .TotalTaxDeducted) & "," _
                & IIf(.TaxAmt = 0, "Null", "#" & Format(.DtOfDeduction, "dd/MMM/yyyy") & "#") & "," _
                & IIf(.Remark = vbNullString, "Null", "'" & .Remark & "'") & "," _
                & IIf(.ChallanID = 0, 0, .ChallanID) & ", " _
                & .IsBookEntry & "," _
                & IIf(.CertNo = vbNullString, "Null", "'" & .CertNo & "'") & "," _
                & IIf(.DTAA = vbNullString, "Null", "'" & .DTAA & "'") & "," _
                & IIf(.RemitID = vbNullString, "Null", "'" & .RemitID & "'") & "," _
                & IIf(.UniqueAck = vbNullString, "Null", "'" & .UniqueAck & "'") & "," _
                & IIf(.CountryID = vbNullString, "Null", "'" & .CountryID & "'") & ")"


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
        MsgBox(err.Description, , err.Number)
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

        sql = "Delete * From Deductee27Q Where ID27Q = " & ID
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

    'Public Property Let Remark(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Remark = 5
    '    mvarRemark = vData
    'End Property


    Public Property Remark As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Remark
            Remark = mvarRemark
        End Get
        Set(ByVal vData As String)
            mvarRemark = vData
        End Set

    End Property



            'Public Property Let DtOfDeduction(ByVal vData As Date)
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.DtOfDeduction = 5
            '    mvarDtOfDeduction = vData
            'End Property


    Public Property DtOfDeduction As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfDeduction
            DtOfDeduction = mvarDtOfDeduction
        End Get
        Set(ByVal vData As Date)
            mvarDtOfDeduction = vData
        End Set
    End Property



            'Public Property Let TotalTaxDeducted(ByVal vData As Double)    'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.TotalTaxDeducted = 5
            '    mvarTotalTaxDeducted = vData
            'End Property


    Public Property TotalTaxDeducted As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TotalTaxDeducted
            TotalTaxDeducted = mvarTotalTaxDeducted
        End Get
        Set(ByVal vData As Double)
            mvarTotalTaxDeducted = vData
        End Set
    End Property



            'Public Property Let TotalTaxDeposited(ByVal vData As Double)   'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.TotalTaxDeposited = 5
            '    mvarTotalTaxDeposited = vData
            'End Property


    Public Property TotalTaxDeposited As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TotalTaxDeposited
            TotalTaxDeposited = mvarTotalTaxDeposited
        End Get
        Set(ByVal vData As Double)
            mvarTotalTaxDeposited = vData
        End Set
    End Property



            'Public Property Let ECess(ByVal vData As Double)  'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.ECess = 5
            '    mvarECess = vData
            'End Property


    Public Property ECess As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ECess
            ECess = mvarECess
        End Get
        Set(ByVal vData As Double)
            mvarECess = vData
        End Set
    End Property



            'Public Property Let Surcharge(ByVal vData As Double)   'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.Surcharge = 5
            '    mvarSurcharge = vData
            'End Property


    Public Property Surcharge As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Surcharge
            Surcharge = mvarSurcharge
        End Get
        Set(ByVal vData As Double)
            mvarSurcharge = vData
        End Set
    End Property



            'Public Property Let TaxAmt(ByVal vData As Double)          'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.TaxAmt = 5
            '    mvarTaxAmt = vData
            'End Property


    Public Property TaxAmt As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TaxAmt
            TaxAmt = mvarTaxAmt
        End Get
        Set(ByVal vData As Double)
            mvarTaxAmt = vData
        End Set
    End Property



            'Public Property Let RateOfTDS(ByVal vData As Double)    'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.RateOfTDS = 5
            '    mvarRateOfTDS = vData
            'End Property


    Public Property RateOfTDS As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RateOfTDS
            RateOfTDS = mvarRateOfTDS
        End Get
        Set(ByVal vData As Double)
            mvarRateOfTDS = vData
        End Set
    End Property



            'Public Property Let IsBookEntry(ByVal vData As Boolean)   'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.IsBookEntry = 5
            '    mvarIsBookEntry = vData
            'End Property


    Public Property IsBookEntry As Boolean
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.IsBookEntry
            IsBookEntry = mvarIsBookEntry
        End Get
        Set(ByVal vData As Boolean)
            mvarIsBookEntry = vData
        End Set
    End Property



            'Public Property Let DtOfPayment(ByVal vData As Date)          'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.DtOfPayment = 5
            '    mvarDtOfPayment = vData
            'End Property


    Public Property DtOfPayment As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfPayment
            DtOfPayment = mvarDtOfPayment
        End Get
        Set(ByVal vData As Date)
            mvarDtOfPayment = vData
        End Set
    End Property



            'Public Property Let AmtOfPayment(ByVal vData As Double)   'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.AmtOfPayment = 5
            '    mvarAmtOfPayment = vData
            'End Property


    Public Property AmtOfPayment As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AmtOfPayment
            AmtOfPayment = mvarAmtOfPayment
        End Get
        Set(ByVal vData As Double)
            mvarAmtOfPayment = vData
        End Set
    End Property



            'Public Property Let DCode(ByVal vData As String)
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.DCode = 5
            '    mvarDCode = vData
            'End Property


    Public Property DCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DCode
            DCode = mvarDCode
        End Get
        Set(ByVal vData As String)
            mvarDCode = vData
        End Set
    End Property



            'Public Property Let did(ByVal vData As Long)    'payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.DId = 5
            '    mvarDId = vData
            'End Property


    Public Property did As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            did = mvarDId
        End Get
        Set(ByVal vData As Long)
            mvarDId = vData
        End Set
    End Property



            'Public Property Let RetnID(ByVal vData As Long)
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.RetnID = 5
            '    mvarRetnID = vData
            'End Property


    Public Property RetnID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RetnID
            RetnID = mvarRetnID
        End Get
        Set(ByVal vData As Long)
            mvarRetnID = vData
        End Set
    End Property



            'Public Property Let ID27Q(ByVal vData As Long)   'Payal
            '    'used when assigning a value to the property, on the left side of an assignment.
            '    'Syntax: X.ID26Q = 5
            '    mvarID27Q = vData
            'End Property


    Public Property ID27Q As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ID26Q
            ID27Q = mvarID27Q
        End Get
        Set(ByVal vData As Long)
            mvarID27Q = vData
        End Set
    End Property


    Public Function getChallanNo(ByVal ID As Long) As String
        Dim nds As New DataSet, dt As Date
        nds = FetchDataSet("select BankChallanNo,DtOfChallan, TranVouNo from Challan27Q WHERE challanID=" & ID)
        If nds.Tables(0).Rows.Count > 0 Then
            dt = nds.Tables(0).Rows(0)(1).ToString()

            If nds.Tables(0).Rows(0)(0).ToString() = "" Then
                getChallanNo = SetFormat("00000", nds.Tables(0).Rows(0)(2).ToString()) & " - " & Format(dt, "dd/MM/yy")

            Else

                getChallanNo = SetFormat("00000", nds.Tables(0).Rows(0)(0).ToString()) & " - " & Format(dt, "dd/MM/yy")
            End If
        Else
            getChallanNo = vbNullString
        End If
    End Function
    Public Function GetChallanID(ByVal BChlNo As Long, ChlDt As Date) As Long

        Dim nds As New DataSet

        nds = FetchDataSet("SELECT challanid FROM Challan27Q WHERE BankChallanNo=" & BChlNo & " and DtOfChallan=#" & Format(ChlDt, "MM/dd/yyyy") & "#")
        If nds.Tables(0).Rows.Count > 0 Then
            GetChallanID = nds.Tables(0).Rows(0)(0)
        Else
            GetChallanID = 0
        End If
        nds.Dispose()

    End Function
End Class
