Public Class ClsSalaryDetail24QObj
    'local variable(s) to hold property value(s)
    Private mvarSDID As Long 'local copy
    Private mvarRetnID As Long 'local copy
    Private mvarDId As Long 'local copy
    Private mvarEmpFrDt As Date 'local copy
    Private mvarEmpToDt As Date 'local copy
    Private mvarTotalSalary As Double 'local copy
    Private mvarSec16ii As Double 'local copy
    Private mvarSec16iii As Double 'local copy
    Private mvarOtherIncome As Double 'local copy
    Private mvarSec80CCEAmt As Double 'local copy
    Private mvarSec80CCFAmt As Double 'local copy
    Private mvarSec80CCGAmt As Double 'local copy
    Private mvarOtherVIA As Double 'local copy
    Private mvarTaxAmt As Double 'local copy
    Private mvarSurcharge As Double 'local copy
    Private mvarECess As Double 'local copy
    Private mvarRelief89 As Double 'local copy
    Private mvarTDSAmt As Double 'local copy
    Public Event AfterDelete()
    Public Event AfterSave()
    Public Event BeforeDelete(Cancel As Boolean)
    Public Event BeforeSave(Cancel As Boolean)
    Public Event Fetch(ByVal ID As Long)
    Public Event PrepareDataForSave(Cancel As Boolean)
    'local variable(s) to hold property value(s)
    Private mvarTDSAmtPreEmp As Double 'local copy
    Private mvarTotalSalaryPreEmp As Double 'local copy
    Private mvarHighRatePAN As Boolean 'local copy

    Public Property HighRatePAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            HighRatePAN = mvarHighRatePAN
        End Get
        Set(ByVal vData As String)
            mvarHighRatePAN = vData
        End Set
    End Property


    Public Property TDSAmtPreEmp As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            TDSAmtPreEmp = mvarTDSAmtPreEmp
        End Get
        Set(ByVal vData As String)
            mvarTDSAmtPreEmp = vData
        End Set
    End Property

    Public Property TotalSalaryPreEmp As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            TotalSalaryPreEmp = mvarTotalSalaryPreEmp
        End Get
        Set(ByVal vData As String)
            mvarTotalSalaryPreEmp = vData
        End Set
    End Property

    Public Property TDSAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            TDSAmt = mvarTDSAmt
        End Get
        Set(ByVal vData As String)
            mvarTDSAmt = vData
        End Set
    End Property

    Public Property Relief89 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Relief89 = mvarRelief89
        End Get
        Set(ByVal vData As String)
            mvarRelief89 = vData
        End Set
    End Property


    Public Property ECess As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            ECess = mvarECess
        End Get
        Set(ByVal vData As String)
            mvarECess = vData
        End Set
    End Property

    Public Property Surcharge As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Surcharge = mvarSurcharge
        End Get
        Set(ByVal vData As String)
            mvarSurcharge = vData
        End Set
    End Property


    Public Property TaxAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            TaxAmt = mvarTaxAmt
        End Get
        Set(ByVal vData As String)
            mvarTaxAmt = vData
        End Set
    End Property

    Public Property OtherVIA As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            OtherVIA = mvarOtherVIA
        End Get
        Set(ByVal vData As String)
            mvarOtherVIA = vData
        End Set
    End Property

    Public Property Sec80CCEAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Sec80CCEAmt = mvarSec80CCEAmt
        End Get
        Set(ByVal vData As String)
            mvarSec80CCEAmt = vData
        End Set
    End Property

    Public Property Sec80CCFAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Sec80CCFAmt = mvarSec80CCFAmt
        End Get
        Set(ByVal vData As String)
            mvarSec80CCFAmt = vData
        End Set
    End Property

    Public Property Sec80CCGAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Sec80CCGAmt = mvarSec80CCGAmt
        End Get
        Set(ByVal vData As String)
            mvarSec80CCGAmt = vData
        End Set
    End Property

    Public Property OtherIncome As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            OtherIncome = mvarOtherIncome
        End Get
        Set(ByVal vData As String)
            mvarOtherIncome = vData
        End Set
    End Property

    Public Property Sec16ii As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Sec16ii = mvarSec16ii
        End Get
        Set(ByVal vData As String)
            mvarSec16ii = vData
        End Set
    End Property

    Public Property Sec16iii As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Sec16iii = mvarSec16iii
        End Get
        Set(ByVal vData As String)
            mvarSec16iii = vData
        End Set
    End Property

    Public Property TotalSalary As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            TotalSalary = mvarTotalSalary
        End Get
        Set(ByVal vData As String)
            mvarTotalSalary = vData
        End Set
    End Property


    Public Property EmpToDt As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            EmpToDt = mvarEmpToDt
        End Get
        Set(ByVal vData As Date)
            mvarEmpToDt = vData
        End Set
    End Property


    Public Property EmpFrDt As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            EmpFrDt = mvarEmpFrDt
        End Get
        Set(ByVal vData As Date)
            mvarEmpFrDt = vData
        End Set
    End Property

    Public Property RetnID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            RetnID = mvarRetnID
        End Get
        Set(ByVal vData As String)
            mvarRetnID = vData
        End Set
    End Property


    Public Property did As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            did = mvarDId
        End Get
        Set(ByVal vData As String)
            mvarDId = vData
        End Set
    End Property

    Public Property SDID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            SDID = mvarSDID
        End Get
        Set(ByVal vData As String)
            mvarSDID = vData
        End Set
    End Property

    Public Function Update(SD24Q As ClsSalaryDetail24QObj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
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
        With SD24Q
            sql = "Update SalaryDetail24Q Set RetnId = " & IIf(.RetnID = 0, 0, .RetnID) & "," _
        & " DId = " & IIf(.did = 0, 0, .did) & "," _
        & " EmpFromDt = #" & Format(.EmpFrDt, "dd/MMM/yyyy") & "#," _
        & " EmpToDt = #" & Format(.EmpToDt, "dd/MMM/yyyy") & "#," _
        & " TotalSalary = " & IIf(.TotalSalary = 0, 0, .TotalSalary) & "," _
        & " Sec16ii = " & IIf(.Sec16ii = 0, 0, .Sec16ii) & "," _
        & " Sec16iii = " & IIf(.Sec16iii = 0, 0, .Sec16iii) & "," _
        & " OtherIncome = " & IIf(.OtherIncome = 0, 0, .OtherIncome) & "," _
        & " Sec80CCEAmt = " & IIf(.Sec80CCEAmt = 0, 0, .Sec80CCEAmt) & "," _
        & " OtherVIA = " & IIf(.OtherVIA = 0, 0, .OtherVIA) & "," _
        & " TaxAmt = " & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
        & " Surcharge  = " & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
        & " ECess  = " & IIf(.ECess = 0, 0, .ECess) & "," _
        & " Relief89  = " & IIf(.Relief89 = 0, 0, .Relief89) & "," _
        & " TDSAmt  = " & IIf(.TDSAmt = 0, 0, .TDSAmt) & "," _
        & " Sec80CCFAmt = " & IIf(.Sec80CCFAmt = 0, 0, .Sec80CCFAmt) & "," _
        & " Sec80CCGAmt = " & IIf(.Sec80CCGAmt = 0, 0, .Sec80CCGAmt) & "," _
        & " TotalSalaryPreEmp = " & IIf(.TotalSalaryPreEmp = 0, 0, .TotalSalaryPreEmp) & "," _
        & " TDSAmtPreEmp = " & IIf(.TDSAmtPreEmp = 0, 0, .TDSAmtPreEmp) & "," _
        & " HighRatePAN = " & IIf(.HighRatePAN = True, vbTrue, vbFalse) & "" _
        & " Where SDID = " & .SDID
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

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(sdid) as Id From salarydetail24q"
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            MaxID = nds.Tables(0).Rows(0)(0)
        Else
            MaxID = 0
        End If
        nds.Dispose()
    End Function

    Public Function Insert(SD24Q As ClsSalaryDetail24QObj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction
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
        With SD24Q
            .SDID = MaxID() + 1
            sql = "Insert Into salarydetail24q (sdid,RetnId,DId,EmpFromDt,EmpToDt,TotalSalary, " _
            & "Sec16ii,Sec16iii,OtherIncome,Sec80CCEAmt,OtherVIA,TaxAmt,Surcharge,ECess, " _
            & "Relief89,TDSAmt,Sec80CCFAmt,Sec80CCGAmt,TotalSalaryPreEmp,TDSAmtPreEmp,HighRatePAN) Values ( " _
            & .SDID & "," & IIf(.RetnID = 0, 0, .RetnID) & "," _
            & IIf(.did = 0, 0, .did) & ",#" _
            & Format(.EmpFrDt, "dd/MMM/yyyy") & "#,#" _
            & Format(.EmpToDt, "dd/MMM/yyyy") & "#," _
            & IIf(.TotalSalary = 0, 0, .TotalSalary) & "," _
            & IIf(.Sec16ii = 0, 0, .Sec16ii) & "," _
            & IIf(.Sec16iii = 0, 0, .Sec16iii) & "," _
            & IIf(.OtherIncome = 0, 0, .OtherIncome) & "," _
            & IIf(.Sec80CCEAmt = 0, 0, .Sec80CCEAmt) & "," _
            & IIf(.OtherVIA = 0, 0, .OtherVIA) & "," _
            & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
            & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
            & IIf(.ECess = 0, 0, .ECess) & "," _
            & IIf(.Relief89 = 0, 0, .Relief89) & "," _
            & IIf(.TDSAmt = 0, 0, .TDSAmt) & "," _
            & IIf(.Sec80CCFAmt = 0, 0, .Sec80CCFAmt) & "," _
            & IIf(.Sec80CCGAmt = 0, 0, .Sec80CCGAmt) & "," _
            & IIf(.TotalSalaryPreEmp = 0, 0, .TotalSalaryPreEmp) & "," _
            & IIf(.TDSAmtPreEmp = 0, 0, .TDSAmtPreEmp) & "," _
            & HighRatePAN & ")"
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
    End Function

    Public Function Delete(ID As Long) As Boolean
        Dim cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
        '        On Error GoTo DelErr
        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then
            Delete = False
            Exit Function
        End If

        Dim sql As String
        sql = "Delete * From SalaryDetail24Q Where SDID = " & ID
        'Cnn.BeginTrans
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
End Class
