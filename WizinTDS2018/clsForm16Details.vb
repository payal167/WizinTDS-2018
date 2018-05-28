Imports System.Data.OleDb

Public Class Form16Details
    'Dim oF16Challan As New clsF16Challan
    'local variable(s) to hold property value(s)
    Private mvarF16ID As Long 'local copy
    Private mvarDId As Long 'local copy
    Private mvarRetnID As Double 'local copy
    Private mvarDDesgn As String 'local copy
    Private mvarEmpFromDt As Date 'local copy
    Private mvarEmpToDt As Date 'local copy
    Private mvarGross1 As Double 'local copy
    Private mvarGross2 As Double 'local copy
    Private mvarGross3 As Double 'local copy
    Private mvarSec16ii As Double 'local copy
    Private mvarSec16iii As Double 'local copy
    Private mvarTaxAmt As Double 'local copy
    Private mvarSurcharge As Double 'local copy
    Private mvarECess As Double 'local copy
    Private mvarRelief89 As Double 'local copy
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent AfterDelete[(arg1, arg2, ... , argn)]
    Public Event AfterDelete()
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent AfterSave[(arg1, arg2, ... , argn)]
    Public Event AfterSave()
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent BeforeDelete[(arg1, arg2, ... , argn)]
    Public Event BeforeDelete(Cancel As Boolean)
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent BeforeSave[(arg1, arg2, ... , argn)]
    Public Event BeforeSave(Cancel As Boolean)
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent PrepareDataForSave[(arg1, arg2, ... , argn)]
    Public Event PrepareDataForSave(Cancel As Boolean)
    'local variable(s) to hold property value(s)
    Private mvarSignByName As String 'local copy
    Private mvarSignByFatherName As String 'local copy
    Private mvarSignByCapacity As String 'local copy
    Private mvarPlaceOfForm As String
    Private mvarDateOfForm As Date
    Private mvarTDSOnPerks As Double
    Private mvarTotalSalaryPreEmp As Double 'local copy
    Private mvarTDSAmtPreEmp As Double 'local copy
    Private mvarHighRatePAN As Boolean 'local copy
    Private mvarDAdd1 As String 'local copy
    Private mvarDAdd2 As String 'local copy
    Private mvarDAdd3 As String 'local copy
    Private mvarDAdd4 As String 'local copy
    Private mvarDAdd5 As String 'local copy
    Private mvarDPan As String 'local copy
    Private mvarDTan As String 'local copy
    Private mvarDPin As String 'local copy
    Private mvarDStatus As String 'local copy
    Private mvarCoName As String 'local copy
    Private mvarPR24Name As String
    Private mvarPR24Desg As String

    Public Property HighRatePAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            HighRatePAN = mvarHighRatePAN
            'Category = mvarCategory
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

    Public Property SignByCapacity As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            SignByCapacity = mvarSignByCapacity
        End Get
        Set(ByVal vData As String)
            mvarSignByCapacity = vData
        End Set
    End Property

    Public Property SignByFatherName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            SignByFatherName = mvarSignByFatherName
        End Get
        Set(ByVal vData As String)
            mvarSignByFatherName = vData
        End Set
    End Property

    Public Property SignByName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            SignByName = mvarSignByName
        End Get
        Set(ByVal vData As String)
            mvarSignByName = vData
        End Set
    End Property

    Public Property mCoAdd5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd5
            mCoAdd5 = mvarDAdd5
        End Get
        Set(ByVal vData As String)
            mvarDAdd5 = vData
        End Set
    End Property

    Public Property mCoAdd4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd4
            mCoAdd4 = mvarDAdd4
        End Get
        Set(ByVal vData As String)
            mvarDAdd4 = vData
        End Set
    End Property

    Public Property mCoAdd3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd3
            mCoAdd3 = mvarDAdd3
        End Get
        Set(ByVal vData As String)
            mvarDAdd3 = vData
        End Set
    End Property

    Public Property mCoAdd2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd2
            mCoAdd2 = mvarDAdd2
        End Get
        Set(ByVal vData As String)
            mvarDAdd2 = vData
        End Set
    End Property

    Public Property mCoAdd1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd1
            mCoAdd1 = mvarDAdd1
        End Get
        Set(ByVal vData As String)
            mvarDAdd1 = vData
        End Set
    End Property

    Public Property mCoPin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            mCoPin = mvarDPin
        End Get
        Set(ByVal vData As String)
            mvarDPin = vData
        End Set
    End Property

    Public Property mCoPAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            mCoPAN = mvarDPin
        End Get
        Set(ByVal vData As String)
            mvarDPin = vData
        End Set
    End Property

    Public Property mCoTAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            mCoTAN = mvarDPin
        End Get
        Set(ByVal vData As String)
            mvarDPin = vData
        End Set
    End Property

    Public Property PR24Desg As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            PR24Desg = mvarPR24Desg
        End Get
        Set(ByVal vData As String)
            mvarPR24Desg = vData
        End Set
    End Property

    Public Property mCoName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DName
            mCoName = mvarCoName
        End Get
        Set(ByVal vData As String)
            mvarCoName = vData
        End Set
    End Property

    Public Property mCoStatus As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            mCoStatus = mvarDStatus
        End Get
        Set(ByVal vData As String)
            mvarDStatus = vData
        End Set
    End Property

    Public Function Update(ByVal F16Details As Form16Details, Allowances As Collection_Allowances,
    OtherIncomes As Collection_OtherIncomes, Sec80CDeductions As Collection_Sec80CDed,
    sec80CCFDeductions As Collection_Sec80CCFDed, sec80CCGDeductions As Collection_Sec80CCGDed,
    VI_A_Deductions As Collection_VI_A_Deductions, grdF16 As DataGridView, TotalSalaryPreEmp As Double, TDSAmtPreEmp As Double, HighRatePAN As Boolean) As Boolean
        ' Public Function Update(F16Details As clsForm16Details, Allowances As Collection_Allowances,
        'OtherIncomes As Collection_OtherIncomes, Sec80CDeductions As Collection_Sec80CDed,
        'sec80CCFDeductions As Collection_Sec80CCFDed, sec80CCGDeductions As Collection_Sec80CCGDed,
        'VI_A_Deductions As Collection_VI_A_Deductions, grdF16 As Variant, TotalSalaryPreEmp As Double, TDSAmtPreEmp As Double, HighRatePAN As Boolean) As Boolean
        On Error GoTo UpErr
        Dim sql As String, cnl As Boolean
        Dim cmd As New OleDbCommand
        Dim norow As Integer
        Dim mFId As Long 'MoreDetails As New clsForm16MoreDetails
        Dim sql_OthIncomes As String, sql_Allw As String, sql_sec80C As String, sql_sec80CCF As String, sql_6A As String
        Dim sql_sec80CCG As String
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then Exit Function
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then Exit Function
        With F16Details
            sql = "Update Form16Details Set RetnId = " & IIf(.RetnID = 0, 0, .RetnID) & "," _
        & " DId = " & IIf(.did = vbNullString, "Null", "'" & .did & "'") & "," _
        & " DDesgn = '" & IIf(Trim(.DDesgn) = "", "", .DDesgn) & "'," _
        & " EmpFromDt = #" & Format(.EmpFromDt, "dd/MMM/yyyy") & "#," _
        & " EmpToDt = #" & Format(.EmpToDt, "dd/MMM/yyyy") & "#," _
        & " Gross1 = " & IIf(.Gross1 = 0, 0, .Gross1) & "," _
        & " Gross2 = " & IIf(.Gross2 = 0, 0, .Gross2) & "," _
        & " Gross3 = " & IIf(.Gross3 = 0, 0, .Gross3) & "," _
        & " Sec16ii = " & IIf(.Sec16ii = 0, 0, .Sec16ii) & "," _
        & " Sec16iii = " & IIf(.Sec16iii = 0, 0, .Sec16iii) & "," _
        & " TaxAmt = " & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
        & " Surcharge  = " & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
        & " ECess  = " & IIf(.ECess = 0, 0, .ECess) & "," _
        & " Relief89  = " & IIf(.Relief89 = 0, 0, .Relief89) & "," _
        & " TDSOnPerks = " & IIf(.TDSOnPerks = 0, 0, .TDSOnPerks) & "," _
        & " TotalSalaryPreEmp = " & IIf(.TotalSalaryPreEmp = 0, 0, .TotalSalaryPreEmp) & "," _
        & " TDSAmtPreEmp = " & IIf(.TDSAmtPreEmp = 0, 0, .TDSAmtPreEmp) & "," _
        & " HighRatePAN = " & IIf(.HighRatePAN = True, vbYes, vbNo) _
        & " Where F16ID =" & .F16ID
            mFId = .F16ID
            cmd.CommandText = sql
            cmd.Connection = cn
            norow = cmd.ExecuteNonQuery()
            If norow > 0 Then
                Update = True
                RaiseEvent AfterSave()
            Else
                Update = False
            End If
        End With
        Exit Function
        'mFId = .F16ID
        '    Cnn.Execute sql
        ''Following Changes made by Prakash
        '    '        'Delete records before update
        '    sql = "Delete * From Form16MoreDetails Where F16ID = " & .F16ID
        '    Cnn.Execute sql
        'sql = "Delete * From F16Challan Where F16ID = " & .F16ID
        '    Cnn.Execute sql

        'Master Data written sucessfully, now write Allowances Data of Form 16
        'Dim Allw As Long
        'For Allw = 1 To Allowances.Count
        '    Allowances(Allw).ID = MoreDetails.MaxID + 1
        '    sql_Allw = "Insert into Form16MoreDetails (ID, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
        '    " Values ( " _
        '    & Allowances(Allw).ID & "," _
        '    & .F16ID & "," _
        '    & "'A','" _
        '    & IIf(Trim(Allowances(Allw).Particulars) = "", "", Allowances(Allw).Particulars) & "'," _
        '    & IIf(Allowances(Allw).GrossAmt = 0, 0, Allowances(Allw).GrossAmt) & "," _
        '    & IIf(Allowances(Allw).QualifyAmt = 0, 0, Allowances(Allw).QualifyAmt) & "," _
        '    & IIf(Allowances(Allw).DeductibleAmt = 0, 0, Allowances(Allw).DeductibleAmt) & ")"
        '    Cnn.Execute sql_Allw
        '    Next
        ''Allowances Data written sucessfully, now write Other Income Data of Form 16
        'Dim OthInc As Long
        'For OthInc = 1 To OtherIncomes.Count
        '    OtherIncomes(OthInc).ID = MoreDetails.MaxID + 1
        '    sql_OthInc = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
        '    " Values ( " _
        '    & OtherIncomes(OthInc).ID & "," _
        '    & .F16ID & "," _
        '    & "'O','" _
        '    & IIf(Trim(OtherIncomes(OthInc).Particulars) = "", "", OtherIncomes(OthInc).Particulars) & "'," _
        '    & IIf(OtherIncomes(OthInc).GrossAmt = 0, 0, OtherIncomes(OthInc).GrossAmt) & "," _
        '    & IIf(OtherIncomes(OthInc).QualifyAmt = 0, 0, OtherIncomes(OthInc).QualifyAmt) & "," _
        '    & IIf(OtherIncomes(OthInc).DeductibleAmt = 0, 0, OtherIncomes(OthInc).DeductibleAmt) & ")"
        '    Cnn.Execute sql_OthInc
        '    Next
        ''Other Income written sucessfully, now write Section 80C Data of Form 16
        'Dim Sec80C As Long
        'For Sec80C = 1 To Sec80CDeductions.Count
        '    Sec80CDeductions(Sec80C).ID = MoreDetails.MaxID + 1
        '    sql_sec80C = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
        '    " Values ( " _
        '    & Sec80CDeductions(Sec80C).ID & "," _
        '    & .F16ID & "," _
        '    & "'E','" _
        '    & IIf(Trim(Sec80CDeductions(Sec80C).Particulars) = "", "", Sec80CDeductions(Sec80C).Particulars) & "'," _
        '    & IIf(Sec80CDeductions(Sec80C).GrossAmt = 0, 0, Sec80CDeductions(Sec80C).GrossAmt) & "," _
        '    & IIf(Sec80CDeductions(Sec80C).QualifyAmt = 0, 0, Sec80CDeductions(Sec80C).QualifyAmt) & "," _
        '    & IIf(Sec80CDeductions(Sec80C).DeductibleAmt = 0, 0, Sec80CDeductions(Sec80C).DeductibleAmt) & ")"
        '    Cnn.Execute sql_sec80C
        '    Next
        ''Section 80C written sucessfully, now write Section 80CCF Data of Form 16
        'Dim Sec80CCF As Long
        'For Sec80CCF = 1 To sec80CCFDeductions.Count
        '    sec80CCFDeductions(Sec80CCF).ID = MoreDetails.MaxID + 1
        '    sql_sec80CCF = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
        '    " Values ( " _
        '    & sec80CCFDeductions(Sec80CCF).ID & "," _
        '    & .F16ID & "," _
        '    & "'F','" _
        '    & IIf(Trim(sec80CCFDeductions(Sec80CCF).Particulars) = "", "", sec80CCFDeductions(Sec80CCF).Particulars) & "'," _
        '    & IIf(sec80CCFDeductions(Sec80CCF).GrossAmt = 0, 0, sec80CCFDeductions(Sec80CCF).GrossAmt) & "," _
        '    & IIf(sec80CCFDeductions(Sec80CCF).QualifyAmt = 0, 0, sec80CCFDeductions(Sec80CCF).QualifyAmt) & "," _
        '    & IIf(sec80CCFDeductions(Sec80CCF).DeductibleAmt = 0, 0, sec80CCFDeductions(Sec80CCF).DeductibleAmt) & ")"
        '    Cnn.Execute sql_sec80CCF
        '    Next
        ''Section 80CCF written sucessfully, now write Section 80CCG Data of Form 16
        'Dim Sec80CCG As Long
        'For Sec80CCG = 1 To sec80CCGDeductions.Count
        '    sec80CCGDeductions(Sec80CCG).ID = MoreDetails.MaxID + 1
        '    sql_sec80CCG = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
        '    " Values ( " _
        '    & sec80CCGDeductions(Sec80CCG).ID & "," _
        '    & .F16ID & "," _
        '    & "'G','" _
        '    & IIf(Trim(sec80CCGDeductions(Sec80CCG).Particulars) = "", "", sec80CCGDeductions(Sec80CCG).Particulars) & "'," _
        '    & IIf(sec80CCGDeductions(Sec80CCG).GrossAmt = 0, 0, sec80CCGDeductions(Sec80CCG).GrossAmt) & "," _
        '    & IIf(sec80CCGDeductions(Sec80CCG).QualifyAmt = 0, 0, sec80CCGDeductions(Sec80CCG).QualifyAmt) & "," _
        '    & IIf(sec80CCGDeductions(Sec80CCG).DeductibleAmt = 0, 0, sec80CCGDeductions(Sec80CCG).DeductibleAmt) & ")"

        '    Cnn.Execute sql_sec80CCG
        '    Next

        ''Sec 80CCG data written sucessfully, now write Chapter VI-A Data of Form 16
        'Dim Chp6A As Long
        'For Chp6A = 1 To VI_A_Deductions.Count
        '    VI_A_Deductions(Chp6A).ID = MoreDetails.MaxID + 1
        '    sql_6A = "Insert into Form16MoreDetails (id,F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt)" &
        '    " Values ( " _
        '    & VI_A_Deductions(Chp6A).ID & "," _
        '    & .F16ID & "," _
        '    & "'V','" _
        '    & IIf(Trim(VI_A_Deductions(Chp6A).Particulars) = "", "", VI_A_Deductions(Chp6A).Particulars) & "'," _
        '    & IIf(VI_A_Deductions(Chp6A).GrossAmt = 0, 0, VI_A_Deductions(Chp6A).GrossAmt) & "," _
        '    & IIf(VI_A_Deductions(Chp6A).QualifyAmt = 0, 0, VI_A_Deductions(Chp6A).QualifyAmt) & "," _
        '    & IIf(VI_A_Deductions(Chp6A).DeductibleAmt = 0, 0, VI_A_Deductions(Chp6A).DeductibleAmt) & ")"
        '    Cnn.Execute sql_6A
        '    Next

        'Dim R As Integer
        'For R = 1 To grdF16.Rows - 2
        '    If grdF16.ValueMatrix(R, 1) > 0 Then
        '        oF16Challan.Insert.F16ID, grdF16.ValueMatrix(R, 1), grdF16.ValueMatrix(R, 2), grdF16.ValueMatrix(R, 3), grdF16.ValueMatrix(R, 5), grdF16.ValueMatrix(R, 8), grdF16.ValueMatrix(R, 6), grdF16.TextMatrix(R, 7), vbNullString
        '        End If
        'Next R
UpErr:
        MsgBox(Err.Description, , Err.Number)
    End Function

    Public Function MaxID() As Long
        Dim nds As New DataSet
        nds = FetchDataSet("Select Max(F16Id) as Id From Form16Details")
        If Not String.IsNullOrEmpty(nds.Tables(0).Rows(0)("id").Value) Then
            MaxID = nds.Tables(0).Rows(0)("id").Value
        Else
            MaxID = 0
        End If

        nds = Nothing
    End Function

    Public Function Insert(F16Details As Form16Details, Allowances As Collection_Allowances,
    OtherIncomes As Collection_OtherIncomes, Sec80CDeductions As Collection_Sec80CDed,
    sec80CCFDeductions As Collection_Sec80CCFDed, sec80CCGDeductions As Collection_Sec80CCGDed,
    VI_A_Deductions As Collection_VI_A_Deductions, grdF16 As DataGridView, TotalSalaryPreEmp As Double, TDSAmtPreEmp As Double, HighRatePAN As Boolean) As Boolean
        '    , Allowances As Collection_Allowances,
        'OtherIncomes As Collection_OtherIncomes, Sec80CDeductions As Collection_Sec80CDed,
        'sec80CCFDeductions As Collection_Sec80CCFDed, sec80CCGDeductions As Collection_Sec80CCGDed,
        'VI_A_Deductions As Collection_VI_A_Deductions, grdF16 As Variant, TotalSalaryPreEmp As Double, TDSAmtPreEmp As Double, HighRatePAN As Boolean) As Boolean
        'On Error GoTo InErr
        Dim sql As String ', cnl As Boolean
        'Dim sql_OthInc As String, sql_Allw As String, sql_sec80C As String, sql_sec80CCF As String, sql_6A As String
        'Dim sql_sec80CCG As String
        ' Dim MoreDetails As New clsForm16MoreDetails
        'cnl = False : RaiseEvent BeforeSave(cnl)
        'If cnl = True Then Exit Function
        'cnl = False : RaiseEvent PrepareDataForSave(cnl)
        'If cnl = True Then Exit Function
        'Master Data of Form 16
        With F16Details
            .F16ID = .MaxID + 1
            sql = "Insert Into Form16Details (F16Id,DId,RetnID,DDesgn,EmpFromDt,EmpToDt,Gross1,Gross2,Gross3,Sec16ii,Sec16iii, " _
        & "TaxAmt,Surcharge,ECess,Relief89,TDSOnPerks,TotalSalaryPreEmp,TDSAmtPreEmp,HighRatePAN) Values ( " _
        & .F16ID & "," _
        & IIf(.did = 0, 0, .did) & "," _
        & IIf(.RetnID = 0, 0, .RetnID) & ",'" _
        & IIf(Trim(.DDesgn) = "", "", .DDesgn) & "', #" _
        & Format(.EmpFromDt, "dd/MMM/yyyy") & "#,#" _
        & Format(.EmpToDt, "dd/MMM/yyyy") & "#," _
        & IIf(.Gross1 = 0, 0, .Gross1) & "," _
        & IIf(.Gross2 = 0, 0, .Gross2) & "," _
        & IIf(.Gross3 = 0, 0, .Gross3) & "," _
        & IIf(.Sec16ii = 0, 0, .Sec16ii) & "," _
        & IIf(.Sec16iii = 0, 0, .Sec16iii) & "," _
        & IIf(.TaxAmt = 0, 0, .TaxAmt) & "," _
        & IIf(.Surcharge = 0, 0, .Surcharge) & "," _
        & IIf(.ECess = 0, 0, .ECess) & "," _
        & IIf(.Relief89 = 0, 0, .Relief89) & "," _
        & IIf(.TDSOnPerks = 0, 0, .TDSOnPerks) & "," _
        & IIf(.TotalSalaryPreEmp = 0, 0, .TotalSalaryPreEmp) & "," _
        & IIf(.TDSAmtPreEmp = 0, 0, .TDSAmtPreEmp) & "," _
        & IIf(.HighRatePAN = True, vbTrue, vbFalse) & ")"

            Dim cmd As New OleDbCommand
            Try
                cmd.CommandText = sql
                cmd.Connection = cn
                cmd.ExecuteNonQuery()
                Insert = True
            Catch ex As Exception
                Dim merror As String
                merror = ex.Message
                MsgBox(merror)
                Insert = False
            End Try
        End With
    End Function

    'Master Data written sucessfully, now write Allowances Data of Form 16
    '    Dim Allw As Long
    '            For Allw = 1 To Allowances.Count
    '                Allowances(Allw).ID = MoreDetails.MaxID + 1
    '                sql_Allw = "Insert into Form16MoreDetails (ID, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
    '                " Values ( " _
    '                & Allowances(Allw).ID & "," _
    '                & .F16ID & "," _
    '                & "'A','" _
    '                & IIf(Trim(Allowances(Allw).Particulars) = "", "", Allowances(Allw).Particulars) & "'," _
    '                & IIf(Allowances(Allw).GrossAmt = 0, 0, Allowances(Allw).GrossAmt) & "," _
    '                & IIf(Allowances(Allw).QualifyAmt = 0, 0, Allowances(Allw).QualifyAmt) & "," _
    '                & IIf(Allowances(Allw).DeductibleAmt = 0, 0, Allowances(Allw).DeductibleAmt) & ")"
    '                Cnn.Execute sql_Allw
    '        Next
    '            'Allowances Data written sucessfully, now write Other Income Data of Form 16
    '            Dim OthInc As Long
    '            For OthInc = 1 To OtherIncomes.Count
    '                OtherIncomes(OthInc).ID = MoreDetails.MaxID + 1
    '                sql_OthInc = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
    '                " Values ( " _
    '                & OtherIncomes(OthInc).ID & "," _
    '                & .F16ID & "," _
    '                & "'O','" _
    '                & IIf(Trim(OtherIncomes(OthInc).Particulars) = "", "", OtherIncomes(OthInc).Particulars) & "'," _
    '                & IIf(OtherIncomes(OthInc).GrossAmt = 0, 0, OtherIncomes(OthInc).GrossAmt) & "," _
    '                & IIf(OtherIncomes(OthInc).QualifyAmt = 0, 0, OtherIncomes(OthInc).QualifyAmt) & "," _
    '                & IIf(OtherIncomes(OthInc).DeductibleAmt = 0, 0, OtherIncomes(OthInc).DeductibleAmt) & ")"
    '                Cnn.Execute sql_OthInc
    '        Next
    '            'Other Income written sucessfully, now write Section 80C Data of Form 16
    '            Dim Sec80C As Long
    '            For Sec80C = 1 To Sec80CDeductions.Count
    '                Sec80CDeductions(Sec80C).ID = MoreDetails.MaxID + 1
    '                sql_sec80C = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
    '                " Values ( " _
    '                & Sec80CDeductions(Sec80C).ID & "," _
    '                & .F16ID & "," _
    '                & "'E','" _
    '                & IIf(Trim(Sec80CDeductions(Sec80C).Particulars) = "", "", Sec80CDeductions(Sec80C).Particulars) & "'," _
    '                & IIf(Sec80CDeductions(Sec80C).GrossAmt = 0, 0, Sec80CDeductions(Sec80C).GrossAmt) & "," _
    '                & IIf(Sec80CDeductions(Sec80C).QualifyAmt = 0, 0, Sec80CDeductions(Sec80C).QualifyAmt) & "," _
    '                & IIf(Sec80CDeductions(Sec80C).DeductibleAmt = 0, 0, Sec80CDeductions(Sec80C).DeductibleAmt) & ")"
    '                Cnn.Execute sql_sec80C
    '        Next
    '            'Sec 80C data written sucessfully, now write 80CCF Data of Form 16
    '            Dim Sec80CCF As Long
    '            For Sec80CCF = 1 To sec80CCFDeductions.Count
    '                sec80CCFDeductions(Sec80CCF).ID = MoreDetails.MaxID + 1
    '                sql_sec80CCF = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
    '                " Values ( " _
    '                & sec80CCFDeductions(Sec80CCF).ID & "," _
    '                & .F16ID & "," _
    '                & "'F','" _
    '                & IIf(Trim(sec80CCFDeductions(Sec80CCF).Particulars) = "", "", sec80CCFDeductions(Sec80CCF).Particulars) & "'," _
    '                & IIf(sec80CCFDeductions(Sec80CCF).GrossAmt = 0, 0, sec80CCFDeductions(Sec80CCF).GrossAmt) & "," _
    '                & IIf(sec80CCFDeductions(Sec80CCF).QualifyAmt = 0, 0, sec80CCFDeductions(Sec80CCF).QualifyAmt) & "," _
    '                & IIf(sec80CCFDeductions(Sec80CCF).DeductibleAmt = 0, 0, sec80CCFDeductions(Sec80CCF).DeductibleAmt) & ")"
    '                Cnn.Execute sql_sec80CCF
    '        Next
    '            'Sec 80CCF data written sucessfully, now write 80CCG Data of Form 16
    '            Dim Sec80CCG As Long
    '            For Sec80CCG = 1 To sec80CCGDeductions.Count
    '                sec80CCGDeductions(Sec80CCG).ID = MoreDetails.MaxID + 1
    '                sql_sec80CCG = "Insert into Form16MoreDetails (id, F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt) " &
    '                " Values ( " _
    '                & sec80CCGDeductions(Sec80CCG).ID & "," _
    '                & .F16ID & "," _
    '                & "'G','" _
    '                & IIf(Trim(sec80CCGDeductions(Sec80CCG).Particulars) = "", "", sec80CCGDeductions(Sec80CCG).Particulars) & "'," _
    '                & IIf(sec80CCGDeductions(Sec80CCG).GrossAmt = 0, 0, sec80CCGDeductions(Sec80CCG).GrossAmt) & "," _
    '                & IIf(sec80CCGDeductions(Sec80CCG).QualifyAmt = 0, 0, sec80CCGDeductions(Sec80CCG).QualifyAmt) & "," _
    '                & IIf(sec80CCGDeductions(Sec80CCG).DeductibleAmt = 0, 0, sec80CCGDeductions(Sec80CCG).DeductibleAmt) & ")"
    '                Cnn.Execute sql_sec80CCG
    '        Next
    '            'Sec 80CCG data written sucessfully, now write Chapter VI-A Data of Form 16
    '            Dim Chp6A As Long
    '            For Chp6A = 1 To VI_A_Deductions.Count
    '                VI_A_Deductions(Chp6A).ID = MoreDetails.MaxID + 1
    '                sql_6A = "Insert into Form16MoreDetails (id,F16ID,TypeOfDetail,Particulars,GrossAmt,QualifyAmt,DeductibleAmt)" &
    '                " Values ( " _
    '                & VI_A_Deductions(Chp6A).ID & "," _
    '                & .F16ID & "," _
    '                & "'V','" _
    '                & IIf(Trim(VI_A_Deductions(Chp6A).Particulars) = "", "", VI_A_Deductions(Chp6A).Particulars) & "'," _
    '                & IIf(VI_A_Deductions(Chp6A).GrossAmt = 0, 0, VI_A_Deductions(Chp6A).GrossAmt) & "," _
    '                & IIf(VI_A_Deductions(Chp6A).QualifyAmt = 0, 0, VI_A_Deductions(Chp6A).QualifyAmt) & "," _
    '                & IIf(VI_A_Deductions(Chp6A).DeductibleAmt = 0, 0, VI_A_Deductions(Chp6A).DeductibleAmt) & ")"
    '                Cnn.Execute sql_6A
    '        Next

    '            Dim R As Integer
    '            For R = 1 To grdF16.Rows - 2
    '                If grdF16.ValueMatrix(R, 1) > 0 Then
    '                    oF16Challan.Insert.F16ID, grdF16.ValueMatrix(R, 1), grdF16.ValueMatrix(R, 2), grdF16.ValueMatrix(R, 3), grdF16.ValueMatrix(R, 5), grdF16.ValueMatrix(R, 8), grdF16.ValueMatrix(R, 6), grdF16.TextMatrix(R, 7), vbNullString
    '            End If
    '            Next R

    '            RaiseEvent AfterSave()
    '            Insert = True
    '        End With
    '        Exit Function
    'InErr:
    '        MsgBox Err.Description, , Err.Number
    '    Insert = False
    '    End Function

    Public Function Delete(ByVal ID As String) As Boolean
        Dim cnl As Boolean
        Dim cmd As New OleDbCommand
        On Error GoTo DelErr
        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then Exit Function
        Dim sql, norow As String
        sql = "Delete * From Form16Details Where F16ID = " & ID
        cmd.CommandText = sql
        cmd.Connection = cn
        norow = cmd.ExecuteNonQuery()
        sql = "Delete * From Form16MoreDetails WHERE F16ID = " & ID
        cmd.CommandText = sql
        cmd.Connection = cn
        norow = cmd.ExecuteNonQuery()
        If norow > 0 Then
            Delete = True
            RaiseEvent AfterDelete()
        Else
            Delete = False
        End If
        Exit Function
DelErr:
        MsgBox(Err.Description, , Err.Number)
        Delete = False
    End Function


    Public Function Fetch(ByVal dname As String) As Form16Details
        Dim DM As New Form16Details
        Dim ds As New DataSet
        ds = FetchDataSet("SELECT CoMst.*, DeductMst.DId, DeductMst.DName, Form16Details.*
FROM CoMst INNER JOIN (DeductMst INNER JOIN Form16Details ON DeductMst.DId = Form16Details.DId) ON CoMst.CoID = DeductMst.CoID where DName like'" & dname & "'")
        'ds = FetchDataSet("select * From form16details")
        If ds.Tables(0).Rows.Count > 0 Then
            With DM
                .SignByFatherName = ds.Tables(0).Rows(0)("SignByFatherName").ToString() 'rst!PR24Name
                .PlaceOfForm = ds.Tables(0).Rows(0)("Place").ToString() 'rst!PR24Desg
                .SignByName = ds.Tables(0).Rows(0)("SignByName").ToString()
                .DateOfForm = ds.Tables(0).Rows(0)("DateOfCertificate").ToString()
                .DDesgn = ds.Tables(0).Rows(0)("DDesgn").ToString()
                .SignByCapacity = ds.Tables(0).Rows(0)("SignByCapacity").ToString()
                .PR24Name = ds.Tables(0).Rows(0)("PR24Name")
                .PR24Desg = ds.Tables(0).Rows(0)("PR24Desg")
                '  .DateOfForm = ds.Tables(0).Rows(0)("")
                .mCoName = ds.Tables(0).Rows(0)("CoName").ToString() 'IIf(IsNull(rst!CoName), vbNullString, rst!CoName)
                .mCoAdd1 = ds.Tables(0).Rows(0)("CoAdd1").ToString() 'IIf(IsNull(rst!CoAdd1), vbNullString, rst!CoAdd1)
                .mCoAdd2 = ds.Tables(0).Rows(0)("CoAdd2").ToString() 'IIf(IsNull(rst!CoAdd2), vbNullString, rst!CoAdd2)
                .mCoAdd3 = ds.Tables(0).Rows(0)("CoAdd3").ToString() 'IIf(IsNull(rst!CoAdd3), vbNullString, rst!CoAdd3)
                .mCoAdd4 = ds.Tables(0).Rows(0)("CoAdd4").ToString() 'IIf(IsNull(rst!CoAdd4), vbNullString, rst!CoAdd4)
                .mCoAdd5 = ds.Tables(0).Rows(0)("CoAdd5").ToString() 'IIf(IsNull(rst!CoAdd5), vbNullString, rst!CoAdd5)
                .mCoPin = ds.Tables(0).Rows(0)("CoPin").ToString() 'IIf(IsNull(rst!CoPin), vbNullString, rst!CoPin)
                .mCoPAN = ds.Tables(0).Rows(0)("CoPAN").ToString() 'IIf(IsNull(rst!CoPAN), vbNullString, rst!CoPAN)
                .mCoTAN = ds.Tables(0).Rows(0)("CoTAN").ToString() 'IIf(IsNull(rst!CoTAN), vbNullString, rst!CoTAN)
                .mCoStatus = ds.Tables(0).Rows(0)("CoStatus").ToString() 'IIf(IsNull(rst!CoStatus), vbNullString, rst!CoStatus)
                .RetnID = ds.Tables(0).Rows(0)("RetnID")
                .F16ID = ds.Tables(0).Rows(0)("F16ID")
                .did = ds.Tables(0).Rows(0)("DeductMst.DId")
                .EmpFromDt = ds.Tables(0).Rows(0)("EmpFromDt")
                .EmpToDt = ds.Tables(0).Rows(0)("EmpToDt")
                .Gross1 = ds.Tables(0).Rows(0)("Gross1")
                .Gross2 = ds.Tables(0).Rows(0)("Gross2")
                .Gross3 = ds.Tables(0).Rows(0)("Gross3")
                .TotalSalaryPreEmp = ds.Tables(0).Rows(0)("TotalSalaryPreEmp").ToString()
                .Sec16ii = ds.Tables(0).Rows(0)("Sec16ii")
                .Sec16iii = ds.Tables(0).Rows(0)("Sec16iii")
                .TaxAmt = ds.Tables(0).Rows(0)("TaxAmt").ToString()
                .Surcharge = ds.Tables(0).Rows(0)("Surcharge").ToString()
                .ECess = ds.Tables(0).Rows(0)("ECess").ToString()
                .Relief89 = ds.Tables(0).Rows(0)("Relief89").ToString()
                .TDSOnPerks = ds.Tables(0).Rows(0)("TDSOnPerks").ToString
                .RetnID = ds.Tables(0).Rows(0)("RetnId")
            End With
            Fetch = DM
        Else

            Fetch = Nothing
        End If
        ds.Dispose()
        DM = Nothing
    End Function

    Public Property PR24Name As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            PR24Name = mvarPR24Name
        End Get
        Set(ByVal vData As String)
            mvarPR24Name = vData
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

    Public Property Gross3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Gross3 = mvarGross3
        End Get
        Set(ByVal vData As String)
            mvarGross3 = vData
        End Set
    End Property

    Public Property Gross2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Gross2 = mvarGross2
        End Get
        Set(ByVal vData As String)
            mvarGross2 = vData
        End Set
    End Property

    Public Property Gross1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Gross1 = mvarGross1
        End Get
        Set(ByVal vData As String)
            mvarGross1 = vData
        End Set
    End Property

    Public Property EmpToDt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            EmpToDt = mvarEmpToDt
        End Get
        Set(ByVal vData As String)
            mvarEmpToDt = vData
        End Set
    End Property

    Public Property EmpFromDt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            EmpFromDt = mvarEmpFromDt
        End Get
        Set(ByVal vData As String)
            mvarEmpFromDt = vData
        End Set
    End Property

    Public Property DDesgn As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            DDesgn = mvarDDesgn
        End Get
        Set(ByVal vData As String)
            mvarDDesgn = vData
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

    Public Property F16ID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            F16ID = mvarF16ID
        End Get
        Set(ByVal vData As String)
            mvarF16ID = vData
        End Set
    End Property

    Public Property DateOfForm() As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            DateOfForm = mvarDateOfForm
        End Get
        Set(ByVal vNewValue As Date)
            mvarDateOfForm = vNewValue
        End Set
    End Property

    Public Property PlaceOfForm() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            PlaceOfForm = mvarPlaceOfForm
        End Get
        Set(ByVal vNewValue As String)
            mvarPlaceOfForm = vNewValue
        End Set
    End Property

    Public Property TDSOnPerks() As VariantType
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            TDSOnPerks = mvarTDSOnPerks
        End Get
        Set(ByVal vNewValue As VariantType)
            mvarTDSOnPerks = vNewValue
        End Set
    End Property
End Class