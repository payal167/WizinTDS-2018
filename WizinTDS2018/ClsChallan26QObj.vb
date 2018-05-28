Public Class ClsChallan26QObj
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

    'Public Property Let MinorHead(ByVal vData As String)  'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.MinorHead = 5
    '    mvarMinorHead = vData
    'End Property


    Public Property MinorHead As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            MinorHead = mvarMinorHead
        End Get
        Set(ByVal vData As String)
            mvarMinorHead = vData
        End Set
    End Property



    'Public Property Let AFees(ByVal vData As Double)   'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.AFees = 5
    '    mvarAFees = vData
    'End Property


    Public Property AFees As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AFees
            AFees = mvarAFees
        End Get
        Set(ByVal vData As Double)
            mvarAFees = vData
        End Set
    End Property



    'Public Property Let BankChallanNo(ByVal vData As Long)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.BankChallanNo = 5
    '    mvarBankChallanNo = vData
    'End Property


    Public Property BankChallanNo As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BankChallanNo
            BankChallanNo = mvarBankChallanNo
        End Get
        Set(ByVal vData As Long)
            mvarBankChallanNo = vData
        End Set
    End Property



    'Public Property Let Remark(ByVal vData As String)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Remark = 5
    '    mvarRemark = vData
    'End Property


    Public Property Remark() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Remark
            Remark = mvarRemark
        End Get
        Set(ByVal vData As String)
            mvarRemark = vData
        End Set
    End Property



    'Public Property Let TranVouNo(ByVal vData As Long)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.TranVouNo = 5
    '    mvarTranVouNo = vData
    'End Property


    Public Property TranVouNo As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TranVouNo
            TranVouNo = mvarTranVouNo
        End Get
        Set(ByVal vData As Long)
            mvarTranVouNo = vData
        End Set
    End Property



    'Public Property Let AOthers(ByVal vData As Double)   'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.AECess = 5
    '    mvarAECess = vData
    'End Property


    Public Property AOthers As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AECess
            AOthers = mvarAECess
        End Get
        Set(ByVal vData As Double)
            mvarAECess = vData
        End Set
    End Property



    'Public Property Let AInterest(ByVal vData As Double)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.AInterest = 5
    '    mvarAInterest = vData
    'End Property


    Public Property AInterest As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AInterest
            AInterest = mvarAInterest
        End Get
        Set(ByVal vData As Double)
            mvarAInterest = vData
        End Set
    End Property



    'Public Property Let TotalTax(ByVal vData As Double)  'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.TotalTax = 5
    '    mvarTotalTax = vData
    'End Property


    Public Property TotalTax As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TotalTax
            TotalTax = mvarTotalTax
        End Get
        Set(ByVal vData As Double)
            mvarTotalTax = vData
        End Set
    End Property




    'Public Property Let CollCode(ByVal vData As String)    'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.CollCode = 5
    '    mvarCollCode = vData
    'End Property


    Public Property CollCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CollCode
            CollCode = mvarCollCode
        End Get
        Set(ByVal vData As String)
            mvarCollCode = vData
        End Set
    End Property



    'Public Property Let IsBookEntry(ByVal vData As Boolean)   'payal
    '        'used when assigning a value to the property, on the left side of an assignment.
    '        'Syntax: X.IsBookEntry = 5
    '    mvarIsBookEntry = vData
    '    End Property


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



    'Public Property Let ChqDDNo(ByVal vData As Double)        'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.ChqNo = 5
    '    mvarChqNo = vData
    'End Property


    Public Property ChqDDNo As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.ChqNo
            ChqDDNo = mvarChqNo
        End Get
        Set(ByVal vData As Double)
            mvarChqNo = vData
        End Set
    End Property



    'Public Property Let Others(ByVal vData As Double)    'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Others = 5
    '    mvarOthers = vData
    'End Property


    Public Property Others As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Others
            Others = mvarOthers
        End Get
        Set(ByVal vData As Double)
            mvarOthers = vData
        End Set
    End Property



    'Public Property Let Interest(ByVal vData As Double)      'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Interest = 5
    '    mvarInterest = vData
    'End Property


    Public Property Interest As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Interest
            Interest = mvarInterest
        End Get
        Set(ByVal vData As Double)
            mvarInterest = vData
        End Set
    End Property



    'Public Property Let ECess(ByVal vData As Double)    'Payal
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



    'Public Property Let Surcharge(ByVal vData As Double)
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
        Set(ByVal vdata As Double)
            mvarSurcharge = vData
        End Set
    End Property




    'Public Property Let BankBrCode(ByVal vData As Long)   'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.BankBrCode = 5
    '    mvarBankBrCode = vData
    'End Property


    Public Property BankBrCode As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BankBrCode
            BankBrCode = mvarBankBrCode
        End Get
        Set(ByVal vData As Long)
            mvarBankBrCode = vData
        End Set
    End Property




    'Public Property Let DtOfChallan(ByVal vData As Date)    'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DtOfVoucher = 5
    '    mvarDtOfVoucher = vData
    'End Property


    Public Property DtOfChallan As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfVoucher
            DtOfChallan = mvarDtOfVoucher
        End Get
        Set(ByVal vData As Date)
            mvarDtOfVoucher = vData
        End Set
    End Property



    'Public Property Let TaxAmt(ByVal vData As Double)     'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Amt = 5
    '    mvarAmt = vData
    'End Property


    Public Property TaxAmt As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Amt
            TaxAmt = mvarAmt
        End Get
        Set(ByVal vData As Double)
            mvarAmt = vData
        End Set
    End Property



    'Public Property Let Sec(ByVal vData As String)    "payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Section = 5
    '    mvarSection = vData
    'End Property


    Public Property Sec As String
    get

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Section
            Sec = mvarSection
        End Get
        Set(ByVal vData As String)
            mvarSection = vData
        End Set
    End Property



    'Public Property Let RetnID(ByVal vData As Long)    'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.RetnId = 5
    '    mvarRetnID = vData
    'End Property


    Public Property RetnID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RetnId
            RetnID = mvarRetnID
        End Get
        Set(ByVal vData As Long)
            mvarRetnID = vData
        End Set
    End Property



    'Public Property Let ChallanID(ByVal vData As Long)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Id = 5
    '    mvarId = vData
    'End Property


    Public Property ChallanID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Id
            ChallanID = mvarId
        End Get
        Set(ByVal vData As Long)
            mvarId = vData
        End Set
    End Property

    'Public Function Fetch(ByVal ID As Long) As ClsChallan26QObj
    '    Dim rst As New ADODB.Recordset, Chln As New ClsChallan26QObj
    '    rst.Open("SELECT * FROM Challan26Q WHERE ChallanID=" & ID, Cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    '    If rst.RecordCount > 0 Then
    '        With Chln
    '            .ChallanID = rst("ChallanID").Value & ""
    '            .RetnID = rst("RetnID").Value & ""
    '            .Sec = rst("Sec").Value & ""
    '            .TaxAmt = rst("TaxAmt").Value & ""
    '            .BankChallanNo = rst("BankChallanNo").Value & ""
    '            .DtOfChallan = Format(rst("DtOfChallan").Value, "dd/MMM/yyyy") & ""
    '            .BankBrCode = rst("BankBrCode").Value & ""
    '            .AFees = rst("AFees").Value & ""
    '            .MinorHead = rst("MinorHead").Value & ""
    '            '.ChallanID = rst!ChallanID & ""
    '            '.RetnID = rst!RetnID & ""
    '            '.Sec = rst!Sec & ""
    '            '.TaxAmt = rst!TaxAmt & ""
    '            '.BankChallanNo = rst!BankChallanNo & ""
    '            '.DtOfChallan = Format(rst!DtOfChallan, "dd/MMM/yyyy") & ""
    '            '.BankBrCode = rst!BankBrCode & ""
    '            '.AFees = rst!AFees & ""
    '            '.MinorHead = rst!MinorHead & ""
    '        End With
    '        Fetch = Chln
    '    Else
    '        Fetch = Nothing
    '    End If
    '    If rst.State = ADODB.ObjectStateEnum.adStateOpen Then rst.Close()
    '    rst = Nothing
    '    Chln = Nothing
    'End Function

    Public Function Delete(ByVal ID As Long) As Boolean
        Dim cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
        'On Error GoTo DelErr
        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then
            Delete = False
            Exit Function
        End If

        Dim sql As String
        sql = "Delete * From Challan26Q Where ChallanID = " & ID

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

    Public Function Update(ByVal Challan As ClsChallan26QObj) As Boolean
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
            sql = "Update Challan26Q Set RetnId = " & IIf(.RetnID = 0, 0, .RetnID) & "," _
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
                & " MinorHead = " & IIf(String.IsNullOrEmpty(.MinorHead), vbNullString, "'" & .MinorHead & "'") _
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


    Public Function Insert(ByVal Challan As ClsChallan26QObj) As Boolean
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
            Insert=False
            Exit Function
        End If

        With Challan
            .ChallanID = MaxID + 1

            sql = " Insert Into Challan26Q (ChallanId,RetnId,Sec,TaxAmt, Surcharge, Ecess, Interest, Others, "
            sql = sql & "TotalTax,AInterest,AOthers,ChqDDNo, BankChallanNo,TranVouNo, DtOfChallan, BankBrCode, IsBookEntry,Remark,"
            sql = sql & "AFees, MinorHead) Values("
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
            sql = sql & IIf(String.IsNullOrEmpty(.MinorHead), vbNullString, "'" & .MinorHead & "'") & ")"
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

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(ChallanId) as ID From Challan26Q"
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

    Public Function LinkDed26Q(ByVal ID As Long) As Boolean
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select ChallanID  From Deductee26Q Where ChallanID = " & ID
        nds = FetchDataSet(sql)

        If nds.Tables(0).Rows.Count > 0 Then

            LinkDed26Q = True
        Else
            LinkDed26Q = False
        End If
        nds.Dispose()
    End Function

    Private Sub ClsChallan26QObj_PrepareDataForSave(Cancel As Boolean) Handles Me.PrepareDataForSave

    End Sub
End Class
