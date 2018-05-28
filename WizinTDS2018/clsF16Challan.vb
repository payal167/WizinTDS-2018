Public Class clsF16Challan
    Private mvarChallanId As Long
    Private mvarF16ID As Long
    Private mvarTaxAmt As Double
    Private mvarSurcharge As Double
    Private mvarECess As Double
    Private mvarChqDDNo As Long
    Private mvarBankChallanNo As Long
    Private mvarBankBrCode As Long
    Private mvarDtOfChallan As Date
    Private mvarRemark As String

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


    Public Property ChallanID() As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Id
            ChallanID = mvarChallanId
        End Get
        Set(ByVal vData As Long)
            mvarChallanId = vData
        End Set
    End Property

    Public Property F16ID() As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Id
            F16ID = mvarF16ID
        End Get
        Set(ByVal vData As Long)
            mvarF16ID = vData
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

    Public Property ChqDDNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            ChqDDNo = mvarChqDDNo
        End Get
        Set(ByVal vData As String)
            mvarChqDDNo = vData
        End Set
    End Property

    Public Property BankChallanNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            BankChallanNo = mvarBankChallanNo
        End Get
        Set(ByVal vData As String)
            mvarBankChallanNo = vData
        End Set
    End Property

    Public Property BankBrCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            BankBrCode = mvarBankBrCode
        End Get
        Set(ByVal vData As String)
            mvarBankBrCode = vData
        End Set
    End Property

    Public Property DtOfChallan As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            DtOfChallan = mvarDtOfChallan
        End Get
        Set(ByVal vData As String)
            mvarDtOfChallan = vData
        End Set
    End Property

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

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(ChallanID) as Id From F16Challan"
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            MaxID = nds.Tables(0).Rows(0)(0)
        Else
            MaxID = 0
        End If
        nds.Dispose()
    End Function

    Public Function Insert(mF16ID As Long, mTaxAmt As Double, mSurcharge As Double, mECess As Double, mChqDDNo As Long, mBankChallanNo As Long, mBankBrCode As Long, mDtOfChallan As Date, mRemark As String) As Boolean
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
        Dim mChallanID = MaxID() + 1

        sql = "Insert into F16Challan(ChallanID, F16ID, TaxAmt, Surcharge, ECess, ChqDDNo, BankChallanNo, BankBrCode, DtOfChallan, Remark) values( " _
    & mChallanID & "," & mF16ID & "," & mTaxAmt & "," & mSurcharge & "," _
    & mECess & "," & mChqDDNo & "," & mBankChallanNo & "," & mBankBrCode & "," _
    & "#" & Format(mDtOfChallan, "dd/mmm/yyyy") & "#," _
    & IIf(mRemark = vbNullString, "Null", "'" & mRemark & "'") & ")"
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
        sql = "Delete * From F16Challan Where F16ID = " & ID
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
