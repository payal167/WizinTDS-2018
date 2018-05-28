Public Class ClsBankMstObj
    'local variable(s) to hold property value(s)
    Private mvarBankBrCode As String 'local copy
    Private mvarBType As String 'local copy
    Private mvarBankName As String 'local copy
    Private mvarBranch As String 'local copy
    Private mvarAddress As String 'local copy
    Private mvarCity As String 'local copy
    Private mvarState As String 'local copy
    Private mvarRegion As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarCoID As Long 'local copy
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent AfterDelete[(arg1, arg2, ... , argn)]
    Public Event AfterDelete()
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent AfterSave[(arg1, arg2, ... , argn)]
    Public Event AfterSave()
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent BeforeDelete[(arg1, arg2, ... , argn)]
    Public Event BeforeDelete(ByRef Cancel As Boolean)
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent PrepareDataForSave[(arg1, arg2, ... , argn)]
    Public Event PrepareDataForSave(ByRef Cancel As Boolean)
    'To fire this event, use RaiseEvent with the following syntax:
    'RaiseEvent BeforeSave[(arg1, arg2, ... , argn)]
    Public Event BeforeSave(ByRef Cancel As Boolean)





    Public Property coid As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CoID
            coid = mvarCoID
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.CoID = 5
            mvarCoID = vData
        End Set
    End Property

    Public Property Region As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Region
            Region = mvarRegion
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Region = 5
            mvarRegion = vData
        End Set
    End Property
    Public Property State As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.State
            State = mvarState
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.State = 5
            mvarState = vData
        End Set
    End Property

    Public Property City As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.City
            City = mvarCity
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.City = 5
            mvarCity = vData
        End Set
    End Property

    Public Property Address As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Address
            Address = mvarAddress
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Address = 5
            mvarAddress = vData
        End Set
    End Property

    Public Property Branch As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Branch
            Branch = mvarBranch
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Branch = 5
            mvarBranch = vData
        End Set
    End Property
    Public Property BankName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BankName
            BankName = mvarBankName
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.BankName = 5
            mvarBankName = vData
        End Set
    End Property

    Public Property BType As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BType
            BType = mvarBType
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.BType = 5
            mvarBType = vData
        End Set
    End Property

    Public Property BankBrCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.BankBrCode
            BankBrCode = mvarBankBrCode
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.BankBrCode = 5
            mvarBankBrCode = vData
        End Set
    End Property
    Public Function Fetch(ByVal BId As String) As ClsBankMstObj
        Dim nds As New DataSet, BNK As New ClsBankMstObj
        nds = FetchDataSet("SELECT * FROM BankMst WHERE BankBrCode='" & BId & "'")
        If nds.Tables(0).Rows.Count > 0 Then
            With BNK

                .BankName = nds.Tables(0).Rows(0)("BankName") & ""
                .Branch = nds.Tables(0).Rows(0)("Branch") & ""
                .City = nds.Tables(0).Rows(0)("City") & ""
                .State = nds.Tables(0).Rows(0)("State") & ""
                .BankBrCode = Format(nds.Tables(0).Rows(0)("BankBrCode"), "0000000")
                .BType = nds.Tables(0).Rows(0)("BType") & ""
                .Address = nds.Tables(0).Rows(0)("Address") & ""
                .Region = nds.Tables(0).Rows(0)("Region") & ""
            End With
            Fetch = BNK
        Else
            Fetch = Nothing
        End If
        nds.Dispose()
        BNK = Nothing

    End Function

    Public Function Delete(ByVal BCode As String) As Boolean
        Dim cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand

        cnl = False
        RaiseEvent BeforeDelete(cnl)
        If cnl = True Then
            Delete = False
            Exit Function
        End If
        Dim sql As String


        sql = "Delete * From BankMst Where BankBrCode = '" & BCode & "'"

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

    Public Function Update(ByVal BankMst As ClsBankMstObj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand

        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then Exit Function
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then Exit Function
        With BankMst
            sql = "Update BankMst Set BType = " & IIf(.BType = vbNullString, "Null", "'" & .BType & "'") & "," _
                & " BankName = " & IIf(.BankName = vbNullString, "Null", "'" & .BankName & "'") & "," _
                & " Branch = " & IIf(.Branch = vbNullString, "Null", "'" & .Branch & "'") & "," _
                & " Address = " & IIf(.Address = vbNullString, "Null", "'" & .Address & "'") & "," _
                & " City = " & IIf(.City = vbNullString, "Null", "'" & .City & "'") & "," _
                & " State = " & IIf(.State = vbNullString, "Null", "'" & .State & "'") & "," _
                & " Region = " & IIf(.Region = vbNullString, "Null", "'" & .Region & "'") & "," _
                & " Coid = " & .coid _
                & " Where BankBrCode = '" & .BankBrCode & "'"
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
        End Try

        cmd.Dispose()
        transaction.Dispose()




    End Function

    Public Function Insert(ByVal Bank As ClsBankMstObj) As Boolean
        Dim sql As String, cnl As Boolean
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction
        Dim err As ErrObject
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then Exit Function
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then Exit Function

        With Bank
            '.BankBrCode = MaxID + 1
            sql = "Insert Into BankMst (BankBrCode,BType,BankName,Branch,Address,City,State,Region,Coid) Values ('" _
                & String.Format("{0:0000000}", Val(.BankBrCode)) & "'," _
                & IIf(.BType = vbNullString, "Null", "'" & .BType & "'") & "," _
                & IIf(.BankName = vbNullString, "Null", "'" & .BankName & "'") & "," _
                & IIf(.Branch = vbNullString, "Null", "'" & .Branch & "'") & "," _
                & IIf(.Address = vbNullString, "Null", "'" & .Address & "'") & "," _
                & IIf(.City = vbNullString, "Null", "'" & .City & "'") & "," _
                & IIf(.State = vbNullString, "Null", "'" & .State & "'") & "," _
                & IIf(.Region = vbNullString, "Null", "'" & .Region & "'") & "," & .coid & ")"
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
        If err.Number = "-2147467259" Then
            MessageBox.Show("This BSR Code alredy exist for the Company ", "Duplicate BSR Code Err")
        Else
            MessageBox.Show(err.Description)
        End If
    End Function

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(BankBrCode) as BankBrCode From BankMst"
        nds = FetchDataSet(sql)

        If Not (nds.Tables(0).Rows.Count > 0) Then
            MaxID = nds.Tables(0).Rows(0)("BankBrCode").Value
        Else
            MaxID = 0
        End If
        nds.Dispose()

    End Function

    Public Function LinkC26Q(ByVal BankBrCode As String) As Boolean
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select BankBrCode  From Challan26Q Where BankBrCode = " & Val(BankBrCode)
        nds = FetchDataSet(sql)

        If nds.Tables(0).Rows.Count > 0 Then
            LinkC26Q = True
        Else
            LinkC26Q = False
        End If
        nds.Dispose()
    End Function
    Public Function LinkC24Q(ByVal BankBrCode As String) As Boolean
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select BankBrCode  From Challan24Q Where BankBrCode = " & Val(BankBrCode)
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            LinkC24Q = True
        Else
            LinkC24Q = False
        End If
        nds.Dispose()
    End Function
    Public Sub Fill_cmbState(frm As frmBankMst)

        Dim nds As New DataSet
        Dim QueSt As String = "Select StateID,StateName from StateMst"
        nds = FetchDataSet(QueSt)

        If nds.Tables(0).Rows.Count > 0 Then
            frm.cboDState.DataSource = nds.Tables(0)
            frm.cboDState.ValueMember = "StateID"
            frm.cboDState.DisplayMember = "StateName"
        End If
        frm.cboDState.SelectedIndex = -1
        nds.Dispose()

    End Sub
    Public Sub Fill_CmbBrnchCode(frm As frmBankMst)

        Dim nds As New DataSet
        Dim Sql As String = "SELECT BankMst.BankBrCode FROM BankMst"
        frm.cbobanbrcode.DataSource = Nothing
        nds = FetchDataSet(Sql)
        If nds.Tables(0).Rows.Count > 0 Then
            frm.cbobanbrcode.DataSource = nds.Tables(0)
            frm.cbobanbrcode.ValueMember = "BankBrCode"
            frm.cbobanbrcode.DisplayMember = "BankBrCode"
        End If
        frm.cbobanbrcode.SelectedIndex = -1
        nds.Dispose()


    End Sub


End Class
