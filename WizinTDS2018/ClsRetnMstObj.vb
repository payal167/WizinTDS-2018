Public Class ClsRetnMstObj
    'local variable(s) to hold property value(s)
    Private mvarRetnID As Long 'local copy
    Private mvarCoID As Long 'local copy
    Private mvarAYear As String 'local copy
    Private mvarFrmType As String 'local copy
    Private mvarTxtFileName As String 'local copy
    Private mvarDtOfFiling As Date 'local copy
    Private mvarPRN As String 'local copy
    Private mvarRPRN As String 'local copy
    Private mvarReceiptNo As String 'local copy


    Public Property TxtFileName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TxtFileName
            TxtFileName = mvarTxtFileName
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TxtFileName = 5
            mvarTxtFileName = vData
        End Set
    End Property

    Public Property FrmType As String
    Get 
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.FrmType
            FrmType = mvarFrmType
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.FrmType = 5
            mvarFrmType = vData
        End Set
    End Property
    Public Property AYear As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.AYear
            AYear = mvarAYear
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.AYear = 5
            mvarAYear = vData
        End Set
    End Property
    Public Property coid As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.CoId
            coid = mvarCoID
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.CoId = 5
            mvarCoID = vData
        End Set
    End Property

    Public Property RetnID As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.RetnId
            RetnID = mvarRetnID
        End Get
        Set(ByVal vData As Long)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.RetnId = 5
            mvarRetnID = vData
        End Set
    End Property

    Public Function Fetch(ByVal ID As Long) As ClsRetnMstObj
        Dim ret As New ClsRetnMstObj
        Dim nds As New DataSet

        nds = FetchDataSet("SELECT * FROM RetnMst WHERE RetnId=" & ID)
        If nds.Tables(0).Rows.Count > 0 Then
            With ret
                .RetnID = nds.Tables(0).Rows(0)("RetnID") & ""
                .coid = nds.Tables(0).Rows(0)("coid") & ""
                .AYear = nds.Tables(0).Rows(0)("AYear") & ""
                .FrmType = nds.Tables(0).Rows(0)("FrmType") & ""
                .TxtFileName = nds.Tables(0).Rows(0)("TxtFileName") & ""
                .prn = nds.Tables(0).Rows(0)("prn")

            End With
            Fetch = ret
        Else
            Fetch = Nothing
        End If
        nds.Dispose()
        ret = Nothing

    End Function

    Public Function Delete(ByVal ID As Long) As Boolean
        Dim cnl As Boolean

        If cnl = True Then
            Delete = False
            Exit Function
        End If
        Dim sql As String
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand

        sql = "Delete * From RetnMst Where RetnID = " & ID
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

        Exit Function

    End Function

    Public Function Update(ByVal RetnMst As ClsRetnMstObj) As Boolean
        Dim sql As String
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand

        With RetnMst
            sql = "Update RetnMst Set CoId = " & IIf(.coid = 0, 0, .coid) & "," _
                & " AYear = " & IIf(.AYear = vbNullString, "Null", "'" & .AYear & "'") & "," _
                & " FrmType = " & IIf(.FrmType = vbNullString, "Null", "'" & .FrmType & "'") & "," _
                & " TxtFileName = " & IIf(.TxtFileName = vbNullString, "Null", "'" & .TxtFileName & "'") & "" _
                & " Where RetnId = " & .RetnID
        End With

        cmd.Connection = cn
            cmd.CommandText = sql
            transaction = cn.BeginTransaction()
            cmd.Transaction = transaction

        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
            Update = True


        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage
            Update = False

        End Try


        cmd.Dispose()
            transaction.Dispose()
    End Function


    Public Function Insert(ByVal RetnMst As ClsRetnMstObj) As Boolean
        Dim sql As String
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction

        With RetnMst
            sql = "Insert Into RetnMst (RetnId,Coid,AYear,FrmType,TxtFileName) Values (" _
                & MaxID() + 1 & "," & IIf(.coid = 0, 0, .coid) & "," _
                & IIf(.AYear = vbNullString, "Null", "'" & .AYear & "'") & "," _
                & IIf(.FrmType = vbNullString, "Null", "'" & .FrmType & "'") & "," _
                & IIf(.TxtFileName = vbNullString, "Null", "'" & .TxtFileName & "'") & ")"
            cmd.CommandText = sql
            cmd.Connection = cn
            transaction = cn.BeginTransaction()
            cmd.Transaction = transaction

            Try
                cmd.ExecuteNonQuery()
                transaction.Commit()

                Insert = True
            Catch ex As Exception

                transaction.Rollback()
                MessageBox.Show(ex.Message)
                Insert = False
            End Try

            cmd.Dispose()
            transaction.Dispose()

        End With


    End Function

    Public Function MaxID() As Long
        Dim nds As New DataSet
        Dim sql As String
        sql = "Select Max(RetnId) as ID From RetnMst"
        nds = FetchDataSet(sql)
        If nds.Tables(0).Rows.Count > 0 Then
            If nds.Tables(0).Rows(0)("ID").ToString() = "" Then
                MaxID = 0
            Else
                MaxID = nds.Tables(0).Rows(0)("ID")
            End If
        Else
            MaxID = 0
        End If
        nds.Dispose()
    End Function

    Public Function UpdRetnDet(ByVal RetnMst As ClsRetnMstObj) As Boolean
        Dim sql As String
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand

        With RetnMst
            sql = "Update RetnMst Set dtoffiling=" & IIf(.DtOfFiling = vbNullString, "Null", "#" & Format(.DtOfFiling, "dd/MMM/yyyy") & "#") &
                 ",PRN=" & IIf(Val(.prn) = 0, "Null", "'" & .prn & "'") &
                 ",RPRN=" & IIf(Val(.RPRN) = 0, "Null", "'" & .RPRN & "'") &
                 ",NewReceiptNo=" & IIf(Len(.ReceiptNo) = 0, "Null", "'" & .ReceiptNo & "'") &
                 " Where RetnId = " & .RetnID
        End With

        cmd.Connection = cn
        cmd.CommandText = sql
        transaction = cn.BeginTransaction()
        cmd.Transaction = transaction

        Try
            cmd.ExecuteNonQuery()
            transaction.Commit()
            UpdRetnDet = True


        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show(ex.Message) 'Error MEssage
            UpdRetnDet = False

        End Try


        cmd.Dispose()
        transaction.Dispose()
    End Function


    Public Property DtOfFiling As Date
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DtOfFiling
            DtOfFiling = mvarDtOfFiling
        End Get
        Set(ByVal vData As Date)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DtOfFiling = 01/01/07
            mvarDtOfFiling = vData
        End Set
    End Property

    Public Property prn As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.prn
            prn = mvarPRN
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.prn = 4
            mvarPRN = vData
        End Set
    End Property

    Public Property RPRN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Rprn
            RPRN = mvarRPRN
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Rprn = 4
            mvarRPRN = vData
        End Set
    End Property
    Public Property ReceiptNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Rprn
            ReceiptNo = mvarReceiptNo
        End Get
        Set(ByVal vData As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.Rprn = 4
            mvarReceiptNo = vData
        End Set
    End Property
End Class
