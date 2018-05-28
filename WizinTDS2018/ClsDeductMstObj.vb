Imports System.Data.OleDb

Public Class ClsDeductMstObj
    'local variable(s) to hold property value(s)
    Private mvarDId As Long 'local copy
    Private mvarCoID As Long 'local copy
    Private mvarDName As String 'local copy
    Private mvarDPan As String 'local copy
    Private mvarDAdd1 As String 'local copy
    Private mvarDAdd2 As String 'local copy
    Private mvarDAdd3 As String 'local copy
    Private mvarDAdd4 As String 'local copy
    Private mvarDAdd5 As String 'local copy
    Private mvarDState As String 'local copy
    Private mvarDPin As String 'local copy
    Private mvarDType As String 'local copy
    Private mvarDRef As String 'local copy
    Private mvarDCat As String  'local copy
    Private mvarDeEmail As String
    Private mvarDeTin As String
    Private mvarDePhone As String
    Public Event PrepareDataForSave(ByVal Cancel As Boolean)
    Public Event BeforeSave(ByVal Cancel As Boolean)
    Public Event AfterSave()
    Public Event BeforeDelete(ByVal Cancel As Boolean)
    Public Event AfterDelete()
    'local variable(s) to hold property value(s)
    Private mvarCategory As String 'local copy
    Private mvarDDesgn As String 'local copy
    Private mvarDCollNonRes As String 'Boolean
    Private mvarDPerEstInd As String 'Boolean
    Private mvarDStatenm As String

    'Public Property Let DDesgn(ByVal vData As String) 'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Category = 5
    '    mvarDDesgn = vData
    'End Property

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


    'Public Property Let Category(ByVal vData As String)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.Category = 5
    '    mvarCategory = vData
    'End Property


    Public Property Category As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            Category = mvarCategory
        End Get
        Set(ByVal vData As String)
            mvarCategory = vData
        End Set
    End Property




    'Public Property Let DType(ByVal vData As String)   'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DPin = 5
    '    mvarDType = vData
    'End Property


    Public Property DType As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            DType = mvarDType
        End Get
        Set(ByVal vData As String)
            mvarDType = vData
        End Set
    End Property


    'Public Property Let DPin(ByVal vData As String)   'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DPin = 5
    '    mvarDPin = vData
    'End Property


    Public Property DPin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            DPin = mvarDPin
        End Get
        Set(ByVal vData As String)
            mvarDPin = vData
        End Set
    End Property



    'Public Property Let DState(ByVal vData As String)  'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DState = 5
    '    mvarDState = vData
    'End Property


    Public Property DState As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            DState = mvarDState
        End Get
        Set(ByVal vData As Long)
            mvarDState = vData
        End Set
    End Property



    'Public Property Let DAdd5(ByVal vData As String)    'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DAdd5 = 5
    '    mvarDAdd5 = vData
    ' End Property


    Public Property DAdd5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd5
            DAdd5 = mvarDAdd5
        End Get
        Set(ByVal vData As String)
            mvarDAdd5 = vData
        End Set
    End Property



    'Public Property Let DAdd4(ByVal vData As String)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DAdd4 = 5
    '    mvarDAdd4 = vData
    'End Property


    Public Property DAdd4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd4
            DAdd4 = mvarDAdd4
        End Get
        Set(ByVal vData As String)
            mvarDAdd4 = vData
        End Set
    End Property



    'Public Property Let DAdd3(ByVal vData As String)   'payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DAdd3 = 5
    '    mvarDAdd3 = vData
    'End Property


    Public Property DAdd3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd3
            DAdd3 = mvarDAdd3
        End Get
        Set(ByVal vData As String)
            mvarDAdd3 = vData
        End Set
    End Property



    'Public Property Let DAdd2(ByVal vData As String)   'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DAdd2 = 5
    '    mvarDAdd2 = vData
    'End Property


    Public Property DAdd2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd2
            DAdd2 = mvarDAdd2
        End Get
        Set(ByVal vData As String)
            mvarDAdd2 = vData
        End Set
    End Property



    'Public Property Let DAdd1(ByVal vData As String)    'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DAdd1 = 5
    '    mvarDAdd1 = vData
    'End Property


    Public Property DAdd1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DAdd1
            DAdd1 = mvarDAdd1
        End Get
        Set(ByVal vData As String)
            mvarDAdd1 = vData
        End Set
    End Property



    'Public Property Let DPan(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DPan = 5
    '    mvarDPan = vData
    'End Property


    Public Property DPan As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPan
            DPan = mvarDPan
        End Get
        Set(ByVal vData As String)
            mvarDPan = vData
        End Set
    End Property



    'Public Property Let DName(ByVal vData As String)    'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DName = 5
    '    mvarDName = vData
    'End Property


    Public Property DName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DName
            DName = mvarDName
        End Get
        Set(ByVal vData As String)
            mvarDName = vData
        End Set
    End Property


    'Public Property Let coid(ByVal vData As Long)     'Payal
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DId = 5
    '    mvarCoID = vData
    'End Property


    Public Property coid As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DId
            coid = mvarCoID
        End Get
        Set(ByVal vData As Long)
            mvarCoID = vData
        End Set
    End Property



    'Public Property Let did(ByVal vData As Long)     'payal
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
    'Public Property Let Dref(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DState = 5
    '    mvarDRef = vData
    'End Property


    Public Property Dref As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            Dref = mvarDRef
        End Get
        Set(ByVal vData As String)
            mvarDRef = vData
        End Set
    End Property
    'Public Property Let Dcat(ByVal vData As String)
    '    'used when assigning a value to the property, on the left side of an assignment.
    '    'Syntax: X.DState = 5
    '    mvarDCat = vData
    'End Property

    Public Property Dcat As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            Dcat = mvarDCat
        End Get
        Set(ByVal vData As String)
            mvarDCat = vData
        End Set
    End Property
    Public Property DeEmail As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            DeEmail = mvarDeEmail
        End Get
        Set(ByVal vData As String)
            mvarDeEmail = vData
        End Set
    End Property
    Public Property DePhone As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            DePhone = mvarDePhone
        End Get
        Set(ByVal vData As String)
            mvarDePhone = vData
        End Set
    End Property
    Public Property DeTin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            DeTin = mvarDeTin 
        End Get
        Set(ByVal vData As String)
            mvarDeTin = vData
        End Set
    End Property

    Public Property CollNonRes As String
        'used when assigning a value to the property, on the left side of an assignment.
        'Syntax: X.IsCoAddChg = 5
        Get
            CollNonRes = mvarDCollNonRes
        End Get
        Set(ByVal vData As String)
            mvarDCollNonRes = vData
        End Set
    End Property
    Public Property PerEstInd As String
        'used when assigning a value to the property, on the left side of an assignment.
        'Syntax: X.IsCoAddChg = 5
        Get
            PerEstInd = mvarDPerEstInd
        End Get
        Set(ByVal vData As String)
            mvarDPerEstInd = vData
        End Set
    End Property
    Public Property DStatenm As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DState
            DStatenm = mvarDStatenm
        End Get
        Set(ByVal vData As String)
            mvarDStatenm = vData
        End Set
    End Property

    Public Function Fetch(ByVal dname As String) As ClsDeductMstObj
        Dim DM As New ClsDeductMstObj

        Dim nds As New DataSet
        nds = FetchDataSet("SELECT Deductmst.*, StateMst.StateName FROM Deductmst INNER JOIN StateMst ON Deductmst.DState = StateMst.StateID WHERE dname like'" & dname & "'")
        If nds.Tables(0).Rows.Count > 0 Then
            With DM
                .did = nds.Tables(0).Rows(0)("did") 'rst!did & ""
                .coid = nds.Tables(0).Rows(0)("coid") 'rst!coid & ""
                .DName = nds.Tables(0).Rows(0)("DName") 'rst!DName & ""
                .DPan = nds.Tables(0).Rows(0)("DPan").ToString() 'rst!DPan & ""
                .DAdd1 = nds.Tables(0).Rows(0)("DAdd1").ToString() 'rst!DAdd1 & ""
                .DAdd2 = nds.Tables(0).Rows(0)("DAdd2").ToString() 'rst!DAdd2 & ""
                .DAdd3 = nds.Tables(0).Rows(0)("DAdd3").ToString() 'rst!DAdd3 & ""
                .DAdd4 = nds.Tables(0).Rows(0)("DAdd4").ToString() 'rst!DAdd4 & ""
                .DAdd5 = nds.Tables(0).Rows(0)("DAdd5").ToString() 'rst!DAdd5 & ""
                .DState = nds.Tables(0).Rows(0)("DState") 'rst!DState & ""
                .DStatenm = nds.Tables(0).Rows(0)("Statename").ToString()
                .DPin = nds.Tables(0).Rows(0)("DPin") 'rst!DPin & ""
                .DType = nds.Tables(0).Rows(0)("DType").ToString() 'rst!DType & ""
                .Dref = nds.Tables(0).Rows(0)("Dpanref").ToString() 'rst!Dpanref & ""
                .Dcat = IIf(nds.Tables(0).Rows(0)("DPANCat").ToString() = "", 0, nds.Tables(0).Rows(0)("DPANCat")) 'rst!DPANCat & ""
                .Category = nds.Tables(0).Rows(0)("Category").ToString() 'rst!Category & ""
                .DDesgn = nds.Tables(0).Rows(0)("DDesgn").ToString() 'rst!DDesgn & ""
                .DeEmail = nds.Tables(0).Rows(0)("DeEmail").ToString()  ' If(nds.Tables(0).Rows(0)("DeEmail").ToString() = "", vbNullString, nds.Tables(0).Rows(0)("DeEmail").ToString()) 'IIf(IsNull(rst!DeEmail), "", rst!DeEmail) & ""
                .DePhone = nds.Tables(0).Rows(0)("DePhone").ToString() 'IIf(IsNull(rst!DePhone), "", rst!DePhone) & ""
                .DeTin = nds.Tables(0).Rows(0)("DeTin").ToString()  'IIf(IsNull(rst!DeTin), "", rst!DeTin) & ""
                If DM.Dcat >= 1 Then
                    .CollNonRes = IIf((nds.Tables(0).Rows(0)("CollNonRes")) = True, "Yes", "No")
                    .PerEstInd = IIf((nds.Tables(0).Rows(0)("PerEstInd")) = True, "Yes", "No")
                Else
                    .CollNonRes = ""
                    .PerEstInd = ""
                End If
            End With
            Fetch = DM
        Else

            Fetch = Nothing
        End If
        'If rst.State = adStateOpen Then rst.Close
        'If rst.State = ADODB.ObjectStateEnum.adStateOpen Then rst.Close()
        'rst = Nothing
        DM = Nothing
    End Function

    Public Function Delete(ByVal did As String) As Boolean
        Dim cnl As Boolean
        Dim cmd As New OleDbCommand
        On Error GoTo DelErr
        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then Exit Function
        Dim sql As String, norow As Integer

        sql = "Delete * From DeductMst Where dname = '" & did & "'"
        cmd.CommandText = sql
        cmd.Connection = cn
        norow = cmd.ExecuteNonQuery()
        If norow > 0 Then
            Delete = True
            RaiseEvent AfterDelete()

        Else
            Delete = False
        End If
        'If cn..Count = 0 Then
        '    'Cnn.CommitTrans
        '    Delete = True
        '    RaiseEvent AfterDelete()
        'Else
        '    'Cnn.RollbackTrans
        '    Delete = False
        'End If
        Exit Function
DelErr:
        MsgBox(Err.Description, , Err.Number)
    End Function

    Public Function Update(ByVal DeductMst As ClsDeductMstObj) As Boolean
        Dim cmd As New OleDbCommand
        Dim sql As String, cnl As Boolean
        Dim norow As Integer
        On Error GoTo UpErr
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then Exit Function
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then Exit Function
        With DeductMst
            sql = "Update DeductMst Set DName = " & IIf(.DName = vbNullString, "Null", "'" & DName & "'") & "," _
                & " DPan = " & IIf(.DPan = vbNullString, "Null", "'" & .DPan & "'") & "," _
                & " DAdd1 = " & IIf(.DAdd1 = vbNullString, "Null", "'" & .DAdd1 & "'") & "," _
                & " DAdd2 = " & IIf(.DAdd2 = vbNullString, "Null", "'" & .DAdd2 & "'") & "," _
                & " DAdd3 = " & IIf(.DAdd3 = vbNullString, "Null", "'" & .DAdd3 & "'") & "," _
                & " DAdd4 = " & IIf(.DAdd4 = vbNullString, "Null", "'" & .DAdd4 & "'") & "," _
                & " DAdd5 = " & IIf(.DAdd5 = vbNullString, "Null", "'" & .DAdd5 & "'") & "," _
                & " DState = " & IIf(.DState = vbNullString, "Null", Chr(34) & .DState & Chr(34)) & "," _
                & " DPin = " & IIf(.DPin = vbNullString, "Null", Chr(34) & .DPin & Chr(34)) & "," _
                & " DType = " & IIf(.DType = vbNullString, "Null", Chr(34) & .DType & Chr(34)) & "," _
                & " CoId = " & IIf(.coid = 0, 0, .coid) & "," _
                & " DPANRef = " & IIf(.Dref = vbNullString, "Null", Chr(34) & .Dref & Chr(34)) & "," _
                & " DPANCat = " & IIf(.Dcat = 0, 0, .Dcat) & "," _
                & " Category = " & IIf(.Category = vbNullString, "Null", Chr(34) & .Category & Chr(34)) & "," _
                & " Ddesgn = " & IIf(.DDesgn = vbNullString, "Null", Chr(34) & .DDesgn & Chr(34)) & "," _
                & " DeEmail = " & IIf(.DeEmail = vbNullString, "Null", Chr(34) & .DeEmail & Chr(34)) & "," _
                & " DePhone = " & IIf(.DePhone = vbNullString, "Null", Chr(34) & .DePhone & Chr(34)) & "," _
                & " DeTin = " & IIf(.DeTin = vbNullString, "Null", Chr(34) & .DeTin & Chr(34)) & "," _
                & " CollNonRes= " & IIf(.CollNonRes = True, False, True) & "," _
                & " PerEstInd =" & IIf(.PerEstInd = True, False, True) & "" _
                & " Where DId = " & .did
            'IIf(oCoMst.IsCoAddChg = True, "Y", "N")
            'Cnn.BeginTrans
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
UpErr:
        MsgBox( Err.Description, , Err.Number)

    End Function


    Public Function Insert(ByVal Deductee As ClsDeductMstObj) As Boolean
        Dim sql As String, cnl As Boolean
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

            .did = MaxID() + 1
            sql = "Insert Into DeductMSt (DId,CoId,DName,DAdd1,DAdd2,DAdd3,DAdd4,DAdd5,DState,DPAN,DPin,DType,DPANRef,DPANCat,Category, Ddesgn,DeEmail,DePhone,DeTin,CollNonRes,PerEstInd) Values (" _
        & MaxID() + 1 & "," & IIf(.coid = 0, 0, .coid) & "," _
        & IIf(.DName = vbNullString, "Null", Chr(34) & .DName & Chr(34)) & "," _
        & IIf(.DAdd1 = vbNullString, "Null", Chr(34) & .DAdd1 & Chr(34)) & "," _
        & IIf(.DAdd2 = vbNullString, "Null", Chr(34) & .DAdd2 & Chr(34)) & "," _
        & IIf(.DAdd3 = vbNullString, "Null", Chr(34) & .DAdd3 & Chr(34)) & "," _
        & IIf(.DAdd4 = vbNullString, "Null", Chr(34) & .DAdd4 & Chr(34)) & "," _
        & IIf(.DAdd5 = vbNullString, "Null", Chr(34) & .DAdd5 & Chr(34)) & "," _
        & IIf(.DState = vbNullString, "Null", Chr(34) & .DState & Chr(34)) & "," _
        & IIf(.DPan = vbNullString, "Null", Chr(34) & .DPan & Chr(34)) & "," _
        & IIf(.DPin = vbNullString, "Null", Chr(34) & .DPin & Chr(34)) & "," _
        & IIf(.DType = vbNullString, "Null", Chr(34) & .DType & Chr(34)) & "," _
        & IIf(.Dref = vbNullString, "Null", Chr(34) & .Dref & Chr(34)) & "," _
        & IIf(.Dcat = vbNullString, "Null", .Dcat) & "," _
        & IIf(.Category = vbNullString, "null", Chr(34) & .Category & Chr(34)) & "," _
        & IIf(.DDesgn = vbNullString, "Null", Chr(34) & .DDesgn & Chr(34)) & "," _
        & IIf(.DeEmail = vbNullString, "Null", Chr(34) & .DeEmail & Chr(34)) & "," _
        & IIf(.DePhone = vbNullString, "Null", Chr(34) & .DePhone & Chr(34)) & "," _
        & IIf(.DeTin = vbNullString, "Null", Chr(34) & .DeTin & Chr(34)) & "," _
        & IIf(.CollNonRes = "YES", "Yes", "No") & "," _
        & IIf(.PerEstInd = "YES", "Yes", "No") & ")"

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
        '        Exit Function
        'InErr:
        '        MsgBox(Err.Description, , Err.Number)


    End Function

    Public Function MaxID() As Long

        Dim nds As New DataSet
        nds = FetchDataSet("Select Max(DId) as DId From DeductMst")

        If nds.Tables(0).Rows(0)("did").ToString() <> "" Then
            MaxID = nds.Tables(0).Rows(0)("did")
        Else
            MaxID = 0
        End If

        nds = Nothing
    End Function
    Public Function LinkD24Qannual(ByVal nm As String) As Boolean

        Dim nds As New DataSet
        nds = FetchDataSet("Select h.DId  From deductmst h,SalaryDetail24Q d Where h.did=d.did and h.dname = '" & nm & "'")

        If nds.Tables(0).Rows.Count > 0 Then
            LinkD24Qannual = True
        Else
            LinkD24Qannual = False
        End If

        nds = Nothing
    End Function


    Public Function LinkD24Q(ByVal nm As String) As Boolean
        Dim nds As New DataSet
        nds = FetchDataSet("Select h.DId  From deductmst h,Deductee24Q d  Where h.did=d.did and h.dname = '" & nm & "'")

        If nds.Tables(0).Rows.Count > 0 Then
            LinkD24Q = True
        Else
            LinkD24Q = False
        End If

        nds = Nothing
    End Function


    Public Function LinkD26Q(ByVal nm As String) As Boolean
        Dim nds As New DataSet
        nds = FetchDataSet("Select h.DId  From deductmst h, Deductee26Q d  Where h.did=d.did And h.dname = '" & nm & "'")

        If nds.Tables(0).Rows.Count > 0 Then
            LinkD26Q = True
        Else
            LinkD26Q = False
        End If

        nds = Nothing


    End Function
    Public Function LinkD27Q(ByVal nm As String) As Boolean
        Dim nds As New DataSet
        nds = FetchDataSet("Select h.DId  From  deductmst h,Deductee27Q d Where h.did=d.did and h.dname = '" & nm & "'")

        If nds.Tables(0).Rows.Count > 0 Then
            LinkD27Q = True
        Else
            LinkD27Q = False
        End If

        nds = Nothing
    End Function
    Public Function LinkD27EQ(ByVal nm As String) As Boolean
        Dim nds As New DataSet
        nds = FetchDataSet("Select h.DId  From deductmst h,Deductee27EQ d Where h.did=d.did and h.dname = '" & nm & "'")

        If nds.Tables(0).Rows.Count > 0 Then
            LinkD27EQ = True
        Else
            LinkD27EQ = False
        End If

        nds = Nothing
    End Function
    Public Function LinkP(ByVal id As String) As Boolean
        Dim nds As New DataSet
        nds = FetchDataSet("Select h.DId  From PerqSal Where DId = " & id)

        If nds.Tables(0).Rows.Count > 0 Then
            LinkP = True
        Else
            LinkP = False
        End If

        nds = Nothing
    End Function



End Class
