Public Class clsForm16MoreDetails
    '    'local variable(s) to hold property value(s)
    Private mvarF16ID As Long 'local copy
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
    Private mvarTypeOfDetail As String 'local copy
    Private mvarParticulars As String 'local copy
    Private mvarGrossAmt As Double 'local copy
    Private mvarQualifyAmt As Double 'local copy
    Private mvarDeductibleAmt As Double 'local copy
    'local variable(s) to hold property value(s)
    Private mvarId As Long 'local copy
    Public Property ID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            ID = mvarId
        End Get
        Set(ByVal vData As String)
            mvarId = vData
        End Set
    End Property


    Public Property DeductibleAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            DeductibleAmt = mvarDeductibleAmt
        End Get
        Set(ByVal vData As String)
            mvarDeductibleAmt = vData
        End Set
    End Property


    Public Property QualifyAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            QualifyAmt = mvarQualifyAmt
        End Get
        Set(ByVal vData As String)
            mvarQualifyAmt = vData
        End Set
    End Property

    Public Property GrossAmt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            GrossAmt = mvarGrossAmt
        End Get
        Set(ByVal vData As String)
            mvarGrossAmt = vData
        End Set
    End Property

    Public Property Particulars As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            Particulars = mvarParticulars
        End Get
        Set(ByVal vData As String)
            mvarParticulars = vData
        End Set
    End Property


    Public Property TypeOfDetail As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Category
            TypeOfDetail = mvarTypeOfDetail
        End Get
        Set(ByVal vData As String)
            mvarTypeOfDetail = vData
        End Set
    End Property



    '    Public Sub Update(F16Details As clsForm16Details)
    '    End Sub

    '    Public Function MaxID() As Long
    '        Dim rs As New ADODB.Recordset
    '        Dim sql As String
    '        sql = "Select Max(id) as Id From Form16MoreDetails"
    '        rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    '    If Not IsNull(rs!ID) Then
    '            MaxID = rs!ID
    '        Else
    '            MaxID = 0
    '        End If
    '        If rs.State = adStateOpen Then rs.Close
    'Set rs = Nothing

    'End Function

    '    Public Sub Insert(F16Details As clsForm16Details)
    '    End Sub

    '    Public Function Delete(ID As Long) As Boolean
    '        Dim cnl As Boolean
    '        On Error GoTo DelErr
    '        cnl = False : RaiseEvent BeforeDelete(cnl)
    '        If cnl = True Then Exit Function
    '        Dim sql As String

    '        sql = "Delete * From Form16MoreDetails Where F16ID = " & ID
    '        Cnn.Execute sql
    '    If Cnn.Errors.Count = 0 Then
    '            Delete = True
    '            RaiseEvent AfterDelete()
    '        Else
    '            Delete = False
    '        End If
    '        Exit Function
    'DelErr:
    '        MsgBox Err.Description, , Err.Number
    'End Function

    Public Sub Fetch(ByVal ID As Long)
    End Sub


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
End Class

