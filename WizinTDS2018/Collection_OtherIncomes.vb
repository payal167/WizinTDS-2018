Public Class Collection_OtherIncomes
    'local variable to hold collection
    Private mCol As Collection

    Public Function Add(ID As Long, F16ID As Long, TypeOfDetail As String, Particulars As String, GrossAmt As Double, QualifyAmt As Double, DeductibleAmt As Double, sKey As String) As clsForm16MoreDetails
        'create a new object
        Dim objNewMember As clsForm16MoreDetails
        objNewMember = New clsForm16MoreDetails


        'set the properties passed into the method
        objNewMember.ID = ID
        objNewMember.F16ID = F16ID
        objNewMember.TypeOfDetail = TypeOfDetail
        objNewMember.Particulars = Particulars
        objNewMember.GrossAmt = GrossAmt
        objNewMember.QualifyAmt = QualifyAmt
        objNewMember.DeductibleAmt = DeductibleAmt
        If Len(sKey) = 0 Then
            mCol.Add(objNewMember)
        Else
            mCol.Add(objNewMember, sKey)
        End If


        'return the object created
        Add = objNewMember
        objNewMember = Nothing


    End Function
    Public Property Item(vntIndexKey As VariantType) As clsForm16MoreDetails
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Item = mCol(vntIndexKey)
        End Get
        Set()
            'mCol(vntIndexKey) = vntIndexKey
        End Set
    End Property

    Public Property Count As Long
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            Count = mCol.Count
        End Get
        Set()
            'mCol(vntIndexKey) = vntIndexKey
        End Set
    End Property
    '    Public Property Get Item(vntIndexKey As Variant) As clsForm16MoreDetails
    'Attribute Item.VB_UserMemId = 0
    '    'used when referencing an element in the collection
    '    'vntIndexKey contains either the Index or Key to the collection,
    '    'this is why it is declared as a Variant
    '    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    '  Set Item = mCol(vntIndexKey)
    'End Property



    '    Public Property Get Count() As Long
    '    'used when retrieving the number of elements in the
    '    'collection. Syntax: Debug.Print x.Count
    '    Count = mCol.Count
    'End Property


    Public Sub Remove(vntIndexKey As VariantType)
        mCol.Remove(vntIndexKey)
    End Sub

    Public Property NewEnum() As stdole.IUnknown
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DPin
            NewEnum = mCol.GetEnumerator
        End Get
        Set()

        End Set
    End Property


    Private Sub Class_Initialize()
        'creates the collection when this class is created
        mCol = New Collection
    End Sub


    Private Sub Class_Terminate()
        'destroys collection when this class is terminated
        mCol = Nothing
    End Sub

    Public Sub Clear()
        'used to empty the collection
        Dim i As Long
        For i = Me.Count To 1 Step -1
            Me.Remove(i)
        Next i
    End Sub


End Class
