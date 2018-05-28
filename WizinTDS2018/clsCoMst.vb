Public Class clsCoMst

    'local variable(s) to hold property value(s)
    Private mvarCoID As Integer 'local copy
    Private mvarCoName As String 'local copy
    Private mvarCoAdd1 As String 'local copy
    Private mvarCoAdd2 As String 'local copy
    Private mvarCoAdd3 As String 'local copy
    Private mvarCoAdd4 As String 'local copy
    Private mvarCoAdd5 As String 'local copy
    Private mvarCoStateID As Integer 'local copy
    Private mvarCoPin As String 'local copy
    Private mvarCoTAN As String 'local copy
    Private mvarCoPAN As String 'local copy
    Private mvarIsCoAddChg As Boolean 'local copy
    Private mvarCoStatus As String 'local copy
    Private mvarPRName26 As String 'local copy
    Private mvarPRDesg26 As String 'local copy
    Private mvarPRName27 As String 'local copy
    Private mvarPRDesg27 As String 'local copy
    Private mvarPRName24 As String 'local copy
    Private mvarPRDesg24 As String 'local copy
    Private mvarPR24Add1 As String 'local copy
    Private mvarPR24Add2 As String 'local copy
    Private mvarPR24Add3 As String 'local copy
    Private mvarPR24Add4 As String 'local copy
    Private mvarPR24Add5 As String 'local copy
    Private mvarPR24StateID As Integer 'local copy
    Private mvarPR24Pin As String 'local copy
    Private mvarPR24Email As String 'local copy
    Private mvarPR24Std As String 'local copy
    Private mvarPR24Phone As String 'local copy
    Private mvarIsPR24AddChg As Boolean 'local copy
    Public Event PrepareDataForSave(Cancel As Boolean)
    Public Event BeforeSave(Cancel As Boolean)
    Public Event AfterSave()
    Public Event BeforeDelete(Cancel As Boolean)
    Public Event AfterDelete()
    Private mvarCoBrDiv As String 'local copy
    Private mvarCoStd As String 'local copy
    Private mvarCoPhone As String 'local copy
    Private mvarCoMobile As String 'local copy
    Private mvarCoEmail As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarPR26Add1 As String 'local copy
    Private mvarPR26Add2 As String 'local copy
    Private mvarPR26Add3 As String 'local copy
    Private mvarPR26Add4 As String 'local copy
    Private mvarPR26Add5 As String 'local copy
    Private mvarPR26StateID As Integer 'local copy
    Private mvarPR26Pin As String 'local copy
    Private mvarPR26Email As String 'local copy
    Private mvarPR26Std As String 'local copy
    Private mvarPR26Phone As String 'local copy
    Private mvarIsPR26AddChg As Boolean 'local copy
    Private mvarPR27Add1 As String 'local copy
    Private mvarPR27Add2 As String 'local copy
    Private mvarPR27Add3 As String 'local copy
    Private mvarPR27Add4 As String 'local copy
    Private mvarPR27Add5 As String 'local copy
    Private mvarPR27StateID As Integer 'local copy
    Private mvarPR27Pin As String 'local copy
    Private mvarPR27Email As String 'local copy
    Private mvarPR27Std As String 'local copy
    Private mvarPR27Phone As String 'local copy
    Private mvarIsPR27AddChg As Boolean 'local copy
    Private mvarPR27EName As String 'local copy
    Private mvarPR27EDesg As String 'local copy
    Private mvarPR27EAdd1 As String 'local copy
    Private mvarPR27EAdd2 As String 'local copy
    Private mvarPR27EAdd3 As String 'local copy
    Private mvarPR27EAdd4 As String 'local copy
    Private mvarPR27EAdd5 As String 'local copy
    Private mvarPR27EStateID As Integer 'local copy
    Private mvarPR27EPin As Integer 'local copy
    Private mvarPR27EEmail As String 'local copy
    Private mvarPR27EStd As String 'local copy
    Private mvarPR27EPhone As String 'local copy
    Private mvarIsPR27EAddChg As Boolean 'local copy
    'local variable(s) to hold property value(s)
    Private mvarUseForm16 As Boolean 'local copy
    Private mvarGovtStateID As String
    Private mvarPAOCode As String
    Private mvarDDOCode As String
    Private mvarMinistryID As String
    Private mvarMinistryName As String
    Private mvarPAORegNo As String
    Private mvarDDORegNo As String
    Private mvaruserid As String   'userid
    Private mvaruserpasswrd As String 'useRPASSWRD
    Private mvaruserTAN As String    'user TAN
    'local variable(s) to hold property value(s)
    Private mvarCoSTDAlt As String 'local copy
    Private mvarCoPhoneAlt As String 'local copy
    Private mvarCoEmailAlt As String 'local copy
    Private mvarPR24STDAlt As String 'local copy
    Private mvarPR24PhoneAlt As String 'local copy
    Private mvarPR24EmailAlt As String 'local copy
    Private mvarPR26STDAlt As String 'local copy
    Private mvarPR26PhoneAlt As String 'local copy
    Private mvarPR26EmailAlt As String 'local copy
    Private mvarPR27STDAlt As String 'local copy
    Private mvarPR27PhoneAlt As String 'local copy
    Private mvarPR27EmailAlt As String 'local copy
    Private mvarPR27ESTDAlt As String 'local copy
    Private mvarPR27EPhoneAlt As String 'local copy
    Private mvarPR27EEmailAlt As String 'local copy
    Private mvarAIN As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarTANRegNo As String 'local copy
    'local variable(s) to hold property value(s)
    Private mvarPR24PAN As String 'local copy
    Private mvarPR26PAN As String 'local copy
    Private mvarPR27PAN As String 'local copy
    Private mvarPR27EPAN As String 'local copy
    Private mvarGSTIN As String 'local copy

    Public Property gstin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            gstin = mvarGSTIN
        End Get
        Set(ByVal vData As String)
            mvarGSTIN = vData
        End Set
    End Property

    Public Property PR27EPAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EPAN = mvarPR27EPAN
        End Get
        Set(ByVal vData As String)
            mvarPR27EPAN = vData
        End Set
    End Property

    Public Property PR27PAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27PAN = mvarPR27PAN
        End Get
        Set(ByVal vData As String)
            mvarPR27PAN = vData
        End Set
    End Property

    Public Property PR26PAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26PAN = mvarPR26PAN
        End Get
        Set(ByVal vData As String)
            mvarPR26PAN = vData
        End Set
    End Property

    Public Property PR24PAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24PAN = mvarPR24PAN
        End Get
        Set(ByVal vData As String)
            mvarPR24PAN = vData
        End Set
    End Property

    Public Property TANRegNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            TANRegNo = mvarTANRegNo
        End Get
        Set(ByVal vData As String)
            mvarTANRegNo = vData
        End Set
    End Property
    Public Property AIN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            AIN = mvarAIN
        End Get
        Set(ByVal vData As String)
            mvarAIN = vData
        End Set
    End Property
    Public Property PR27EEmailAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EEmailAlt = mvarPR27EEmailAlt
        End Get
        Set(ByVal vData As String)
            mvarPR27EEmailAlt = vData
        End Set
    End Property


    Public Property PR27EPhoneAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EPhoneAlt = mvarPR27EPhoneAlt
        End Get
        Set(ByVal vData As String)
            mvarPR27EPhoneAlt = vData
        End Set
    End Property

    Public Property PR27EPhone As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EPhone = mvarPR27EPhone
        End Get
        Set(ByVal vData As String)
            mvarPR27EPhone = vData
        End Set
    End Property
    Public Property PR27ESTDAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27ESTDAlt = mvarPR27ESTDAlt
        End Get
        Set(ByVal vData As String)
            mvarPR27ESTDAlt = vData
        End Set
    End Property
    Public Property PR27EmailAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EmailAlt = mvarPR27EmailAlt
        End Get
        Set(ByVal vData As String)
            mvarPR27EmailAlt = vData
        End Set
    End Property
    Public Property PR27PhoneAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27PhoneAlt = mvarPR27PhoneAlt
        End Get
        Set(ByVal vData As String)
            mvarPR27PhoneAlt = vData
        End Set
    End Property

    Public Property PR27STDAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27STDAlt = mvarPR27STDAlt
        End Get
        Set(ByVal vData As String)
            mvarPR27STDAlt = vData
        End Set
    End Property

    Public Property PR26EmailAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26EmailAlt = mvarPR26EmailAlt
        End Get
        Set(ByVal vData As String)
            mvarPR26EmailAlt = vData
        End Set
    End Property

    Public Property PR26PhoneAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26PhoneAlt = mvarPR26PhoneAlt
        End Get
        Set(ByVal vData As String)
            mvarPR26PhoneAlt = vData
        End Set
    End Property

    Public Property PR26STDAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26STDAlt = mvarPR26STDAlt
        End Get
        Set(ByVal vData As String)
            mvarPR26STDAlt = vData
        End Set
    End Property

    Public Property PR24EmailAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24EmailAlt = mvarPR24EmailAlt
        End Get
        Set(ByVal vData As String)
            mvarPR24EmailAlt = vData
        End Set
    End Property

    Public Property PR24PhoneAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24PhoneAlt = mvarPR24PhoneAlt
        End Get
        Set(ByVal vData As String)
            mvarPR24PhoneAlt = vData
        End Set
    End Property

    Public Property PR24STDAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24STDAlt = mvarPR24STDAlt
        End Get
        Set(ByVal vData As String)
            mvarPR24STDAlt = vData
        End Set
    End Property


    Public Property CoEmailAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoEmailAlt = mvarCoEmailAlt
        End Get
        Set(ByVal vData As String)
            mvarCoEmailAlt = vData
        End Set
    End Property

    Public Property CoPhoneAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoPhoneAlt = mvarCoPhoneAlt
        End Get
        Set(ByVal vData As String)
            mvarCoPhoneAlt = vData
        End Set
    End Property

    Public Property CoSTDAlt As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoSTDAlt = mvarCoSTDAlt
        End Get
        Set(ByVal vData As String)
            mvarCoSTDAlt = vData
        End Set
    End Property

    Public Property UseForm16 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            UseForm16 = mvarUseForm16
        End Get
        Set(ByVal vData As String)
            mvarUseForm16 = vData
        End Set
    End Property

    Public Property IsPR27EPhone As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            IsPR27EPhone = mvarPR27EPhone
        End Get
        Set(ByVal vData As Integer)
            mvarPR27EPhone = vData
        End Set
    End Property
    Public Property PR27EStd As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EStd = mvarPR27EStd
        End Get
        Set(ByVal vData As String)
            mvarPR27EStd = vData
        End Set
    End Property

    Public Property PR27EEmail As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EEmail = mvarPR27EEmail
        End Get
        Set(ByVal vData As String)
            mvarPR27EEmail = vData
        End Set
    End Property

    Public Property PR27EPin As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EPin = mvarPR27EPin
        End Get
        Set(ByVal vData As Integer)
            mvarPR27EPin = vData
        End Set
    End Property

    Public Property PR27EStateID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EStateID = mvarPR27EStateID
        End Get
        Set(ByVal vData As String)
            mvarPR27EStateID = vData
        End Set
    End Property

    Public Property PR27EAdd5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EAdd5 = mvarPR27EAdd5
        End Get
        Set(ByVal vData As String)
            mvarPR27EAdd5 = vData
        End Set
    End Property

    Public Property PR27EAdd4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EAdd4 = mvarPR27EAdd4
        End Get
        Set(ByVal vData As String)
            mvarPR27EAdd4 = vData
        End Set
    End Property


    Public Property PR27EAdd3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EAdd3 = mvarPR27EAdd3
        End Get
        Set(ByVal vData As String)
            mvarPR27EAdd3 = vData
        End Set
    End Property
    Public Property PR27EAdd2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EAdd2 = mvarPR27EAdd2
        End Get
        Set(ByVal vData As String)
            mvarPR27EAdd2 = vData
        End Set
    End Property

    Public Property PR27EAdd1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EAdd1 = mvarPR27EAdd1
        End Get
        Set(ByVal vData As String)
            mvarPR27EAdd1 = vData
        End Set
    End Property
    Public Property PR27EDesg As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27EDesg = mvarPR27EDesg
        End Get
        Set(ByVal vData As String)
            mvarPR27EDesg = vData
        End Set
    End Property
    Public Property IsPR27EAddChg As Boolean
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            IsPR27EAddChg = mvarIsPR27EAddChg
        End Get
        Set(ByVal vData As Boolean)
            mvarIsPR27EAddChg = vData
        End Set
    End Property


    Public Property PR27Ename As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Ename = mvarPR27EName
        End Get
        Set(ByVal vData As String)
            mvarPR27EName = vData
        End Set
    End Property



    Public Property IsPR27AddChg As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            IsPR27AddChg = mvarIsPR27AddChg
        End Get
        Set(ByVal vData As String)
            mvarIsPR27AddChg = vData
        End Set
    End Property
    Public Property PR27Phone As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Phone = mvarPR27Phone
        End Get
        Set(ByVal vData As String)
            mvarPR27Phone = vData
        End Set
    End Property
    Public Property PR27Std As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Std = mvarPR27Std
        End Get
        Set(ByVal vData As String)
            mvarPR27Std = vData
        End Set
    End Property

    Public Property PR27Email As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Email = mvarPR27Email
        End Get
        Set(ByVal vData As String)
            mvarPR27Email = vData
        End Set
    End Property


    Public Property PR27Pin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Pin = mvarPR27Pin
        End Get
        Set(ByVal vData As String)
            mvarPR27Pin = vData
        End Set
    End Property

    Public Property PR27StateID As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27StateID = mvarPR27StateID
        End Get
        Set(ByVal vData As Integer)
            mvarPR27StateID = vData
        End Set
    End Property
    Public Property PR27Add5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Add5 = mvarPR27Add5
        End Get
        Set(ByVal vData As String)
            mvarPR27Add5 = vData
        End Set
    End Property


    Public Property PR27Add4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Add4 = mvarPR27Add4
        End Get
        Set(ByVal vData As String)
            mvarPR27Add4 = vData
        End Set
    End Property

    Public Property PR27Add3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Add3 = mvarPR27Add3
        End Get
        Set(ByVal vData As String)
            mvarPR27Add3 = vData
        End Set
    End Property

    Public Property PR27Add2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Add2 = mvarPR27Add2
        End Get
        Set(ByVal vData As String)
            mvarPR27Add2 = vData
        End Set
    End Property


    Public Property PR27Add1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR27Add1 = mvarPR27Add1
        End Get
        Set(ByVal vData As String)
            mvarPR27Add1 = vData
        End Set
    End Property



    Public Property IsPR26AddChg As Boolean
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            IsPR26AddChg = mvarIsPR26AddChg
        End Get
        Set(ByVal vData As Boolean)
            mvarIsPR26AddChg = vData
        End Set
    End Property

    Public Property PR26Phone As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Phone = mvarPR26Phone
        End Get
        Set(ByVal vData As String)
            mvarPR26Phone = vData
        End Set
    End Property
    Public Property PR26Std As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Std = mvarPR26Std
        End Get
        Set(ByVal vData As String)
            mvarPR26Std = vData
        End Set
    End Property
    Public Property PR26Email As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Email = mvarPR26Email
        End Get
        Set(ByVal vData As String)
            mvarPR26Email = vData
        End Set
    End Property

    Public Property PR26Pin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Pin = mvarPR26Pin
        End Get
        Set(ByVal vData As String)
            mvarPR26Pin = vData
        End Set
    End Property

    Public Property PR26StateID As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26StateID = mvarPR26StateID
        End Get
        Set(ByVal vData As Integer)
            mvarPR26StateID = vData
        End Set
    End Property

    Public Property PR26Add5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Add5 = mvarPR26Add5
        End Get
        Set(ByVal vData As String)
            mvarPR26Add5 = vData
        End Set
    End Property

    Public Property PR26Add4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Add4 = mvarPR26Add4
        End Get
        Set(ByVal vData As String)
            mvarPR26Add4 = vData
        End Set
    End Property

    Public Property PR26Add3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Add3 = mvarPR26Add3
        End Get
        Set(ByVal vData As String)
            mvarPR26Add3 = vData
        End Set
    End Property

    Public Property PR26Add2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Add2 = mvarPR26Add2
        End Get
        Set(ByVal vData As String)
            mvarPR26Add2 = vData
        End Set
    End Property

    Public Property PR26Add1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR26Add1 = mvarPR26Add1
        End Get
        Set(ByVal vData As String)
            mvarPR26Add1 = vData
        End Set
    End Property
    Public Property CoEmail As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoEmail = mvarCoEmail
        End Get
        Set(ByVal vData As String)
            mvarCoEmail = vData
        End Set
    End Property
    Public Property CoPhone As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoPhone = mvarCoPhone
        End Get
        Set(ByVal vData As String)
            mvarCoPhone = vData
        End Set
    End Property
    Public Property Comobile As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            Comobile = mvarCoMobile
        End Get
        Set(ByVal vData As String)
            mvarCoMobile = vData
        End Set
    End Property

    Public Property CoStd As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoStd = mvarCoStd
        End Get
        Set(ByVal vData As String)
            mvarCoStd = vData
        End Set
    End Property

    Public Property CoBrDiv As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoBrDiv = mvarCoBrDiv
        End Get
        Set(ByVal vData As String)
            mvarCoBrDiv = vData
        End Set
    End Property
    Public Property PR24Phone As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Phone = mvarPR24Phone
        End Get
        Set(ByVal vData As String)
            mvarPR24Phone = vData
        End Set
    End Property
    Public Property PR24Std As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Std = mvarPR24Std
        End Get
        Set(ByVal vData As String)
            mvarPR24Std = vData
        End Set
    End Property
    Public Property TuserId As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            TuserId = mvaruserid
        End Get
        Set(ByVal vData As String)
            mvaruserid = vData
        End Set
    End Property
    Public Property Tuserpsswrd As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            Tuserpsswrd = mvaruserpasswrd
        End Get
        Set(ByVal vData As String)
            mvaruserpasswrd = vData
        End Set
    End Property
    Public Property PR24Email As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Email = mvarPR24Email
        End Get
        Set(ByVal vData As String)
            mvarPR24Email = vData
        End Set
    End Property
    Public Sub FillCosInLvw(lvw As ListView)
        Dim rst As New DataSet, lstitm As ListView
        lvw.Items.Clear()
        rst = FetchDataSet("SELECT * FROM CoMst ORDER BY CoName")
        Do While Not rst.Tables(0).Rows.Count
            lstitm = lvw.Items.Add("C" & rst.Tables(0).Rows(0)("coid"), rst.Tables(0).Rows(0)("CoName"))
            'lstitm.Items(1) = lstitm.
            'rst.MoveNext
        Loop
        rst.Dispose()
        rst = Nothing
    End Sub



    Public Function FetchCo(ByVal coid As Integer) As clsCoMst
        Dim cCo As New clsCoMst
        Dim rst As New DataSet
        rst = FetchDataSet("SELECT * FROM CoMst WHERE CoID=" & coid)
        If rst.Tables(0).Rows.Count > 0 Then
            With cCo
                .CoAdd1 = rst.Tables(0).Rows(0)("CoAdd1").ToString()
                .CoAdd2 = rst.Tables(0).Rows(0)("CoAdd2").ToString()
                .CoAdd3 = rst.Tables(0).Rows(0)("CoAdd3").ToString()
                .CoAdd4 = rst.Tables(0).Rows(0)("CoAdd4").ToString()
                .CoAdd5 = rst.Tables(0).Rows(0)("CoAdd5").ToString()
                .coid = rst.Tables(0).Rows(0)("coid").ToString()
                .CoName = rst.Tables(0).Rows(0)("CoName").ToString()
                .CoBrDiv = rst.Tables(0).Rows(0)("CoBrDiv").ToString()
                .CoPAN = rst.Tables(0).Rows(0)("CoPAN").ToString()
                .CoPin = rst.Tables(0).Rows(0)("CoPin").ToString()
                .CoEmail = rst.Tables(0).Rows(0)("CoEmail").ToString()
                .CoStd = rst.Tables(0).Rows(0)("CoStd").ToString()
                .CoPhone = rst.Tables(0).Rows(0)("CoPhone").ToString()
                .Comobile = rst.Tables(0).Rows(0)("Mobile").ToString()
                .CoStateID = rst.Tables(0).Rows(0)("CoStateID").ToString()
                .CoStatus = rst.Tables(0).Rows(0)("CoStatus").ToString()
                .CoTAN = rst.Tables(0).Rows(0)("CoTAN").ToString()
                .IsCoAddChg = rst.Tables(0).Rows(0)("IsCoAddChg").ToString()
                .PRDesg24 = rst.Tables(0).Rows(0)("PR24desg").ToString()
                .PRDesg26 = rst.Tables(0).Rows(0)("PR26desg").ToString()
                .PRDesg27 = rst.Tables(0).Rows(0)("PR27desg").ToString()
                .PR27EDesg = rst.Tables(0).Rows(0)("PR27EDesg").ToString()
                .PRName24 = rst.Tables(0).Rows(0)("Pr24name").ToString()
                .PRName26 = rst.Tables(0).Rows(0)("Pr26name").ToString()
                .PRName27 = rst.Tables(0).Rows(0)("PR27name").ToString()
                .PR27Ename = rst.Tables(0).Rows(0)("PR27EName").ToString()


                .PR24Add1 = rst.Tables(0).Rows(0)("PR24Add1").ToString()
                .PR24Add2 = rst.Tables(0).Rows(0)("PR24Add2").ToString()
                .PR24Add3 = rst.Tables(0).Rows(0)("PR24Add3").ToString()
                .PR24Add4 = rst.Tables(0).Rows(0)("PR24Add4").ToString()
                .PR24Add5 = rst.Tables(0).Rows(0)("PR24Add5").ToString()
                .PR24StateID = rst.Tables(0).Rows(0)("PR24StateID").ToString()
                .PR24Pin = rst.Tables(0).Rows(0)("PR24Pin").ToString()
                .PR24Email = rst.Tables(0).Rows(0)("PR24Email").ToString()
                .PR24Std = rst.Tables(0).Rows(0)("PR24Std").ToString()
                .PR24Phone = rst.Tables(0).Rows(0)("PR24Phone").ToString()
                .IsPR24AddChg = rst.Tables(0).Rows(0)("IsPR24AddChg").ToString()

                .PR26Add1 = rst.Tables(0).Rows(0)("PR26Add1").ToString()
                .PR26Add2 = rst.Tables(0).Rows(0)("PR26Add2").ToString()
                .PR26Add3 = rst.Tables(0).Rows(0)("PR26Add3").ToString()
                .PR26Add4 = rst.Tables(0).Rows(0)("PR26Add4").ToString()
                .PR26Add5 = rst.Tables(0).Rows(0)("PR26Add5").ToString()
                .PR26StateID = rst.Tables(0).Rows(0)("PR26StateID").ToString()
                .PR26Pin = rst.Tables(0).Rows(0)("PR26Pin").ToString()
                .PR26Email = rst.Tables(0).Rows(0)("PR26Email").ToString()
                .PR26Std = rst.Tables(0).Rows(0)("PR26Std").ToString()
                .PR26Phone = rst.Tables(0).Rows(0)("PR26Phone").ToString()
                .IsPR26AddChg = rst.Tables(0).Rows(0)("IsPR26AddChg").ToString()

                .PR27Add1 = rst.Tables(0).Rows(0)("PR27Add1").ToString()
                .PR27Add2 = rst.Tables(0).Rows(0)("PR27Add2").ToString()
                .PR27Add3 = rst.Tables(0).Rows(0)("PR27Add3").ToString()
                .PR27Add4 = rst.Tables(0).Rows(0)("PR27Add4").ToString()
                .PR27Add5 = rst.Tables(0).Rows(0)("PR27Add5").ToString()
                .PR27StateID = rst.Tables(0).Rows(0)("PR27StateID").ToString()
                .PR27Pin = rst.Tables(0).Rows(0)("PR27Pin").ToString()
                .PR27Email = rst.Tables(0).Rows(0)("PR27Email").ToString()
                .PR27Std = rst.Tables(0).Rows(0)("PR27Std").ToString()
                .PR27Phone = rst.Tables(0).Rows(0)("PR27Phone").ToString()
                .IsPR27AddChg = rst.Tables(0).Rows(0)("IsPR27AddChg").ToString()

                .PR27EAdd1 = rst.Tables(0).Rows(0)("PR27EAdd1").ToString()
                .PR27EAdd2 = rst.Tables(0).Rows(0)("PR27EAdd2").ToString()
                .PR27EAdd3 = rst.Tables(0).Rows(0)("PR27EAdd3").ToString()
                .PR27EAdd4 = rst.Tables(0).Rows(0)("PR27EAdd4").ToString()
                .PR27EAdd5 = rst.Tables(0).Rows(0)("PR27EAdd5").ToString()
                .PR27EStateID = rst.Tables(0).Rows(0)("PR27EStateID").ToString()
                .PR27EPin = rst.Tables(0).Rows(0)("PR27EPin").ToString()
                .PR27EEmail = rst.Tables(0).Rows(0)("PR27EEmail").ToString()
                .PR27EStd = rst.Tables(0).Rows(0)("PR27EStd").ToString()
                .PR27EPhone = rst.Tables(0).Rows(0)("PR27EPhone").ToString()
                .IsPR27EAddChg = rst.Tables(0).Rows(0)("IsPR27EAddChg").ToString()
                .UseForm16 = rst.Tables(0).Rows(0)("UseForm16").ToString()
                .GovtStateID = rst.Tables(0).Rows(0)("GovtStateID").ToString()
                .PAOCode = rst.Tables(0).Rows(0)("PAOCode").ToString()
                .DDOCode = rst.Tables(0).Rows(0)("DDOCode").ToString()
                '.MinistryID = String.IsNullOrEmpty(rst.Tables(0).Rows(0)("MinistryID").ToString())
                .MinistryID = rst.Tables(0).Rows(0)("MinistryID").ToString()
                If String.IsNullOrEmpty(rst.Tables(0).Rows(0)("MinistryID").ToString()) = 99 Then
                    .MinistryName = rst.Tables(0).Rows(0)("MinistryName").ToString()
                Else
                    .MinistryName = vbNullString
                End If
                .PAORegNo = rst.Tables(0).Rows(0)("PAORegNo").ToString()
                .DDORegNo = rst.Tables(0).Rows(0)("DDORegNo").ToString()
                .TuserId = rst.Tables(0).Rows(0)("TanUserID").ToString()
                .Tuserpsswrd = rst.Tables(0).Rows(0)("TANPAssword").ToString()
                .TANRegNo = rst.Tables(0).Rows(0)("TANRegNo").ToString()
                'Alternate Contact Details - Added on 05/07/2013 - by Nitin
                .CoSTDAlt = rst.Tables(0).Rows(0)("CoSTDAlt").ToString()
                .CoPhoneAlt = rst.Tables(0).Rows(0)("CoPhoneAlt").ToString()
                .CoEmailAlt = rst.Tables(0).Rows(0)("CoEmailAlt").ToString()

                .PR24STDAlt = rst.Tables(0).Rows(0)("PR24STDAlt").ToString()
                .PR24PhoneAlt = rst.Tables(0).Rows(0)("PR24PhoneAlt").ToString()
                .PR24EmailAlt = rst.Tables(0).Rows(0)("PR24EmailAlt").ToString()

                .PR26STDAlt = rst.Tables(0).Rows(0)("PR26STDAlt").ToString()
                .PR26PhoneAlt = rst.Tables(0).Rows(0)("PR26PhoneAlt").ToString()
                .PR26EmailAlt = rst.Tables(0).Rows(0)("PR26EmailAlt").ToString()

                .PR27STDAlt = rst.Tables(0).Rows(0)("PR27STDAlt").ToString()
                .PR27PhoneAlt = rst.Tables(0).Rows(0)("PR27PhoneAlt").ToString()
                .PR27EmailAlt = rst.Tables(0).Rows(0)("PR27EmailAlt").ToString()

                .PR27ESTDAlt = rst.Tables(0).Rows(0)("PR27ESTDAlt").ToString()
                .PR27EPhoneAlt = rst.Tables(0).Rows(0)("PR27EPhoneAlt").ToString()
                .PR27EEmailAlt = rst.Tables(0).Rows(0)("PR27EEmailAlt").ToString()

                .AIN = rst.Tables(0).Rows(0)("AIN").ToString()
                .PR24PAN = rst.Tables(0).Rows(0)("PR24PAN").ToString()
                .PR26PAN = rst.Tables(0).Rows(0)("PR26PAN").ToString()
                .PR27PAN = rst.Tables(0).Rows(0)("PR27PAN").ToString()
                .PR27EPAN = rst.Tables(0).Rows(0)("PR27EPAN").ToString()
                .gstin = rst.Tables(0).Rows(0)("gstin").ToString()
            End With
            FetchCo = cCo
        Else
            FetchCo = Nothing
        End If
        rst.Dispose()
        rst = Nothing
        cCo = Nothing
    End Function

    Public Function Delete(coid As Integer) As Boolean
        Dim cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
        Dim sql As String
        'On Error GoTo DelErr
        cnl = False : RaiseEvent BeforeDelete(cnl)
        If cnl = True Then
            Delete = False
            Exit Function
        End If
        sql = " Delete * from challan26Q where retnid in(select retnid from retnmst where coid=" & coid & ")"
        sql = " Delete * from challan24Q where retnid in(select retnid from retnmst where coid=" & coid & ")"
        sql = " Delete * from deductee24Q where did in(select did from deductmst where coid=" & coid & ")"
        sql = " Delete * from deductee26Q where did in(select did from deductmst where coid=" & coid & ")"
        sql = "Delete * From deductMst Where CoID = " & coid
        sql = "Delete * From BankMst Where CoID = " & coid
        sql = "Delete * From retnMst Where CoID = " & coid
        sql = "Delete * From CoMst Where CoID = " & coid


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
DelErr:
        MsgBox(Err.Description, , Err.Number)
    End Function

    Public Function Update(ByVal CoMstObj As clsCoMst) As Boolean
        Dim sql As String, cnl As Boolean
        Dim transaction As OleDb.OleDbTransaction
        Dim cmd As New OleDb.OleDbCommand
        'On Error GoTo UpErr
        cnl = False : RaiseEvent PrepareDataForSave(cnl)
        If cnl = True Then Exit Function
        cnl = False : RaiseEvent BeforeSave(cnl)
        If cnl = True Then Exit Function

        With CoMstObj
            sql = "Update CoMSt Set CoName = " & IIf(.CoName = vbNullString, "Null", Chr(34) & .CoName & Chr(34)) & "," _
        & " CoBrDiv = " & IIf(.CoBrDiv = vbNullString, "Null", Chr(34) & .CoBrDiv & Chr(34)) & "," _
        & " CoAdd1 = " & IIf(.CoAdd1 = vbNullString, "Null", Chr(34) & .CoAdd1 & Chr(34)) & "," _
        & " CoAdd2 = " & IIf(.CoAdd2 = vbNullString, "Null", Chr(34) & .CoAdd2 & Chr(34)) & "," _
        & " CoAdd3 = " & IIf(.CoAdd3 = vbNullString, "Null", Chr(34) & .CoAdd3 & Chr(34)) & "," _
        & " CoAdd4 = " & IIf(.CoAdd4 = vbNullString, "Null", Chr(34) & .CoAdd4 & Chr(34)) & "," _
        & " CoAdd5 = " & IIf(.CoAdd5 = vbNullString, "Null", Chr(34) & .CoAdd5 & Chr(34)) & "," _
        & " CoStateID = " & IIf(.CoStateID = 0, 0, .CoStateID) & ", CoPin = " & IIf(.CoPin = 0, 0, .CoPin) & "," _
        & " CoTAN = " & IIf(.CoTAN = vbNullString, "Null", Chr(34) & .CoTAN & Chr(34)) & "," _
        & " CoPAN = " & IIf(.CoPAN = vbNullString, "Null", Chr(34) & .CoPAN & Chr(34)) & "," _
        & " IsCoAddChg = " & IIf(.IsCoAddChg = True, True, False) & "," _
        & " CoStatus = " & IIf(.CoStatus = vbNullString, "Null", Chr(34) & .CoStatus & Chr(34)) & "," _
        & " CoEmail = " & IIf(.CoEmail = vbNullString, "Null", Chr(34) & .CoEmail & Chr(34)) & "," _
        & " CoStd = " & IIf(.CoStd = vbNullString, "Null", Chr(34) & .CoStd & Chr(34)) & "," _
        & " CoPhone = " & IIf(.CoPhone = vbNullString, "Null", Chr(34) & .CoPhone & Chr(34)) & "," _
        & " PR26Name = " & IIf(.PRName26 = vbNullString, "Null", Chr(34) & .PRName26 & Chr(34)) & "," _
        & " PR26Desg = " & IIf(.PRDesg26 = vbNullString, "Null", Chr(34) & .PRDesg26 & Chr(34)) & "," _
        & " PR27Name = " & IIf(.PRName27 = vbNullString, "Null", Chr(34) & .PRName27 & Chr(34)) & "," _
        & " PR27Desg = " & IIf(.PRDesg27 = vbNullString, "Null", Chr(34) & .PRDesg27 & Chr(34)) & ","

            sql = sql & " PR24Name = " & IIf(.PRName24 = vbNullString, "Null", Chr(34) & .PRName24 & Chr(34)) & "," _
        & " PR24Desg = " & IIf(.PRDesg24 = vbNullString, "Null", Chr(34) & .PRDesg24 & Chr(34)) & "," _
        & " PR27EName = " & IIf(.PR27Ename = vbNullString, "Null", Chr(34) & .PR27Ename & Chr(34)) & "," _
        & " PR27EDesg = " & IIf(.PR27EDesg = vbNullString, "Null", Chr(34) & .PR27EDesg & Chr(34)) & "," _
        & " PR24Add1 = " & IIf(.PR24Add1 = vbNullString, "Null", Chr(34) & .PR24Add1 & Chr(34)) & "," _
        & " PR24Add2 = " & IIf(.PR24Add2 = vbNullString, "Null", Chr(34) & .PR24Add2 & Chr(34)) & "," _
        & " PR24Add3 = " & IIf(.PR24Add3 = vbNullString, "Null", Chr(34) & .PR24Add3 & Chr(34)) & "," _
        & " PR24Add4 = " & IIf(.PR24Add4 = vbNullString, "Null", Chr(34) & .PR24Add4 & Chr(34)) & "," _
        & " PR24Add5 = " & IIf(.PR24Add5 = vbNullString, "Null", Chr(34) & .PR24Add5 & Chr(34)) & "," _
        & " PR24StateID = " & IIf(.PR24StateID = 0, 0, .PR24StateID) & "," _
        & " PR24Pin = " & IIf(.PR24Pin = 0, 0, .PR24Pin) & ", IsPR24AddChg = " & IIf(.IsPR24AddChg = True, True, False) & "," _
        & " PR24Email = " & IIf(.PR24Email = vbNullString, "Null", Chr(34) & .PR24Email & Chr(34)) & "," _
        & " PR24Std = " & IIf(.PR24Std = vbNullString, "Null", .PR24Std) & "," _
        & " PR24Phone = " & IIf(.PR24Phone = vbNullString, "Null", .PR24Phone) & ","

            sql = sql & " PR26Add1 = " & IIf(.PR26Add1 = vbNullString, "Null", Chr(34) & .PR26Add1 & Chr(34)) & "," _
        & " PR26Add2 = " & IIf(.PR26Add2 = vbNullString, "Null", Chr(34) & .PR26Add2 & Chr(34)) & "," _
        & " PR26Add3 = " & IIf(.PR26Add3 = vbNullString, "Null", Chr(34) & .PR26Add3 & Chr(34)) & "," _
        & " PR26Add4 = " & IIf(.PR26Add4 = vbNullString, "Null", Chr(34) & .PR26Add4 & Chr(34)) & "," _
        & " PR26Add5 = " & IIf(.PR26Add5 = vbNullString, "Null", Chr(34) & .PR26Add5 & Chr(34)) & "," _
        & " PR26StateID = " & IIf(.PR26StateID = 0, 0, .PR26StateID) & "," _
        & " PR26Pin = " & IIf(.PR26Pin = 0, 0, .PR26Pin) & ", IsPR26AddChg = " & IIf(.IsPR26AddChg = True, True, False) & "," _
        & " PR26Email = " & IIf(.PR26Email = vbNullString, "Null", Chr(34) & .PR26Email & Chr(34)) & "," _
        & " PR26Std = " & IIf(.PR26Std = vbNullString, "Null", .PR26Std) & "," _
        & " PR26Phone = " & IIf(.PR26Phone = vbNullString, "Null", .PR26Phone) & ","

            sql = sql & " PR27Add1 = " & IIf(.PR27Add1 = vbNullString, "Null", Chr(34) & .PR27Add1 & Chr(34)) & "," _
        & " PR27Add2 = " & IIf(.PR27Add2 = vbNullString, "Null", Chr(34) & .PR27Add2 & Chr(34)) & "," _
        & " PR27Add3 = " & IIf(.PR27Add3 = vbNullString, "Null", Chr(34) & .PR27Add3 & Chr(34)) & "," _
        & " PR27Add4 = " & IIf(.PR27Add4 = vbNullString, "Null", Chr(34) & .PR27Add4 & Chr(34)) & "," _
        & " PR27Add5 = " & IIf(.PR27Add5 = vbNullString, "Null", Chr(34) & .PR27Add5 & Chr(34)) & "," _
        & " PR27StateID = " & IIf(.PR27StateID = 0, 0, .PR27StateID) & "," _
        & " PR27Pin = " & IIf(.PR27Pin = 0, 0, .PR27Pin) & ", IsPR27AddChg = " & IIf(.IsPR27AddChg = True, True, False) & "," _
        & " PR27Email = " & IIf(.PR27Email = vbNullString, "Null", Chr(34) & .PR27Email & Chr(34)) & "," _
        & " PR27Std = " & IIf(.PR27Std = vbNullString, "Null", .PR27Std) & "," _
        & " PR27Phone = " & IIf(.PR27Phone = vbNullString, "Null", .PR27Phone) & ","

            sql = sql & " PR27EAdd1 = " & IIf(.PR27EAdd1 = vbNullString, "Null", Chr(34) & .PR27EAdd1 & Chr(34)) & "," _
        & " PR27EAdd2 = " & IIf(.PR27EAdd2 = vbNullString, "Null", Chr(34) & .PR27EAdd2 & Chr(34)) & "," _
        & " PR27EAdd3 = " & IIf(.PR27EAdd3 = vbNullString, "Null", Chr(34) & .PR27EAdd3 & Chr(34)) & "," _
        & " PR27EAdd4 = " & IIf(.PR27EAdd4 = vbNullString, "Null", Chr(34) & .PR27EAdd4 & Chr(34)) & "," _
        & " PR27EAdd5 = " & IIf(.PR27EAdd5 = vbNullString, "Null", Chr(34) & .PR27EAdd5 & Chr(34)) & "," _
        & " PR27EStateID = " & IIf(.PR27EStateID = 0, 0, .PR27EStateID) & "," _
        & " PR27EPin = " & IIf(.PR27EPin = 0, 0, .PR27EPin) & ", IsPR27EAddChg = " & IIf(.IsPR27EAddChg = True, True, False) & "," _
        & " PR27EEmail = " & IIf(.PR27EEmail = vbNullString, "Null", Chr(34) & .PR27EEmail & Chr(34)) & "," _
        & " PR27EStd = " & IIf(.PR27EStd = vbNullString, "Null", .PR27EStd) & "," _
        & " PR27EPhone = " & IIf(.PR27EPhone = vbNullString, "Null", .PR27EPhone) & "," _
        & " UseForm16 = " & IIf(.UseForm16 = True, True, False) & "," _
        & " GovtStateID = " & IIf(.GovtStateID = -1, "Null", .GovtStateID) & "," _
        & " PAOCode = " & IIf(.PAOCode = vbNullString, "Null", Chr(34) & .PAOCode & Chr(34)) & "," _
        & " DDOCode = " & IIf(.DDOCode = vbNullString, "Null", Chr(34) & .DDOCode & Chr(34)) & "," _
        & " MinistryID = " & IIf(.MinistryID = -1, "Null", .MinistryID) & "," _
        & " MinistryName = " & IIf(.MinistryName = vbNullString, "Null", Chr(34) & .MinistryName & Chr(34)) & "," _
        & " PAORegNo = " & IIf(.PAORegNo = 0, 0, .PAORegNo) & "," _
        & " DDORegNo = " & IIf(.DDORegNo = vbNullString, "Null", Chr(34) & .DDORegNo & Chr(34)) & ", Mobile=" & .Comobile & ","

            'Added for Alternative Contact details - 06/07/2013 - By Nitin
            sql = sql _
        & " CoEmailAlt = " & IIf(.CoEmailAlt = vbNullString, "Null", Chr(34) & .CoEmailAlt & Chr(34)) & "," _
        & " CoSTDAlt = " & IIf(.CoSTDAlt = vbNullString, "Null", Chr(34) & .CoSTDAlt & Chr(34)) & "," _
        & " CoPhoneAlt = " & IIf(.CoPhoneAlt = vbNullString, "Null", Chr(34) & .CoPhoneAlt & Chr(34)) & "," _
        & " PR24EmailAlt = " & IIf(.PR24EmailAlt = vbNullString, "Null", Chr(34) & .PR24EmailAlt & Chr(34)) & "," _
        & " PR24STDAlt = " & IIf(.PR24STDAlt = vbNullString, "Null", Chr(34) & .PR24STDAlt & Chr(34)) & "," _
        & " PR24PhoneAlt = " & IIf(.PR24PhoneAlt = vbNullString, "Null", Chr(34) & .PR24PhoneAlt & Chr(34)) & "," _
        & " PR26EmailAlt = " & IIf(.PR26EmailAlt = vbNullString, "Null", Chr(34) & .PR26EmailAlt & Chr(34)) & "," _
        & " PR26STDAlt = " & IIf(.PR26STDAlt = vbNullString, "Null", Chr(34) & .PR26STDAlt & Chr(34)) & "," _
        & " PR26PhoneAlt = " & IIf(.PR26PhoneAlt = vbNullString, "Null", Chr(34) & .PR26PhoneAlt & Chr(34)) & "," _
        & " PR27EmailAlt = " & IIf(.PR27EmailAlt = vbNullString, "Null", Chr(34) & .PR27EmailAlt & Chr(34)) & "," _
        & " PR27STDAlt = " & IIf(.PR27STDAlt = vbNullString, "Null", Chr(34) & .PR27STDAlt & Chr(34)) & "," _
        & " PR27PhoneAlt = " & IIf(.PR27PhoneAlt = vbNullString, "Null", Chr(34) & .PR27PhoneAlt & Chr(34)) & "," _
        & " PR27EEmailAlt = " & IIf(.PR27EEmailAlt = vbNullString, "Null", Chr(34) & .PR27EEmailAlt & Chr(34)) & "," _
        & " PR27ESTDAlt = " & IIf(.PR27ESTDAlt = vbNullString, "Null", Chr(34) & .PR27ESTDAlt & Chr(34)) & "," _
        & " PR27EPhoneAlt = " & IIf(.PR27EPhoneAlt = vbNullString, "Null", Chr(34) & .PR27EPhoneAlt & Chr(34)) & "," _
        & " AIN = " & IIf(.AIN = vbNullString, "Null", Chr(34) & .AIN & Chr(34)) & "," _
        & " PR24PAN = " & IIf(.PR24PAN = vbNullString, "Null", Chr(34) & .PR24PAN & Chr(34)) & "," _
        & " PR26PAN = " & IIf(.PR26PAN = vbNullString, "Null", Chr(34) & .PR26PAN & Chr(34)) & "," _
        & " PR27PAN = " & IIf(.PR27PAN = vbNullString, "Null", Chr(34) & .PR27PAN & Chr(34)) & "," _
        & " PR27EPAN = " & IIf(.PR27EPAN = vbNullString, "Null", Chr(34) & .PR27EPAN & Chr(34)) & "," _
        & " gstin = " & IIf(.gstin = vbNullString, "Null", Chr(34) & .gstin & Chr(34)) _
        & " Where CoId = " & .coid
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

    Public Function Insert(ByVal CoMstObj As clsCoMst) As Boolean
        Dim sql As String, cnl As Boolean
        Dim cmd As New OleDb.OleDbCommand
        Dim transaction As OleDb.OleDbTransaction
        '  On Error GoTo InErr
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
        With CoMstObj
            sql = "Insert Into CoMSt (CoId,CoName,CoBrDiv,CoAdd1,CoAdd2,CoAdd3,CoAdd4,CoAdd5,CoStateID, " _
        & "CoPin,CoTAN,CoPAN,IsCoAddChg,CoEmail,CoStd,CoPhone,CoStatus,PR26Name,PR26Desg,PR27Name,PR27Desg,PR24Name,PR24Desg,PR27EName,PR27EDesg, " _
        & "PR24Add1,PR24Add2,PR24Add3,PR24Add4,PR24Add5,PR24StateID,PR24Pin,IsPR24AddChg,PR24Email,PR24Std,PR24Phone, " _
        & "PR26Add1,PR26Add2,PR26Add3,PR26Add4,PR26Add5,PR26StateID,PR26Pin,IsPR26AddChg,PR26Email,PR26Std,PR26Phone, " _
        & "PR27Add1,PR27Add2,PR27Add3,PR27Add4,PR27Add5,PR27StateID,PR27Pin,IsPR27AddChg,PR27Email,PR27Std,PR27Phone, " _
        & "PR27EAdd1,PR27EAdd2,PR27EAdd3,PR27EAdd4,PR27EAdd5,PR27EStateID,PR27EPin,IsPR27EAddChg,PR27EEmail,PR27EStd, " _
        & "PR27EPhone,UseForm16,GovtStateID,PAOCode,DDOCode,MinistryID,MinistryName,PAORegNo,DDORegNo,mobile, " _
        & "CoEmailAlt,CoSTDAlt, CoPhoneAlt, PR24EmailAlt, PR24STDAlt, PR24PhoneAlt, PR26EmailAlt,PR26STDAlt, PR26PhoneAlt, " _
        & "PR27EmailAlt, PR27STDAlt, PR27PhoneAlt, PR27EEmailAlt, PR27ESTDAlt, PR27EPhoneAlt, AIN, PR24PAN, PR26PAN, PR27PAN, PR27EPAN ,GSTIN" _
        & ") Values ( "

            sql = sql & MaxID() + 1 & "," & IIf(.CoName = vbNullString, "Null", Chr(34) & .CoName & Chr(34)) & "," _
        & IIf(.CoBrDiv = vbNullString, "Null", Chr(34) & .CoBrDiv & Chr(34)) & "," _
        & IIf(.CoAdd1 = vbNullString, "Null", Chr(34) & .CoAdd1 & Chr(34)) & "," & IIf(.CoAdd2 = vbNullString, "Null", Chr(34) & .CoAdd2 & Chr(34)) & "," _
        & IIf(.CoAdd3 = vbNullString, "Null", Chr(34) & .CoAdd3 & Chr(34)) & "," & IIf(.CoAdd4 = vbNullString, "Null", Chr(34) & .CoAdd4 & Chr(34)) & "," _
        & IIf(.CoAdd5 = vbNullString, "Null", Chr(34) & .CoAdd5 & Chr(34)) & "," & IIf(.CoStateID = 0, 0, .CoStateID) & ", " & IIf(.CoPin = 0, 0, .CoPin) & "," _
        & IIf(.CoTAN = vbNullString, "Null", Chr(34) & .CoTAN & Chr(34)) & "," & IIf(.CoPAN = vbNullString, "Null", Chr(34) & .CoPAN & Chr(34)) & "," _
        & IIf(.IsCoAddChg = True, True, False) & "," _
        & IIf(.CoEmail = vbNullString, "Null", Chr(34) & .CoEmail & Chr(34)) & "," _
        & IIf(.CoStd = vbNullString, "Null", Chr(34) & .CoStd & Chr(34)) & "," _
        & IIf(.CoPhone = vbNullString, "Null", Chr(34) & .CoPhone & Chr(34)) & ","

            sql = sql & IIf(.CoStatus = vbNullString, "Null", Chr(34) & .CoStatus & Chr(34)) & "," _
        & IIf(.PRName26 = vbNullString, "Null", Chr(34) & .PRName26 & Chr(34)) & "," _
        & IIf(.PRDesg26 = vbNullString, "Null", Chr(34) & .PRDesg26 & Chr(34)) & "," _
        & IIf(.PRName27 = vbNullString, "Null", Chr(34) & .PRName27 & Chr(34)) & "," _
        & IIf(.PRDesg27 = vbNullString, "Null", Chr(34) & .PRDesg27 & Chr(34)) & "," _
        & IIf(.PRName24 = vbNullString, "Null", Chr(34) & .PRName24 & Chr(34)) & "," _
        & IIf(.PRDesg24 = vbNullString, "Null", Chr(34) & .PRDesg24 & Chr(34)) & "," _
        & IIf(.PR27Ename = vbNullString, "Null", Chr(34) & .PR27Ename & Chr(34)) & "," _
        & IIf(.PR27EDesg = vbNullString, "Null", Chr(34) & .PR27EDesg & Chr(34)) & ","

            sql = sql & IIf(.PR24Add1 = vbNullString, "Null", Chr(34) & .PR24Add1 & Chr(34)) & "," _
        & IIf(.PR24Add2 = vbNullString, "Null", Chr(34) & .PR24Add2 & Chr(34)) & "," _
        & IIf(.PR24Add3 = vbNullString, "Null", Chr(34) & .PR24Add3 & Chr(34)) & "," _
        & IIf(.PR24Add4 = vbNullString, "Null", Chr(34) & .PR24Add4 & Chr(34)) & "," _
        & IIf(.PR24Add5 = vbNullString, "Null", Chr(34) & .PR24Add5 & Chr(34)) & "," _
        & IIf(.PR24StateID = 0, 0, .PR24StateID) & "," _
        & IIf(.PR24Pin = 0, 0, .PR24Pin) & ", " & IIf(.IsPR24AddChg = True, True, False) & "," _
        & IIf(.PR24Email = vbNullString, "Null", Chr(34) & .PR24Email & Chr(34)) & "," _
        & IIf(.PR24Std = vbNullString, "Null", Chr(34) & .PR24Std & Chr(34)) & "," _
        & IIf(.PR24Phone = vbNullString, "Null", Chr(34) & .PR24Phone & Chr(34)) & ","

            sql = sql & IIf(.PR26Add1 = vbNullString, "Null", Chr(34) & .PR26Add1 & Chr(34)) & "," _
        & IIf(.PR26Add2 = vbNullString, "Null", Chr(34) & .PR26Add2 & Chr(34)) & "," _
        & IIf(.PR26Add3 = vbNullString, "Null", Chr(34) & .PR26Add3 & Chr(34)) & "," _
        & IIf(.PR26Add4 = vbNullString, "Null", Chr(34) & .PR26Add4 & Chr(34)) & "," _
        & IIf(.PR26Add5 = vbNullString, "Null", Chr(34) & .PR26Add5 & Chr(34)) & "," _
        & IIf(.PR26StateID = 0, 0, .PR26StateID) & "," _
        & IIf(.PR26Pin = 0, 0, .PR26Pin) & ", " & IIf(.IsPR26AddChg = True, True, False) & "," _
        & IIf(.PR26Email = vbNullString, "Null", Chr(34) & .PR26Email & Chr(34)) & "," _
        & IIf(.PR26Std = vbNullString, "Null", Chr(34) & .PR26Std & Chr(34)) & "," _
        & IIf(.PR26Phone = vbNullString, "Null", Chr(34) & .PR26Phone & Chr(34)) & ","

            sql = sql & IIf(.PR27Add1 = vbNullString, "Null", Chr(34) & .PR27Add1 & Chr(34)) & "," _
        & IIf(.PR27Add2 = vbNullString, "Null", Chr(34) & .PR27Add2 & Chr(34)) & "," _
        & IIf(.PR27Add3 = vbNullString, "Null", Chr(34) & .PR27Add3 & Chr(34)) & "," _
        & IIf(.PR27Add4 = vbNullString, "Null", Chr(34) & .PR27Add4 & Chr(34)) & "," _
        & IIf(.PR27Add5 = vbNullString, "Null", Chr(34) & .PR27Add5 & Chr(34)) & "," _
        & IIf(.PR27StateID = 0, 0, .PR27StateID) & "," _
        & IIf(.PR27Pin = 0, 0, .PR27Pin) & ", " & IIf(.IsPR27AddChg = True, True, False) & "," _
        & IIf(.PR27Email = vbNullString, "Null", Chr(34) & .PR27Email & Chr(34)) & "," _
        & IIf(.PR27Std = vbNullString, "Null", Chr(34) & .PR27Std & Chr(34)) & "," _
        & IIf(.PR27Phone = vbNullString, "Null", Chr(34) & .PR27Phone & Chr(34)) & ","


            sql = sql & IIf(.PR27EAdd1 = vbNullString, "Null", Chr(34) & .PR27EAdd1 & Chr(34)) & "," _
        & IIf(.PR27EAdd2 = vbNullString, "Null", Chr(34) & .PR27EAdd2 & Chr(34)) & "," _
        & IIf(.PR27EAdd3 = vbNullString, "Null", Chr(34) & .PR27EAdd3 & Chr(34)) & "," _
        & IIf(.PR27EAdd4 = vbNullString, "Null", Chr(34) & .PR27EAdd4 & Chr(34)) & "," _
        & IIf(.PR27EAdd5 = vbNullString, "Null", Chr(34) & .PR27EAdd5 & Chr(34)) & "," _
        & IIf(.PR27EStateID = 0, 0, .PR27EStateID) & "," _
        & IIf(.PR27EPin = 0, 0, .PR27EPin) & ", " & IIf(.IsPR27EAddChg = True, True, False) & "," _
        & IIf(.PR27EEmail = vbNullString, "Null", Chr(34) & .PR27EEmail & Chr(34)) & "," _
        & IIf(.PR27EStd = vbNullString, "Null", Chr(34) & .PR27EStd & Chr(34)) & "," _
        & IIf(.PR27EPhone = vbNullString, "Null", Chr(34) & .PR27EPhone & Chr(34)) & "," _
        & IIf(.UseForm16 = True, True, False) & "," _
        & IIf(.GovtStateID = -1, "Null", .GovtStateID) & "," _
        & IIf(.PAOCode = vbNullString, "Null", Chr(34) & .PAOCode & Chr(34)) & "," _
        & IIf(.DDOCode = vbNullString, "Null", Chr(34) & .DDOCode & Chr(34)) & "," _
        & IIf(.MinistryID = -1, "Null", .MinistryID) & "," _
        & IIf(.MinistryName = vbNullString, "Null", Chr(34) & .MinistryName & Chr(34)) & "," _
        & IIf(.PAORegNo = 0, 0, .PAORegNo) & "," _
        & IIf(.DDORegNo = vbNullString, "Null", Chr(34) & .DDORegNo & Chr(34)) & "," & Chr(34) & .Comobile & Chr(34) & ","

            'Added for Alternative Contact details - 06/07/2013 - By Nitin
            sql = sql & IIf(.CoEmailAlt = vbNullString, "Null", Chr(34) & .CoEmailAlt & Chr(34)) & "," _
        & IIf(.CoSTDAlt = vbNullString, "Null", Chr(34) & .CoSTDAlt & Chr(34)) & "," _
        & IIf(.CoPhoneAlt = vbNullString, "Null", Chr(34) & .CoPhoneAlt & Chr(34)) & "," _
        & IIf(.PR24EmailAlt = vbNullString, "Null", Chr(34) & .PR24EmailAlt & Chr(34)) & "," _
        & IIf(.PR24STDAlt = vbNullString, "Null", Chr(34) & .PR24STDAlt & Chr(34)) & "," _
        & IIf(.PR24PhoneAlt = vbNullString, "Null", Chr(34) & .PR24PhoneAlt & Chr(34)) & "," _
        & IIf(.PR26EmailAlt = vbNullString, "Null", Chr(34) & .PR26EmailAlt & Chr(34)) & "," _
        & IIf(.PR26STDAlt = vbNullString, "Null", Chr(34) & .PR26STDAlt & Chr(34)) & "," _
        & IIf(.PR26PhoneAlt = vbNullString, "Null", Chr(34) & .PR26PhoneAlt & Chr(34)) & "," _
        & IIf(.PR27EmailAlt = vbNullString, "Null", Chr(34) & .PR27EmailAlt & Chr(34)) & "," _
        & IIf(.PR27STDAlt = vbNullString, "Null", Chr(34) & .PR27STDAlt & Chr(34)) & "," _
        & IIf(.PR27PhoneAlt = vbNullString, "Null", Chr(34) & .PR27PhoneAlt & Chr(34)) & "," _
        & IIf(.PR27EEmailAlt = vbNullString, "Null", Chr(34) & .PR27EEmailAlt & Chr(34)) & "," _
        & IIf(.PR27ESTDAlt = vbNullString, "Null", Chr(34) & .PR27ESTDAlt & Chr(34)) & "," _
        & IIf(.PR27EPhoneAlt = vbNullString, "Null", Chr(34) & .PR27EPhoneAlt & Chr(34)) & "," _
        & IIf(.AIN = vbNullString, "Null", Chr(34) & .AIN & Chr(34)) & "," _
        & IIf(.PR24PAN = vbNullString, "Null", Chr(34) & .PR24PAN & Chr(34)) & "," _
        & IIf(.PR26PAN = vbNullString, "Null", Chr(34) & .PR26PAN & Chr(34)) & "," _
        & IIf(.PR27PAN = vbNullString, "Null", Chr(34) & .PR27PAN & Chr(34)) & "," _
        & IIf(.PR27EPAN = vbNullString, "Null", Chr(34) & .PR27EPAN & Chr(34)) & "," _
        & IIf(.gstin = vbNullString, "Null", Chr(34) & .gstin & Chr(34)) & ")"
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

        'InErr:
        '        MsgBox(Err.Description, , Err.Number)
    End Function

    Public Function MaxID() As Integer
        Dim rs As New DataSet
        Dim sql As String
        sql = "Select Max(CoId) as ID From CoMSt"
        rs = FetchDataSet(sql)
        If Not String.IsNullOrEmpty(rs.Tables(0).Rows(0)("ID").ToString()) Then
            MaxID = rs.Tables(0).Rows(0)("ID").ToString()
        Else
            MaxID = 0
        End If

        rs.Dispose()
        rs = Nothing
    End Function
    Public Property IsPR24AddChg As Boolean
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            IsPR24AddChg = mvarIsPR24AddChg
        End Get
        Set(ByVal vData As Boolean)
            mvarIsPR24AddChg = vData
        End Set
    End Property
    Public Property PR24Pin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Pin = mvarPR24Pin
        End Get
        Set(ByVal vData As String)
            mvarPR24Pin = vData
        End Set
    End Property
    Public Property PR24StateID As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24StateID = mvarPR24StateID
        End Get
        Set(ByVal vData As Integer)
            mvarPR24StateID = vData
        End Set
    End Property
    Public Property PR24Add5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Add5 = mvarPR24Add5
        End Get
        Set(ByVal vData As String)
            mvarPR24Add5 = vData
        End Set
    End Property
    Public Property PR24Add4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Add4 = mvarPR24Add4
        End Get
        Set(ByVal vData As String)
            mvarPR24Add4 = vData
        End Set
    End Property

    Public Property PR24Add3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Add3 = mvarPR24Add3
        End Get
        Set(ByVal vData As String)
            mvarPR24Add3 = vData
        End Set
    End Property
    Public Property PR24Add2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Add2 = mvarPR24Add2
        End Get
        Set(ByVal vData As String)
            mvarPR24Add2 = vData
        End Set
    End Property

    Public Property PR24Add1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PR24Add1 = mvarPR24Add1
        End Get
        Set(ByVal vData As String)
            mvarPR24Add1 = vData
        End Set
    End Property

    Public Property PRDesg24 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PRDesg24 = mvarPRDesg24
        End Get
        Set(ByVal vData As String)
            mvarPRDesg24 = vData
        End Set
    End Property


    Public Property PRName24 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PRName24 = mvarPRName24
        End Get
        Set(ByVal vData As String)
            mvarPRName24 = vData
        End Set
    End Property

    Public Property PRDesg27 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PRDesg27 = mvarPRDesg27
        End Get
        Set(ByVal vData As String)
            mvarPRDesg27 = vData
        End Set
    End Property

    Public Property PRName27 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PRName27 = mvarPRName27
        End Get
        Set(ByVal vData As String)
            mvarPRName27 = vData
        End Set
    End Property

    Public Property PRDesg26 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PRDesg26 = mvarPRDesg26
        End Get
        Set(ByVal vData As String)
            mvarPRDesg26 = vData
        End Set
    End Property
    Public Property PRName26 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PRName26 = mvarPRName26
        End Get
        Set(ByVal vData As String)
            mvarPRName26 = vData
        End Set
    End Property

    Public Property CoStatus() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoStatus = mvarCoStatus
        End Get
        Set(ByVal vData As String)
            mvarCoStatus = vData
        End Set
    End Property
    Public Property IsCoAddChg As Boolean
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            IsCoAddChg = mvarIsCoAddChg
        End Get
        Set(ByVal vData As Boolean)
            mvarIsCoAddChg = vData
        End Set
    End Property

    Public Property CoPAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoPAN = mvarCoPAN
        End Get
        Set(ByVal vData As String)
            mvarCoPAN = vData
        End Set
    End Property

    Public Property CoTAN As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoTAN = mvarCoTAN
        End Get
        Set(ByVal vData As String)
            mvarCoTAN = vData
        End Set
    End Property

    Public Property CoPin As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoPin = mvarCoPin
        End Get
        Set(ByVal vData As String)
            mvarCoPin = vData
        End Set
    End Property

    Public Property CoStateID As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoStateID = mvarCoStateID
        End Get
        Set(ByVal vData As Integer)
            mvarCoStateID = vData
        End Set
    End Property

    Public Property CoAdd5 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoAdd5 = mvarCoAdd5
        End Get
        Set(ByVal vData As String)
            mvarCoAdd5 = vData
        End Set
    End Property

    Public Property CoAdd4 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoAdd4 = mvarCoAdd4
        End Get
        Set(ByVal vData As String)
            mvarCoAdd4 = vData
        End Set
    End Property

    Public Property CoAdd3 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoAdd3 = mvarCoAdd3
        End Get
        Set(ByVal vData As String)
            mvarCoAdd3 = vData
        End Set
    End Property

    Public Property CoAdd2 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoAdd2 = mvarCoAdd2
        End Get
        Set(ByVal vData As String)
            mvarCoAdd2 = vData
        End Set
    End Property

    Public Property CoAdd1 As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoAdd1 = mvarCoAdd1
        End Get
        Set(ByVal vData As String)
            mvarCoAdd1 = vData
        End Set
    End Property

    Public Property CoName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            CoName = mvarCoName
        End Get
        Set(ByVal vData As String)
            mvarCoName = vData
        End Set
    End Property

    Public Property coid As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            coid = mvarCoID
        End Get
        Set(ByVal vData As Integer)
            mvarCoID = vData
        End Set
    End Property

    Public Property GovtStateID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            GovtStateID = mvarGovtStateID
        End Get
        Set(ByVal vData As String)
            mvarGovtStateID = vData
        End Set
    End Property

    Public Property PAOCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PAOCode = mvarPAOCode
        End Get
        Set(ByVal vData As String)
            mvarPAOCode = vData
        End Set
    End Property

    Public Property DDOCode As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            DDOCode = mvarDDOCode
        End Get
        Set(ByVal vData As String)
            mvarDDOCode = vData
        End Set
    End Property

    Public Property MinistryID As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            MinistryID = mvarMinistryID
        End Get
        Set(ByVal vData As String)
            mvarMinistryID = vData
        End Set
    End Property

    Public Property MinistryName As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            MinistryName = mvarMinistryName
        End Get
        Set(ByVal vData As String)
            mvarMinistryName = vData
        End Set
    End Property

    Public Property PAORegNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            PAORegNo = mvarPAORegNo
        End Get
        Set(ByVal vData As String)
            mvarPAORegNo = vData
        End Set
    End Property

    Public Property DDORegNo As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.MinorHead
            DDORegNo = mvarDDORegNo
        End Get
        Set(ByVal vData As String)
            mvarDDORegNo = vData
        End Set
    End Property

End Class