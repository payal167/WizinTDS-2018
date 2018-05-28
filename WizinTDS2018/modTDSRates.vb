Module modTDSRates
    Public AY As String, ToDate As Date, FromDate As Date
    Public ToDateQ As Date
    Public FromDateQ As Date
    Public InstallPath As String
    Public TDSRates(21) As Rates, FY As String
    Public lastrr As String
    Public lastret As String
    Public Structure Rates
        Public Section As String
        Public RateNonCompany As Double
        Public RateCompany As Double
    End Structure
    Public Sub SetTDSRates()
        Dim dt As Date
        dt = "01/April/2017"
        FromDate = Format(dt, "dd/MMM/yyyy")
        ToDate = Format(CDate("31/March/2018"), "dd/MMM/yyyy")
        FY = CStr(Year(FromDate) & "-" & Year(ToDate))
        AY = CStr(Year(FromDate) + 1 & "-" & Year(ToDate) + 1)
        'InstallPath = Registry_Read("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\WizinTDS" & Year(ToDate) & ".exe\", "")
        'InstallPath = "C:\JAKINFO\WizinTDS " & Year(ToDate)
        If InstallPath = "" Then
            InstallPath = "C:\JAKINFO\WizinTDS " & Year(ToDate)
        Else
            InstallPath = Left(InstallPath, Len(InstallPath) - 17)
        End If
        'InstallPath = IIf(InstallPath = "", "C:\JAKINFO\WizinTDS " & Year(ToDate), Left(InstallPath, Len(InstallPath) - 17))
        With TDSRates(0)
            .Section = "192"
            .RateCompany = 0.0#
            .RateNonCompany = 0.0#
        End With
        With TDSRates(1)
            .Section = "193"
            .RateCompany = 21.0#
            .RateNonCompany = 10.5
        End With
        With TDSRates(2)
            .Section = "194"
            .RateCompany = 10.5
            .RateNonCompany = 10.5
        End With
        With TDSRates(3)
            .Section = "194A"
            .RateCompany = 10
            .RateNonCompany = 10
        End With
        With TDSRates(4)
            .Section = "194B"
            .RateCompany = 30
            .RateNonCompany = 30
        End With
        With TDSRates(5)
            .Section = "194BB"
            .RateCompany = 33.66
            .RateNonCompany = 31.5
        End With
        With TDSRates(6)
            .Section = "194C"
            .RateCompany = 2
            .RateNonCompany = 1
        End With
        With TDSRates(7)
            .Section = "194D"
            .RateCompany = 10
            .RateNonCompany = 10
        End With
        With TDSRates(8)
            .Section = "194E"
            .RateCompany = 10.5
            .RateNonCompany = 10.5
        End With
        With TDSRates(9)
            .Section = "194EE"
            .RateCompany = 0.0#
            .RateNonCompany = 21.0#
        End With
        With TDSRates(10)
            .Section = "194F"
            .RateCompany = 0.0#
            .RateNonCompany = 21.0#
        End With
        With TDSRates(11)
            .Section = "194G"
            .RateCompany = 10.5
            .RateNonCompany = 10.5
        End With
        With TDSRates(12)
            .Section = "194H"
            .RateCompany = 10
            .RateNonCompany = 10
        End With
        With TDSRates(13)
            .Section = "194Ia"
            .RateCompany = 10
            .RateNonCompany = 10
        End With
        With TDSRates(14)
            .Section = "194J"
            .RateCompany = 10
            .RateNonCompany = 10
        End With
        With TDSRates(15)
            .Section = "194K"
            .RateCompany = 10.5
            .RateNonCompany = 10.5
        End With
        With TDSRates(16)
            .Section = "195"
            .RateCompany = 0.0#
            .RateNonCompany = 0.0#
        End With
        With TDSRates(17)
            .Section = "196A"
            .RateCompany = 0.0#
            .RateNonCompany = 0.0#
        End With
        With TDSRates(18)
            .Section = "196B"
            .RateCompany = 0.0#
            .RateNonCompany = 0.0#
        End With
        With TDSRates(19)
            .Section = "196C"
            .RateCompany = 0.0#
            .RateNonCompany = 0.0#
        End With
        With TDSRates(20)
            .Section = "196D"
            .RateCompany = 0.0#
            .RateNonCompany = 0.0#
        End With
        With TDSRates(21)
            .Section = "194Ib"
            .RateCompany = 2.0#
            .RateNonCompany = 0.0#
        End With

    End Sub

    Public Function GetTDSRates(ByVal Section As String) As Rates
        Dim i As Long
        For i = 0 To UBound(TDSRates)
            If TDSRates(i).Section = Section Then
                GetTDSRates = TDSRates(i)
                Exit For
            End If
        Next i
        Return GetTDSRates
    End Function

End Module