Imports System.IO
Imports System.Reflection

Module mod24QConvert
    '    Dim fso As New FileSystemObject
    Dim eTDSPath As String, eFileName As String
    Dim oCoMst As New clsCoMst ', TStrm As TextStream
    Dim TStrm As StreamWriter
    Dim hrp As String
    Dim Value18 As Date
    Public Sub Convert24Q(frmid As String, filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstChln, rstRetn, rstCSum, rstDeductee, rstDSum, rstXD, rstC, rstD, rst16SDSumTDS, rst16SDSumOthInc, rst16SDSum80CCG, rst16SDSum6A, rstDSalSum, rst16SDSumAllow, rst16SDSum80CCF, rstChlnDed, rst16SDSum80C, rstSD, rst16SD, rst16SDSum, rstSDSum As New DataSet
        Dim LNo As Long, RecNo As Long, RecNoDed As Long
        Dim SumC As Double, SumD As Double
        Dim FH(18), BH(70), CD(41), DD(43), SD(67), SD16(8), SD6A(8)
        Dim strFH As String, strBH As String, strCD As String, strDD As String
        Dim i As Long ', fs As New FileSystemObject
        Dim ChlnTotalinDeductee As Double

        'Check whether the return exists..
        rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = '" & frmid & "'")
        If rstRetn.Tables(0).Rows.Count <= 0 Then
            Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
            'GoTo 
            Exit Sub
        End If

        '    End If

        oCoMst = oCoMst.FetchCo(selectedcoid)
        'Start the conversion
        'Open the text file..
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        ' eFileName = eTDSPath & "\" & frmid & "-" & oCoMst.CoName & ".TXT"
        eFileName = eTDSPath & "\" & oCoMst.CoName & "\F" & frmid & "V" & ".txt"
        'Do the conversion
        ' Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        TStrm = File.CreateText(eFileName)
        'Get the related data..
        ' rstC = FetchDataSet("SELECT * FROM Challan24Q WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))   'Challan Records
        rstC = FetchDataSet("SELECT * FROM Challan24Q WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID").ToString()) ', Cnn      'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM Deductee24Q AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("RetnID")) 'Deductee Records

        rstSD = FetchDataSet("SELECT SD.*, DM.*,fx.* " _
        & " FROM (SalaryDetail24Q AS SD INNER JOIN DeductMst AS DM ON SD.DId = DM.Did) LEFT JOIN F16XtraDetails AS fx ON ((sd.RetnID = fx.RetnID) and (sd.DId = fx.DId))" _
        & " WHERE sd.RetnId=" & rstRetn.Tables(0).Rows(0)("RetnID"))  'Salary Detail Records


        rst16SD = FetchDataSet("select sd.*, dm.*, fx.RentExceeds, fx.LandLord1PAN, fx.LandLord1Name, fx.LandLord2PAN, fx.LandLord2Name, fx.LandLord3PAN, fx.LandLord3Name, fx.LandLord4PAN, fx.LandLord4Name, fx.InttPaidOnHP, fx.Lender1PAN, fx.Lender1Name, fx.Lender2PAN, fx.Lender2Name, fx.Lender3PAN, fx.Lender3Name, fx.Lender4PAN, fx.Lender4Name, fx.HasSAFundPaid, fx.FundName, fx.DateFrom, fx.DateTo, fx.AmtRepaid, fx.AvgRate, fx.TaxDedAmt, fx.GrossTotIncome" _
& " FROM DeductMst as dm INNER JOIN (Form16Details as sd LEFT JOIN F16XtraDetails as fx ON (sd.RetnID = fx.RetnID) AND (sd.DId = fx.DId)) ON dm.DId = sd.DId" _
& " WHERE sd.RetnId=" & rstRetn.Tables(0).Rows(0)("RetnID"))


        rstDSalSum = FetchDataSet("SELECT sum(AmtOfPayment) as TotalSal FROM Deductee24Q AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))
        rstCSum = FetchDataSet("SELECT sum(TaxAmt) as TotAmt,sum(Surcharge) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth, Sum(Afees) as TotFees FROM Challan24Q WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))     'Challan Records
        Dim SDGTIAmt As Double
        If oCoMst.UseForm16 = True Then
            '    'use form 16 details table to collect data...

            rst16SDSum = FetchDataSet("SELECT sum(Gross1+Gross2+Gross3+TotalSalaryPreEmp) as TotSal,sum(sec16ii) as Tot16ii, sum(sec16iii) as Tot16iii" &
      " FROM Form16Details WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))     'Form 16 Detail Records7

            rst16SDSumOthInc = FetchDataSet("SELECT Sum([F16MD].[GrossAmt]) AS TotalOthInc" &
      " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
      " WHERE (((F16D.RetnID)=" & rstRetn.Tables(0).Rows(0)("RetnID") & ") AND ((F16MD.TypeOfDetail)='O'))")
            rst16SDSumAllow = FetchDataSet("SELECT Sum([F16MD].[GrossAmt]) AS TotalAllow" &
      " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
      " WHERE (((F16D.RetnID)=" & rstRetn.Tables(0).Rows(0)("RetnID") & ") AND ((F16MD.TypeOfDetail)='A'))")
            'Calculate Gross Total Income from Salary Detail records....applicable in Q4 only
            SDGTIAmt = IIf(String.IsNullOrEmpty(rst16SDSum.Tables(0).Rows(0)("totsal").ToString()), 0, Val(rst16SDSum.Tables(0).Rows(0)("totsal").ToString())) -
                       IIf(String.IsNullOrEmpty(rst16SDSum.Tables(0).Rows(0)("Tot16ii").ToString()), 0, Val(rst16SDSum.Tables(0).Rows(0)("Tot16ii").ToString())) -
                       IIf(String.IsNullOrEmpty(rst16SDSum.Tables(0).Rows(0)("Tot16iii").ToString()), 0, Val(rst16SDSum.Tables(0).Rows(0)("Tot16iii").ToString()))
            SDGTIAmt = SDGTIAmt - IIf(String.IsNullOrEmpty(rst16SDSumAllow.Tables(0).Rows(0)("TotalAllow").ToString()), 0, Val(rst16SDSumAllow.Tables(0).Rows(0)("TotalAllow").ToString()))
            SDGTIAmt = SDGTIAmt + IIf(String.IsNullOrEmpty(rst16SDSumOthInc.Tables(0).Rows(0)("TotalOthInc").ToString()), 0, Val(rst16SDSumOthInc.Tables(0).Rows(0)("TotalOthInc").ToString()))
            'SDGTIAmt = Val(rst16SDSum.Tables(0).Rows(0)("totsal")) -
            '           Val(rst16SDSum.Tables(0).Rows(0)("Tot16ii")) -
            '           Val(rst16SDSum.Tables(0).Rows(0)("Tot16iii"))
            'SDGTIAmt = SDGTIAmt - Val(rst16SDSumAllow.Tables(0).Rows(0)("TotalAllow"))
            'SDGTIAmt = SDGTIAmt + Val(rst16SDSumAllow.Tables(0).Rows(0)("TotalOthInc"))
        Else
            'use summary details table to collect data...
            rstSDSum = FetchDataSet("SELECT sum(TotalSalary) as TotSal,sum(sec16ii) as Tot16ii, sum(sec16iii) as Tot16iii," &
       " Sum(OtherIncome) as TotOI, Sum(TotalSalaryPreEmp) as TotSalPreEmp FROM SalaryDetail24Q WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))    'Salary Detail Records
            'Calculate Gross Total Income from Salary Detail records....applicable in Q4 only
            SDGTIAmt = IIf(String.IsNullOrEmpty(rstSDSum.Tables(0).Rows(0)("totsal").ToString()), 0, rstSDSum.Tables(0).Rows(0)("totsal")) -
        IIf(String.IsNullOrEmpty(rstSDSum.Tables(0).Rows(0)("Tot16ii").ToString()), 0, Val(rstSDSum.Tables(0).Rows(0)("Tot16ii").ToString()) -
        IIf(String.IsNullOrEmpty(rstSDSum.Tables(0).Rows(0)("Tot16iii").ToString()), 0, Val(rstSDSum.Tables(0).Rows(0)("Tot16iii").ToString()) +
        IIf(String.IsNullOrEmpty(rstSDSum.Tables(0).Rows(0)("totOI").ToString()), 0, Val(rstSDSum.Tables(0).Rows(0)("totOI").ToString()) +
        IIf(String.IsNullOrEmpty(rstSDSum.Tables(0).Rows(0)("TotSalPreEmp").ToString()), 0, Val(rstSDSum.Tables(0).Rows(0)("TotSalPreEmp").ToString())))))
            'SDGTIAmt = Val(rstSDSum.Tables(0).Rows(0)("totsal")) -
            '           Val(rstSDSum.Tables(0).Rows(0)("Tot16ii")) -
            '           Val(rstSDSum.Tables(0).Rows(0)("Tot16iii")) +
            '           Val(rstSDSum.Tables(0).Rows(0)("totOI")) +
            '           Val(rstSDSum.Tables(0).Rows(0)("TotSalPreEmp"))
        End If
        Dim rstXDCnt As New DataSet
        rstXD = FetchDataSet("SELECT * FROM F16XtraDetails WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID"))

        'If oCoMst.chkpan = True Then
        If AllPANVerified(rstD, Left(filename, Len(filename) - 5) & "_PAN_VERIFY.CSV") = False Then
            If MsgBox("There are unverified PANs in the return." & vbCrLf &
        "Do you want to take risk of converting the TDS file?", vbYesNo + vbDefaultButton2, "UNVERIFIED PAN EXISTS") = vbNo Then
                IsAllPANVerified = False
                TStrm.Close()
                Call OpenNotePad(Left(filename, Len(filename) - 5) & "_PAN_VERIFY.CSV")
                'GoTo cleanup
            Else
                IsAllPANVerified = True
            End If
        Else
            IsAllPANVerified = True
        End If


        'End If
        'START WRITING THE TEXT FILE NOW.
        '*****************************************************************************************
        'FILE HEADER RECORD...
        FH(1) = 1
        FH(2) = "FH"
        FH(3) = "SL1"
        FH(4) = "R"
        FH(5) = Format(Today(), "ddMMyyyy")
        FH(6) = 1
        FH(7) = "D"
        FH(8) = UCase(oCoMst.CoTAN)
        FH(9) = 1
        Dim versionNumber As String
        Dim a
        versionNumber = Assembly.GetExecutingAssembly.FullName
        versionNumber = Strings.Left(versionNumber, 29)
        a = Split(versionNumber, ",")

        FH(10) = a(0) + " " + Strings.Right(a(1), 7)

        FH(11) = vbNullString
        FH(12) = vbNullString
        FH(13) = vbNullString
        FH(14) = vbNullString
        FH(15) = vbNullString
        FH(16) = vbNullString
        FH(17) = vbNullString
        FH(18) = vbNullString   'Added on 28/01/2013 for new FVU 3.3
        strFH = vbNullString
        For i = 1 To UBound(FH)
            strFH = strFH & FH(i)
            If i <= UBound(FH) - 1 Then strFH = strFH & "^"
        Next i
        TStrm.WriteLine(strFH)
        '*****************************************************************************************
        'BATCH HEADER RECORD...
        'Calculate Challan Total, and Deductee Total
        Dim ChallanTotal As Double
        ChallanTotal = Val(rstCSum.Tables(0).Rows(0)("TotAmt").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totsc").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totecess").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totint").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("tototh").ToString()) +
                Val(rstCSum.Tables(0).Rows(0)("totfees").ToString())
        If IsDBNull(ChallanTotal) = True Then
            SumC = 0
        Else
            SumC = CDbl(ChallanTotal)
        End If
        BH(1) = 2
        BH(2) = "BH"
        BH(3) = 1
        BH(4) = rstC.Tables(0).Rows.Count
        BH(5) = "24Q"
        BH(6) = vbNullString
        BH(7) = vbNullString
        BH(8) = vbNullString
        BH(9) = lastrr 'IIf(IsNull(lastrr), "", vbNullString)
        BH(10) = vbNullString
        BH(11) = vbNullString
        BH(12) = vbNullString
        BH(13) = UCase(oCoMst.CoTAN)
        BH(14) = vbNullString
        BH(15) = IIf(Len(Trim(oCoMst.CoPAN)) = 0, "PANNOTREQD", UCase(oCoMst.CoPAN))    'Changed on 03/10/09, as per new format ver. 4.0
        BH(16) = Left(AY, 4) & Right(AY, 2)     'eg. AY format 200607
        BH(17) = Left(FY, 4) & Right(FY, 2)     'eg. FY format 200506
        BH(18) = "Q" & Right(frmid, 1)
        BH(19) = oCoMst.CoName
        BH(20) = oCoMst.CoBrDiv
        BH(21) = oCoMst.CoAdd1
        BH(22) = oCoMst.CoAdd2
        BH(23) = oCoMst.CoAdd3
        BH(24) = oCoMst.CoAdd4
        BH(25) = oCoMst.CoAdd5
        BH(26) = oCoMst.CoStateID
        BH(27) = oCoMst.CoPin
        BH(28) = oCoMst.CoEmail
        BH(29) = oCoMst.CoStd
        BH(30) = oCoMst.CoPhone
        BH(31) = IIf(oCoMst.IsCoAddChg = True, "Y", "N")
        BH(32) = oCoMst.CoStatus
        BH(33) = oCoMst.PRName24
        BH(34) = oCoMst.PRDesg24
        BH(35) = oCoMst.PR24Add1
        BH(36) = oCoMst.PR24Add2
        BH(37) = oCoMst.PR24Add3
        BH(38) = oCoMst.PR24Add4
        BH(39) = oCoMst.PR24Add5
        BH(40) = oCoMst.PR24StateID
        BH(41) = oCoMst.PR24Pin
        BH(42) = oCoMst.PR24Email
        BH(43) = oCoMst.Comobile  'vbNullString   'Mobile No - not given in this release
        BH(44) = oCoMst.PR24Std
        BH(45) = oCoMst.PR24Phone
        BH(46) = IIf(oCoMst.IsPR24AddChg = True, "Y", "N")
        BH(47) = Format(ChallanTotal, "0.00")
        BH(48) = vbNullString
        BH(49) = IIf(oCoMst.UseForm16, rst16SD.Tables(0).Rows.Count, rstSD.Tables(0).Rows.Count) 'changed by nitin on 16/06/06 as per new structure for Q4
        BH(50) = IIf(SDGTIAmt = 0, vbNullString, Format(SDGTIAmt, "0.00")) 'changed by harsha on 02/02/15 as per new structure for Q4
        BH(51) = "N"
        BH(52) = lastret 'IIf(IsNull(lastrr), "", vbNullString) 'vbNullString change for fvu 4.0
        BH(53) = vbNullString
        BH(54) = IIf(oCoMst.GovtStateID = "", "", Format(oCoMst.GovtStateID, "00"))
        BH(55) = oCoMst.PAOCode
        BH(56) = oCoMst.DDOCode
        BH(57) = IIf(oCoMst.MinistryID = "", "", Format(oCoMst.MinistryID, "00"))
        BH(58) = IIf(Val(BH(57)) = 99, oCoMst.MinistryName, vbNullString)
        BH(59) = Trim(oCoMst.PR24PAN)               'Trim(oCoMst.TANRegNo)
        BH(60) = IIf(oCoMst.PAORegNo = "", "", Format(oCoMst.PAORegNo, "00"))
        BH(61) = oCoMst.DDORegNo
        BH(62) = oCoMst.CoSTDAlt
        BH(63) = oCoMst.CoPhoneAlt
        BH(64) = oCoMst.CoEmailAlt
        BH(65) = oCoMst.PR24STDAlt
        BH(66) = oCoMst.PR24PhoneAlt
        BH(67) = oCoMst.PR24EmailAlt
        BH(68) = oCoMst.AIN
        BH(69) = oCoMst.gstin
        BH(70) = vbNullString

        strBH = vbNullString
        For i = 1 To UBound(BH)
            strBH = strBH & BH(i)
            If i <= UBound(BH) - 1 Then strBH = strBH & "^"
        Next i
        TStrm.WriteLine(strBH)
        '*****************************************************************************************
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.movefirst
        End If
        LNo = 3 : RecNo = 1
        Dim CTotal As Long
        'Do While Not rstC.Tables(0).Rows.Count
        For c = 0 To rstC.Tables(0).Rows.Count - 1
            rstChlnDed = FetchDataSet("SELECT DT.*, DM.* FROM Deductee24Q AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("RetnID") &
      " AND ChallanID=" & rstC.Tables(0).Rows(c)("ChallanID"))  'Deductee Records
            'fees added on 17/05/14
            'CTotal = CLng(IIf(IsDBNull(rstC.Tables(0).Rows(0)("TaxAmt")), 0, rstC.Tables(0).Rows(0)("TaxAmt"))) +
            '         CLng(IIf(IsDBNull(rstC.Tables(0).Rows(0)("Surcharge")), 0, rstC.Tables(0).Rows(0)("Surcharge"))) +
            '         CLng(IIf(IsDBNull(rstC.Tables(0).Rows(0)("ECess")), 0, rstC.Tables(0).Rows(0)("ECess"))) +
            '         CLng(IIf(IsDBNull(rstC.Tables(0).Rows(0)("Interest")), 0, rstC.Tables(0).Rows(0)("Interest"))) +
            '         CLng(IIf(IsDBNull(rstC.Tables(0).Rows(0)("Others")), 0, rstC.Tables(0).Rows(0)("Others"))) +
            '         CLng(IIf(IsDBNull(rstC.Tables(0).Rows(0)("AFees")), 0, rstC.Tables(0).Rows(0)("AFees")))
            CTotal = Val(rstC.Tables(0).Rows(0)("TaxAmt").ToString()) +
                     Val(rstC.Tables(0).Rows(0)("Surcharge").ToString()) +
                     Val(rstC.Tables(0).Rows(0)("ECess").ToString()) +
                     Val(rstC.Tables(0).Rows(0)("Interest").ToString()) +
                     Val(rstC.Tables(0).Rows(0)("Others").ToString()) +
                     Val(rstC.Tables(0).Rows(0)("AFees").ToString())
            CD(1) = LNo
            CD(2) = "CD"
            CD(3) = BH(3)
            CD(4) = RecNo
            CD(5) = rstChlnDed.Tables(0).Rows.Count
            CD(6) = IIf(CTotal <= 0, "Y", "N")
            CD(7) = vbNullString
            CD(8) = vbNullString
            CD(9) = vbNullString
            CD(10) = vbNullString
            CD(11) = vbNullString
            CD(12) = IIf((CD(6) = "Y" Or (oCoMst.CoStatus = "C" And rstC.Tables(0).Rows(c)("IsBookEntry").ToString() = True)), vbNullString, rstC.Tables(0).Rows(c)("BankChallanNo").ToString())
            'CD(12) = IIf((CD(6) = "Y" Or (oCoMst.CoStatus = "C" And rstC.Tables(0).Rows(0)("IsBookEntry") = True)), vbNullString, rstC.Tables(0).Rows(0)("BankChallanNo"))
            CD(13) = vbNullString
            'changed on 03/10/09 as per new format ver 4.0
            If BH(32) = "A" Or BH(32) = "S" Or BH(32) = "D" Or BH(32) = "E" Or BH(32) = "G" Or
        BH(32) = "H" Or BH(32) = "L" Or BH(32) = "N" Then
                CD(14) = rstC.Tables(0).Rows(c)("TranVouNo")
            Else
                CD(14) = vbNullString
            End If
            'CD(14) = IIf(oCoMst.CoStatus = "C" And rstC!IsBookEntry = True, rstC!BankChallanNo, vbNullString)
            CD(15) = vbNullString
            CD(16) = IIf(IsDBNull(rstC.Tables(0).Rows(c)("BankBrCode")), vbNullString, Format(Val(rstC.Tables(0).Rows(c)("BankBrCode").ToString()), "0000000"))
            CD(17) = vbNullString

            If CD(6) = "Y" Then
                Select Case Right(frmid, 1)
                    Case 1
                        Value18 = "30/06/" & Left(FY, 4)
                    Case 2
                        Value18 = "30/09/" & Left(FY, 4)
                    Case 3
                        Value18 = "31/12/" & Left(FY, 4)
                    Case 4
                        Value18 = "31/03/" & Right(FY, 4)
                End Select
            Else
                Value18 = rstC.Tables(0).Rows(c)("DtOfChallan").ToString()
            End If
            CD(18) = Format(Value18, "ddMMyyyy")
            CD(19) = vbNullString
            CD(20) = vbNullString
            'CD(21) = IIf(BH(32) = "A" Or BH(32) = "S" Or BH(32) = "D" Or BH(32) = "E" Or BH(32) = "G" Or BH(32) = "H" Or BH(32) = "L" Or BH(32) = "N", "92A", "92B")
            CD(21) = vbNullString 'changed for AY 13-14
            CD(22) = Format(Val(rstC.Tables(0).Rows(c)("TaxAmt").ToString()), "0.00") 'IIf(IsNull(rstC!TaxAmt), 0, rstC!TaxAmt), "0.00")
            CD(23) = Format(Val(rstC.Tables(0).Rows(c)("Surcharge").ToString()), "0.00")
            CD(24) = Format(Val(rstC.Tables(0).Rows(c)("ECess").ToString()), "0.00")
            CD(25) = Format(Val(rstC.Tables(0).Rows(c)("Interest").ToString()), "0.00")
            CD(26) = Format(Val(rstC.Tables(0).Rows(c)("Others").ToString()), "0.00")
            CD(39) = Format(Val(rstC.Tables(0).Rows(c)("AFees").ToString()), "0.00")
            CD(27) = Format((Val(CD(22)) + Val(CD(23)) + Val(CD(24)) + Val(CD(25)) + Val(CD(26)) + Val(CD(39))), "0.00")
            CD(28) = vbNullString

            'If rstDSum.State = adStateOpen Then 
            rstDSum.Dispose()
            rstDSum = FetchDataSet("SELECT sum(TotalTaxDeposited) as TotDep, sum(TaxAmt) as TotTax, " &
       "sum(Surcharge) as TotSur, sum(Ecess) as TotEcess, sum(0) as TotInt, " &
       "sum(0) as TotOth FROM Deductee24Q WHERE ChallanID=" & rstC.Tables(0).Rows(c)("ChallanID"))
            'CD(29) = Format(IIf(IsDBNull(rstDSum.Tables(0).Rows(0)("totdep")), 0, rstDSum.Tables(0).Rows(0)("totdep")), "0.00")
            'CD(30) = Format(IIf(IsDBNull(rstDSum.Tables(0).Rows(0)("tottax")), 0, rstDSum.Tables(0).Rows(0)("tottax")), "0.00")
            'CD(31) = Format(IIf(IsDBNull(rstDSum.Tables(0).Rows(0)("totsur")), 0, rstDSum.Tables(0).Rows(0)("totsur")), "0.00")
            'CD(32) = Format(IIf(IsDBNull(rstDSum.Tables(0).Rows(0)("totecess")), 0, rstDSum.Tables(0).Rows(0)("totecess")), "0.00")
            CD(29) = Format(Val(rstDSum.Tables(0).Rows(0)("totdep").ToString()), "0.00")
            CD(30) = Format(Val(rstDSum.Tables(0).Rows(0)("tottax").ToString()), "0.00")
            CD(31) = Format(Val(rstDSum.Tables(0).Rows(0)("totsur").ToString()), "0.00")
            CD(32) = Format(Val(rstDSum.Tables(0).Rows(0)("totecess").ToString()), "0.00")
            CD(33) = Format(Val(CD(30)) + Val(CD(31)) + Val(CD(32)), "0.00")
            CD(34) = Format(Val(rstC.Tables(0).Rows(c)("AInterest").ToString()), "0.00")
            CD(35) = Format(Val(rstC.Tables(0).Rows(c)("AOthers").ToString()), "0.00")
            ' CD(34) = Format(IIf(IsDBNull(rstC.Tables(0).Rows(0)("AInterest")), 0, rstC.Tables(0).Rows(0)("AInterest")), "0.00")
            ' CD(35) = Format(IIf(IsDBNull(rstC.Tables(0).Rows(0)("AOthers")), 0, rstC.Tables(0).Rows(0)("AOthers")), "0.00")
            CD(36) = vbNullString     'rstC!ChqDDNo
            CD(37) = IIf(rstC.Tables(0).Rows(c)("IsBookEntry"), "Y", IIf(Val(CD(27)) = 0, "", "N"))  'Value changed to NULL in case of NIL challan FUV3.9
            CD(38) = vbNullString
            'CD(39) - already mentioned above for total tax calc purpose...
            CD(40) = IIf(Val(CD(27)) = 0 Or CD(37) = "Y", "", rstC.Tables(0).Rows(c)("MinorHead") & "")
            CD(41) = vbNullString
            strCD = ""
            For i = 1 To UBound(CD)
                strCD = strCD & CD(i)
                If i <= UBound(CD) - 1 Then strCD = strCD & "^"
            Next i
            TStrm.WriteLine(strCD)
            LNo = LNo + 1
            'Insert Deductee Details for this challan
            '*****************************************************************************************
            'DEDUCTEE DETAIL RECORD
            RecNoDed = 1      'Line no not reset as it will continue from challan detail...
            'open the recordset
            If rstChlnDed.Tables(0).Rows.Count > 0 Then
                'rstChlnDed.MoveFirst
            End If
            ' Do While Not rstChlnDed.Tables(0).Rows.Count
            For k = 0 To rstChlnDed.Tables(0).Rows.Count - 1
                DD(1) = LNo
                DD(2) = "DD"
                DD(3) = BH(3)
                DD(4) = CD(4)
                DD(5) = RecNoDed
                DD(6) = "O"
                DD(7) = LNo            'changed by nitin on 16/06/06 for Q4 compulsion. vbNullString vbNullString    'TO BE CHANGED.....****************** NOW
                DD(8) = vbNullString
                DD(9) = vbNullString
                DD(10) = UCase(rstChlnDed.Tables(0).Rows(k)("DPan").ToString())
                DD(11) = vbNullString
                ' changed to null as per file format ver 4.1
                '        DD(12) = IIf(rstChlnDed!DPANCat <> 0, rstChlnDed!DPanRef, vbNullString)
                DD(12) = vbNullString
                'DD(13) = rstChlnDed.Tables(0).Rows(k)("DName")
                DD(13) = IIf(String.IsNullOrEmpty(rstChlnDed.Tables(0).Rows(k)("DName").ToString()), "", rstChlnDed.Tables(0).Rows(k)("DName").ToString())
                DD(14) = Format(Val(rstChlnDed.Tables(0).Rows(k)("TaxAmt").ToString()), "0.00")
                DD(15) = Format(Val(rstChlnDed.Tables(0).Rows(k)("Surcharge").ToString()), "0.00")
                DD(16) = Format(Val(rstChlnDed.Tables(0).Rows(k)("ECess").ToString()), "0.00")
                DD(17) = Format(Val(DD(14)) + Val(DD(15)) + Val(DD(16)), "0.00")
                DD(18) = vbNullString
                DD(19) = Format(Val(rstChlnDed.Tables(0).Rows(k)("TotalTaxDeposited").ToString()), "0.00")
                DD(20) = vbNullString
                DD(21) = vbNullString
                DD(22) = Format(Val(rstChlnDed.Tables(0).Rows(k)("AmtOfPayment").ToString()), "0.00")
                Dim dt As Date
                dt = rstChlnDed.Tables(0).Rows(k)("DtOfPayment")
                DD(23) = Format(dt, "ddMMyyyy")
                If rstChlnDed.Tables(0).Rows(k)("DtOfDeduction").ToString = "" Then
                    DD(24) = vbNullString
                Else

                    dt = rstChlnDed.Tables(0).Rows(k)("DtOfDeduction")
                    DD(24) = Format(dt, "ddMMyyyy")
                End If
                'DD(23) = Format(rstChlnDed.Tables(0).Rows(k)("DtOfPayment").ToString(), "ddMMyyyy")
                'DD(24) = Format(rstChlnDed.Tables(0).Rows(k)("DtOfDeduction").ToString(), "ddMMyyyy")
                dt = rstC.Tables(0).Rows(c)("DtOfChallan")
                DD(25) = Format(dt, "ddMMyyyy")
                ' DD(25) = Format(rstChlnDed.Tables(0).Rows(k)("DtOfChallan").ToString(), "ddMMyyyy")
                DD(26) = vbNullString
                DD(27) = vbNullString
                DD(28) = vbNullString
                DD(29) = vbNullString
                DD(30) = Trim(rstChlnDed.Tables(0).Rows(k)("Remark").ToString())
                DD(31) = vbNullString
                DD(32) = vbNullString
                DD(33) = Right(rstChlnDed.Tables(0).Rows(k)("Sec"), 3)
                DD(34) = rstChlnDed.Tables(0).Rows(k)("CertNo")
                DD(35) = vbNullString
                DD(36) = vbNullString
                DD(37) = vbNullString
                DD(38) = vbNullString
                DD(39) = vbNullString
                DD(40) = vbNullString
                DD(41) = vbNullString
                DD(42) = vbNullString
                DD(43) = vbNullString

                RecNoDed = RecNoDed + 1
                If RecNoDed > 999999999 Then RecNoDed = 1
                strDD = ""
                For i = 1 To UBound(DD)
                    strDD = strDD & DD(i)
                    If i <= UBound(DD) - 1 Then strDD = strDD & "^"
                Next i
                TStrm.WriteLine(strDD)
                'warn if PAN Available and remark = "C"
                If IsValidPAN(rstChlnDed.Tables(0).Rows(k)("DPan").ToString()) = 0 Then
                    If DD(30) = "C" Then
                        If MsgBox("WARNING: Please remove remark 'C' of Deductee:" & DD(13) & " and then convert the file before filing. " & vbCrLf &
                    "Do You want to Continue", vbYesNo, "WARNING") = vbNo Then
                            CPan = True
                        End If
                    End If
                End If
                LNo = LNo + 1
                'rstChlnDed.MoveNext
            Next k
            'If rstChlnDed.State = adStateOpen Then 
            ' rstChlnDed.Dispose()
            ' rstC.MoveNext
            RecNo = RecNo + 1

        Next c
        'Inserts Salary Detail Records...**********************************
        'Changed by Nitin Betharia on 12/05/2008 for picking up data from new database using form 16 details or summary
        'details...
        Dim Count16 As Integer, Val16ii As Double, val16iii As Double
        Dim Count6A As Integer, Val80CCE As Double, Val80CCF As Double, Val80CCG As Double, ValOtherVIA As Double
        Dim strSD As String, strsd16 As String, strsd6A As String
        If oCoMst.UseForm16 = False Then
            'Insert SD records using the summary tables...
            If rstSD.Tables(0).Rows.Count > 0 Then
                'rstSD.Tables(0).Rows(0)(0)
            End If
            RecNo = 1
            For j = 0 To rstSD.Tables(0).Rows.Count - 1
                SD(1) = LNo
                SD(2) = "SD"
                SD(3) = BH(3)
                SD(4) = RecNo
                SD(5) = "A"
                SD(6) = vbNullString
                SD(7) = UCase(rstSD.Tables(0).Rows(j)("DPan").ToString())
                SD(8) = rstSD.Tables(0).Rows(j)("DPANCat").ToString() '<> 0)', rstSD!Dpanref, vbNullString)
                SD(9) = IIf(String.IsNullOrEmpty(rstSD.Tables(0).Rows(j)("DName").ToString()), "", rstSD.Tables(0).Rows(j)("DName").ToString())
                SD(10) = rstSD.Tables(0).Rows(j)("Category").ToString() & ""
                SD(11) = Format(rstSD.Tables(0).Rows(j)("EmpFromDt").ToString(), "ddMMyyyy")
                SD(12) = Format(rstSD.Tables(0).Rows(j)("EmpToDt").ToString(), "ddMMyyyy")
                ' sd 34 and sd 35 defined here as it is used by sd 13
                SD(34) = Format(Val(rstSD.Tables(0).Rows(j)("TotalSalary").ToString()), "0.00")
                SD(35) = Format(Val(rstSD.Tables(0).Rows(j)("TotalSalaryPreEmp").ToString()), "0.00")
                SD(13) = Format(Val(SD(34)) + Val(SD(35)), "0.00")
                SD(14) = vbNullString
                Count16 = 0 : Val16ii = 0 : val16iii = 0
                Val16ii = rstSD.Tables(0).Rows(j)("Sec16ii").ToString() 'IIf(IsNull(rstSD!), 0, Val(rstSD!Sec16ii))
                val16iii = rstSD.Tables(0).Rows(j)("Sec16iii").ToString() 'IIf(IsNull(rstSD!), 0, Val(rstSD!Sec16iii))
                If Val16ii > 0 Then
                    Count16 = Count16 + 1
                End If
                If val16iii > 0 Then
                    Count16 = Count16 + 1
                End If
                SD(15) = Count16
                SD(16) = Format(Val16ii + val16iii, "0.00")
                SD(17) = Format(Val(SD(13)) - Val(SD(16)), "0.00")
                SD(18) = Format(rstSD.Tables(0).Rows(j)("OtherIncome").ToString()) ', 0, rstSD!OtherIncome), "0.00")
                SD(19) = Format(Val(SD(17)) + Val(SD(18)), "0.00")
                SD(20) = vbNullString
                Count6A = 0 : Val80CCE = 0 : ValOtherVIA = 0 : Val80CCF = 0 : Val80CCG = 0
                Val80CCE = Format(rstSD.Tables(0).Rows(j)("Sec80CCEAmt").ToString()) ', 0, Val(rstSD!)), "0.00")
                ValOtherVIA = Format(rstSD.Tables(0).Rows(j)("OtherVIA").ToString()) ', 0, Val(rstSD!OtherVIA)), "0.00")
                Val80CCF = Format(rstSD.Tables(0).Rows(j)("Sec80CCFAmt").ToString()) ', 0#, rstSD!Sec80CCFAmt), "0.00")
                Val80CCG = Format(rstSD.Tables(0).Rows(j)("Sec80CCGAmt").ToString()) ', 0#, rstSD!Sec80CCGAmt), "0.00")
                If Val80CCE > 0 Then
                    Count6A = Count6A + 1
                End If
                If Val80CCF > 0 Then
                    Count6A = Count6A + 1
                End If
                If Val80CCG > 0 Then
                    Count6A = Count6A + 1
                End If
                If ValOtherVIA > 0 Then
                    Count6A = Count6A + 1
                End If
                SD(21) = Count6A
                SD(22) = Format(Val80CCE + Val80CCF + Val80CCG + ValOtherVIA, "0.00")
                SD(23) = Format(SD(19) - SD(22), "0.00")
                SD(24) = Format(Val(rstSD.Tables(0).Rows(j)("TaxAmt").ToString()), "0.00")
                SD(25) = Format(Val(rstSD.Tables(0).Rows(j)("Surcharge").ToString()), "0.00")
                SD(26) = Format(Val(rstSD.Tables(0).Rows(j)("ECess").ToString()), "0.00")
                SD(27) = Format(Val(rstSD.Tables(0).Rows(j)("Relief89").ToString()), "0.00")
                SD(28) = Format(Val(SD(24)) + Val(SD(25)) + Val(SD(26)) - Val(SD(27)), "0.00")
                'sd 36 and sd 37 defined here as it is used by sd 29
                SD(36) = Format(Val(rstSD.Tables(0).Rows(j)("TDSAmt").ToString()), "0.00")
                SD(37) = Format(Val(rstSD.Tables(0).Rows(j)("TDSAmtPreEmp").ToString()), "0.00")
                SD(29) = Format(Val(SD(36)) + Val(SD(37)), "0.00")
                SD(30) = Format(Val(SD(28)) - Val(SD(29)), "0.00")
                SD(31) = vbNullString
                SD(32) = vbNullString
                SD(33) = vbNullString
                ' sd 34, to 37 defined above.... sd 13 and sd 29
                hrp = ""
                If (Left(FY, 4) >= 2013) Then
                    hrp = IIf(rstSD.Tables(0).Rows(j)("HighRatePAN").ToString() = True, "Y", "N")
                Else
                    hrp = ""
                End If
                SD(38) = hrp

                Dim cnt1, cnt2 As Integer
                SD(39) = IIf(rstSD.Tables(0).Rows(j)("RENTEXCEEDS") = True, "Y", "N")

                If SD(39) = "Y" Then
                    SD(41) = rstSD.Tables(0).Rows(j)("LandLord1PAN").ToString()
                    SD(42) = rstSD.Tables(0).Rows(j)("LandLord1Name").ToString()
                    SD(43) = rstSD.Tables(0).Rows(j)("LANDLORD2PAN").ToString()
                    SD(44) = rstSD.Tables(0).Rows(j)("LANDLORD2NAME").ToString()
                    SD(45) = rstSD.Tables(0).Rows(j)("LANDLORD3PAN").ToString()
                    SD(46) = rstSD.Tables(0).Rows(j)("LANDLORD3NAME").ToString()
                    SD(47) = rstSD.Tables(0).Rows(j)("LANDLORD4PAN").ToString()
                    SD(48) = rstSD.Tables(0).Rows(j)("LandLord4Name").ToString()
                    SD(40) = 0
                    If Not (rstSD.Tables(0).Rows(j)("LandLord1PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("LandLord1PAN").ToString() = "") Then
                        SD(40) = 1
                    End If
                    If Not (rstSD.Tables(0).Rows(j)("LANDLORD2PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("LANDLORD2PAN").ToString() = "") Then
                        SD(40) = 2
                    End If
                    If Not (rstSD.Tables(0).Rows(j)("LANDLORD3PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("LANDLORD3PAN").ToString() = "") Then
                        SD(40) = 3
                    End If
                    If Not (rstSD.Tables(0).Rows(j)("LANDLORD4PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("LANDLORD4PAN").ToString() = "") Then
                        SD(40) = 4
                    End If
                Else
                    SD(41) = vbNullString
                    SD(42) = vbNullString
                    SD(43) = vbNullString
                    SD(44) = vbNullString
                    SD(45) = vbNullString
                    SD(46) = vbNullString
                    SD(47) = vbNullString
                    SD(48) = vbNullString
                    SD(40) = 0
                End If

                SD(49) = IIf(rstSD.Tables(0).Rows(j)("InttPaidOnHP") = True, "Y", "N")
                If SD(49) = "Y" Then
                    SD(51) = rstSD.Tables(0).Rows(j)("Lender1PAN").ToString()
                    SD(52) = rstSD.Tables(0).Rows(j)("Lender1Name").ToString()
                    SD(53) = rstSD.Tables(0).Rows(j)("Lender2PAN").ToString()
                    SD(54) = rstSD.Tables(0).Rows(j)("Lender2Name").ToString()
                    SD(55) = rstSD.Tables(0).Rows(j)("Lender3PAN").ToString()
                    SD(56) = rstSD.Tables(0).Rows(j)("Lender3Name").ToString()
                    SD(57) = rstSD.Tables(0).Rows(j)("Lender4PAN").ToString()
                    SD(58) = rstSD.Tables(0).Rows(j)("Lender4Name").ToString()
                    SD(50) = 0
                    If Not (rstSD.Tables(0).Rows(j)("Lender1PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("Lender1PAN").ToString() = "") Then
                        SD(50) = 1
                    End If
                    If Not (rstSD.Tables(0).Rows(j)("Lender2PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("Lender2PAN").ToString() = "") Then
                        SD(50) = 2
                    End If
                    If Not (rstSD.Tables(0).Rows(j)("Lender3PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("Lender3PAN").ToString() = "") Then
                        SD(50) = 3
                    End If
                    If Not (rstSD.Tables(0).Rows(j)("Lender4PAN").ToString()) And Not (rstSD.Tables(0).Rows(j)("Lender4PAN").ToString() = "") Then
                        SD(50) = 4
                    End If
                Else
                    SD(51) = vbNullString
                    SD(52) = vbNullString
                    SD(53) = vbNullString
                    SD(54) = vbNullString
                    SD(55) = vbNullString
                    SD(56) = vbNullString
                    SD(57) = vbNullString
                    SD(58) = vbNullString
                    SD(50) = 0
                End If
                SD(59) = IIf(rstSD.Tables(0).Rows(j)("HasSAFundPaid") = True, "Y", "N")

                If SD(59) = "Y" Then
                    SD(60) = rstSD.Tables(0).Rows(j)("FundName")
                    SD(61) = Format(rstSD.Tables(0).Rows(j)("DATEFROM"), "ddMMyyyy")
                    SD(62) = Format(rstSD.Tables(0).Rows(j)("DATETO"), "ddMMyyyy")
                    SD(63) = rstSD.Tables(0).Rows(j)("AMTREPAID")

                    If rstSD.Tables(0).Rows(j)("AvgRate") = 0 Or IsDBNull(rstSD.Tables(0).Rows(j)("AvgRate")) Then
                        SD(64) = vbNullString
                    Else
                        SD(64) = Format(Val(rstSD.Tables(0).Rows(0)("AvgRate")), "0.0000")
                    End If
                    SD(65) = Val(rstSD.Tables(0).Rows(0)("TAXDEDAMT"))
                    SD(66) = SD(19) + SD(63)
                Else
                    SD(60) = vbNullString
                    SD(61) = vbNullString
                    SD(62) = vbNullString
                    SD(63) = vbNullString
                    SD(64) = vbNullString
                    SD(65) = vbNullString
                    SD(66) = vbNullString
                End If
                SD(67) = vbNullString
                strSD = ""
                For i = 1 To UBound(SD)
                    strSD = strSD & SD(i)
                    If i <= UBound(SD) - 1 Then strSD = strSD & "^"
                Next i
                TStrm.WriteLine(strSD)
                LNo = LNo + 1
                'Insert Section 16 Details for aforesaid Salary Detail record
                '*****************************************************************************************
                'SEC 16 DETAIL RECORD
                RecNoDed = 1      'Line no not reset as it will continue from salary detail...
                strsd16 = ""
                'write data for section 16(ii)
                If Val16ii > 0 Then
                    SD16(1) = LNo
                    SD16(2) = "S16"
                    SD16(3) = SD(3)
                    SD16(4) = SD(4)
                    SD16(5) = RecNoDed
                    SD16(6) = "16(ii)"
                    SD16(7) = Format(Val16ii, "0.00")
                    SD16(8) = vbNullString
                    RecNoDed = RecNoDed + 1
                    strsd16 = ""
                    For i = 1 To UBound(SD16)
                        strsd16 = strsd16 & SD16(i)
                        If i <= UBound(SD16) - 1 Then strsd16 = strsd16 & "^"
                    Next i
                    TStrm.WriteLine(strsd16)
                    LNo = LNo + 1
                End If
                'write data for section 16(ii1)
                If val16iii > 0 Then
                    SD16(1) = LNo
                    SD16(2) = "S16"
                    SD16(3) = SD(3)
                    SD16(4) = SD(4)
                    SD16(5) = RecNoDed
                    SD16(6) = "16(iii)"
                    SD16(7) = Format(val16iii, "0.00")
                    SD16(8) = vbNullString
                    RecNoDed = RecNoDed + 1
                    strsd16 = ""
                    For i = 1 To UBound(SD16)
                        strsd16 = strsd16 & SD16(i)
                        If i <= UBound(SD16) - 1 Then strsd16 = strsd16 & "^"
                    Next i
                    TStrm.WriteLine(strsd16)
                    LNo = LNo + 1
                End If
                'Insert Chapter 6A Details for aforesaid Salary Detail record
                '*****************************************************************************************
                'CHP 6A DETAIL RECORD
                RecNoDed = 1      'Line no not reset as it will continue from salary detail...
                strsd6A = ""
                'write data for section 80CCE
                If Val80CCE > 0 Then
                    SD6A(1) = LNo
                    SD6A(2) = "C6A"
                    SD6A(3) = SD(3)
                    SD6A(4) = SD(4)
                    SD6A(5) = RecNoDed
                    SD6A(6) = "80CCE"
                    SD6A(7) = Format(Val80CCE, "0.00")
                    SD6A(8) = vbNullString
                    RecNoDed = RecNoDed + 1
                    strsd6A = ""
                    For i = 1 To UBound(SD6A)
                        strsd6A = strsd6A & SD6A(i)
                        If i <= UBound(SD6A) - 1 Then strsd6A = strsd6A & "^"
                    Next i
                    TStrm.WriteLine(strsd6A)
                    LNo = LNo + 1
                End If
                If Val80CCF > 0 Then
                    SD6A(1) = LNo
                    SD6A(2) = "C6A"
                    SD6A(3) = SD(3)
                    SD6A(4) = SD(4)
                    SD6A(5) = RecNoDed
                    SD6A(6) = "80CCF"
                    SD6A(7) = Format(Val80CCF, "0.00")
                    SD6A(8) = vbNullString
                    RecNoDed = RecNoDed + 1
                    strsd6A = ""
                    For i = 1 To UBound(SD6A)
                        strsd6A = strsd6A & SD6A(i)
                        If i <= UBound(SD6A) - 1 Then strsd6A = strsd6A & "^"
                    Next i
                    TStrm.WriteLine(strsd6A)
                    LNo = LNo + 1
                End If
                If Val80CCG > 0 Then
                    SD6A(1) = LNo
                    SD6A(2) = "C6A"
                    SD6A(3) = SD(3)
                    SD6A(4) = SD(4)
                    SD6A(5) = RecNoDed
                    SD6A(6) = "80CCG"
                    SD6A(7) = Format(Val80CCG, "0.00")
                    SD6A(8) = vbNullString
                    RecNoDed = RecNoDed + 1
                    strsd6A = ""
                    For i = 1 To UBound(SD6A)
                        strsd6A = strsd6A & SD6A(i)
                        If i <= UBound(SD6A) - 1 Then strsd6A = strsd6A & "^"
                    Next i
                    TStrm.WriteLine(strsd6A)
                    LNo = LNo + 1
                End If
                'write data for Other VIA Details
                If ValOtherVIA > 0 Then
                    SD6A(1) = LNo
                    SD6A(2) = "C6A"
                    SD6A(3) = SD(3)
                    SD6A(4) = SD(4)
                    SD6A(5) = RecNoDed
                    SD6A(6) = "OTHERS"
                    SD6A(7) = Format(ValOtherVIA, "0.00")
                    SD6A(8) = vbNullString
                    RecNoDed = RecNoDed + 1
                    strsd6A = ""
                    For i = 1 To UBound(SD6A)
                        strsd6A = strsd6A & SD6A(i)
                        If i <= UBound(SD6A) - 1 Then strsd6A = strsd6A & "^"
                    Next i
                    TStrm.WriteLine(strsd6A)
                    LNo = LNo + 1
                End If
                'move to next SD record..
                'rstSD.MoveNext
                RecNo = RecNo + 1
            Next
        Else            'Use Form 16 is ticked on the company...fill in the SD records using details table...
            'Insert SD records using the form 16 Details table.....
            If rst16SD.Tables(0).Rows.Count > 0 Then
                'rst16SD.MoveFirst
            End If
            RecNo = 1
            Dim s As Integer
            For s = 0 To rst16SD.Tables(0).Rows.Count - 1
                'If rst16SDSumAllow.State = adStateOpen Then 
                rst16SDSumAllow.Dispose()
                'If rst16SDSum80C.State = adStateOpen Then
                rst16SDSum80C.Dispose()
                'If rst16SDSum80CCF.State = adStateOpen Then 
                rst16SDSum80CCF.Dispose()
                'If rst16SDSum80CCG.State = adStateOpen Then 
                rst16SDSum80CCG.Dispose()
                'If rst16SDSum6A.State = adStateOpen Then 
                rst16SDSum6A.Dispose()
                'If rst16SDSumOthInc.State = adStateOpen Then 
                rst16SDSumOthInc.Dispose()
                'If rst16SDSumTDS.State = adStateOpen Then 
                rst16SDSumTDS.Dispose()

                rst16SDSumAllow = FetchDataSet("SELECT Sum([F16MD].[GrossAmt]) AS TotalAllow" &
            " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
            " WHERE F16D.f16id = " & rst16SD.Tables(0).Rows(s)("F16ID") & " AND F16MD.TypeOfDetail='A'")
                rst16SDSumOthInc = FetchDataSet("SELECT Sum([F16MD].[GrossAmt]) AS TotalOthInc" &
                " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
                " WHERE (((F16D.f16id)=" & rst16SD.Tables(0).Rows(s)("F16ID") & ") AND ((F16MD.TypeOfDetail)='O'))")
                rst16SDSum80C = FetchDataSet("SELECT Sum([F16MD].[DeductibleAmt]) AS Total80C" &
                " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
                " WHERE (((F16D.f16id)=" & rst16SD.Tables(0).Rows(s)("F16ID") & ") AND ((F16MD.TypeOfDetail)='E'))")
                rst16SDSum80CCF = FetchDataSet("SELECT Sum([F16MD].[DeductibleAmt]) AS Total80CCF" &
                " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
                " WHERE (((F16D.f16id)=" & rst16SD.Tables(0).Rows(s)("F16ID") & ") AND ((F16MD.TypeOfDetail)='F'))")
                rst16SDSum80CCG = FetchDataSet("SELECT Sum([F16MD].[DeductibleAmt]) AS Total80CCG" &
                " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
                " WHERE (((F16D.f16id)=" & rst16SD.Tables(0).Rows(s)("F16ID") & ") AND ((F16MD.TypeOfDetail)='G'))")
                rst16SDSum6A = FetchDataSet("SELECT Sum([F16MD].[DeductibleAmt]) AS Total6A" &
                " FROM Form16Details AS F16D INNER JOIN Form16MoreDetails AS F16MD ON F16D.F16ID = F16MD.F16ID" &
                " WHERE (((F16D.f16id)=" & rst16SD.Tables(0).Rows(s)("F16ID") & ") AND ((F16MD.TypeOfDetail)='V'))")

                rst16SDSumTDS = FetchDataSet("SELECT Sum([c24].[TaxAmt]) AS SumTax, Sum([c24].[Surcharge]) AS SumSur, Sum([c24].[ECess]) AS SumECess FROM Form16Details AS D24 INNER JOIN F16Challan AS C24 ON D24.F16ID = C24.F16ID WHERE D24.DId= " & rst16SD.Tables(0).Rows(s)("SD.DId").Value & " Union All" &
                       " SELECT sum(D24.TaxAmt) as SumTax, sum(D24.Surcharge) as SumSur, Sum(D24.ECess)as SumECess  FROM Challan24Q AS C24 INNER JOIN Deductee24Q AS D24 ON  C24.ChallanID = D24.ChallanId WHERE D24.DId = " & rst16SD.Tables(0).Rows(s)("SD.DId").Value & "")

                SD(1) = LNo
                SD(2) = "SD"
                SD(3) = BH(3)
                SD(4) = RecNo
                SD(5) = "A"
                SD(6) = vbNullString
                SD(7) = UCase(rst16SD.Tables(0).Rows(s)("DPan").ToString())
                SD(8) = IIf(rst16SD.Tables(0).Rows(s)("DPANCat") <> 0, rst16SD.Tables(0).Rows(s)("Dpanref"), vbNullString)
                SD(9) = IIf(String.IsNullOrEmpty(rst16SD.Tables(0).Rows(s)("DName").ToString()), "", rst16SD.Tables(0).Rows(s)("DName").ToString())
                SD(10) = rst16SD.Tables(0).Rows(s)("Category").ToString() & ""
                SD(11) = Format(rst16SD.Tables(0).Rows(s)("EmpFromDt"), "ddMMyyyy")
                SD(12) = Format(rst16SD.Tables(0).Rows(s)("EmpToDt"), "ddMMyyyy")
                'SD 34 AND 35 Defined earlier as it is used by sd(13)
                SD(34) = Format(Val(rst16SD.Tables(0).Rows(s)("F16ID"))) + 'IIf(IsNull(rst16SD!Gross1), 0, rst16SD!Gross1) +
                        Val(rst16SD.Tables(0).Rows(s)("Gross2")) +
                        Val(rst16SD.Tables(0).Rows(s)("Gross3")) -
                       Val(rst16SDSumAllow.Tables(0).Rows(s)("TotalAllow")) ', 0, rst16SDSumAllow!TotalAllow), "0.00")
                SD(35) = Format(Val(rst16SD.Tables(0).Rows(s)("TotalSalaryPreEmp").ToString()), "0.00")
                SD(13) = Format(Val(SD(34)) + Val(SD(35)), "0.00")
                SD(14) = vbNullString
                Count16 = 0 : Val16ii = 0 : val16iii = 0
                Val16ii = 'IIf(IsNull(rst16SD!Sec16ii), 0, Val(rst16SD!Sec16ii))
                    val16iii = Val(rst16SD.Tables(0).Rows(s)("Sec16iii")) ', 0, Val(rst16SD!Sec16iii))
                If Val16ii > 0 Then
                    Count16 = Count16 + 1
                End If
                If val16iii > 0 Then
                    Count16 = Count16 + 1
                End If
                SD(15) = Count16
                SD(16) = Format(Val16ii + val16iii, "0.00")
                SD(17) = Format(Val(SD(13)) - Val(SD(16)), "0.00")
                SD(18) = Format(Val(rst16SDSumOthInc.Tables(0).Rows(0)("TotalOthInc").ToString()), "0.00")
                SD(19) = Format(Val(SD(17)) + Val(SD(18)), "0.00")
                SD(20) = vbNullString
                Count6A = 0 : Val80CCE = 0 : ValOtherVIA = 0
                ' Val80CCE = Format(IIf(IsNull(rst16SDSum80C!Total80C), 0, Val(rst16SDSum80C!Total80C)), "0.00")
                ' Val80CCF = Format(IIf(IsNull(rst16SDSum80CCF!Total80CCF), 0, rst16SDSum80CCF!Total80CCF), "0.00")
            Next

        End If
        TStrm.Close()
    End Sub


End Module
