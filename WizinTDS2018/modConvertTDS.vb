Module modConvertTDS
    ' Dim fs As New FileSystemObject
    Dim eTDSPath As String, eFileName As String
    Dim oCoMst As New clsCoMst
    ', TStrm As TextStream

    Public Sub Convert2Txt(frmid As String, filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstC, rstD, rstRetn, rstCSum, rstDSum As New DataSet
        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        If frmid <> "F26" Then
            MsgBox("Wrong Parameter, Call JAK Infosolutions P Ltd", vbCritical)
            GoTo cleanup
        Else
            'Parameter ok, check if return exists..
            rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = " & IIf(frmid = "F26", 26, 27))
            If rstRetn.Tables(0).Rows.Count <= 0 Then
                Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
                GoTo cleanup
            End If
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        eFileName = eTDSPath & "\" & frmid & "-" & oCoMst.CoName & ".TXT"
        'Do the conversion
        If System.IO.File.Exists(eFileName) = False Then
            System.IO.File.Create(eFileName).Dispose()
        End If
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        ' TStrm = fs.CreateTextFile(filename, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())     'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM DeducteeTDS AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("retnID")) 'Deductee Records
        rstCSum = FetchDataSet("SELECT sum(Amt) as TotAmt,sum(Surcharges) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())      'Challan Records
        rstDSum = FetchDataSet("SELECT sum(AmtOfTDS) as TotD FROM DeducteeTDS WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString()) 'Deductee Records
        'START WRITING THE TEXT FILE NOW.
        Dim txtTANAppNo As Double
        'If oCoMst.CoTAN = "TANAPPLIED" Then
        '   txtTANAppNo = Format(InputBox("Please Enter your TAN Application No", "TANAPPLIED Number", 0), "00000000000000")
        'End If
        'FILE HEADER RECORD...COMMON FOR F26 AND F27
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            ' FILE TYPE CHANGED FROM XNS TO NS3 AS PER NEW FORMAT DT 22/4/2005
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            ' FILE TYPE CHANGED FROM ENS TO NS3 AS PER NEW FORMAT DT 22/4/2005
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        If frmid = "F26" Then
            Dim ChallanTotal As Double
            ChallanTotal = (IIf((rstCSum.Tables(0).Rows(0)("TotAmt").ToString()), 0, rstRetn.Tables(0).Rows(0)("retnID").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totsc").ToString()), 0, (rstCSum.Tables(0).Rows(0)("totsc").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totecess").ToString()), 0, rstCSum.Tables(0).Rows(0)("totecess").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totint").ToString()), 0, rstCSum.Tables(0).Rows(0)("totint").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("tototh").ToString()), 0, rstCSum.Tables(0).Rows(0)("tototh").ToString())))
            '    ChallanTotal = (rstCSum!TotAmt + rstCSum!totsc + rstCSum!Totecess + rstCSum!TotInt + rstCSum!TotOth)
            If String.IsNullOrEmpty(ChallanTotal) = True Then
                SumC = 0
            Else
                SumC = CDbl(ChallanTotal * 100)
            End If
            If rstDSum.Tables(0).Rows(0)("totd").ToString() = True Then
                SumD = 0
            Else
                SumD = CDbl(rstDSum.Tables(0).Rows(0)("totd").ToString() * 100)
            End If

            'BATCH HEADER RECORD.

            TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
   Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format("26", "!@@@@") & Space(8) &
   oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
   Format(oCoMst.CoName, "!" & New String("0", 75)) & Format(oCoMst.CoAdd1, "!" & New String("0", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("0", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("0", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("0", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("0", 25)) &
   Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
   IIf(oCoMst.IsCoAddChg = True, "Y", "N") & oCoMst.CoStatus &
   Format("Y", "!@@") & Format(oCoMst.PRName26, "!" & New String("0", 75)) &
   Format(oCoMst.PRDesg26, "!" & New String("0", 20)) &
   Format(SumC, New String("0", 14)) &
   Format(SumD, New String("0", 14)) & Format(0, New String("0", 14) &
   Space(10) & Format(txtTANAppNo, New String("0", 14)) &
   Format(oldRRRNo, New String("0", 14))))   '' revised return number to be incrop...
        ElseIf frmid = "F27" Then
        End If
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Dim CTotal As Long
        Do While Not rstC.Tables(0).Rows.Count
            CTotal = IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString()) + CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges)").ToString()) + CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()))) * 100
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") & Format(rstC.Tables(0).Rows(0)("sec").ToString(), "!@@@@@") &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()) * 100), New String("0", 14)) &
      Format(CTotal, New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), 0, rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      IIf(rstC.Tables(0).Rows(0)("IsBookEntry").ToString() = True, "Y", "N") &
      Space(1))      'extra space added as per new format, this space is not there in new form no 24

            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstD.Tables(0).Rows.Count
            ' new filler added after pin no as per new format dt 22/04/2005, isbookentry added..
            '
            TStrm.WriteLine(Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(rstD.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(IIf(rstD.Tables(0).Rows(0)("DType").ToString() = "O", 2, 1), "00") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(rstD.Tables(0).Rows(0)("DAdd1").ToString(), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd2").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd2").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd2").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd3").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd3").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd3").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd4").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd4").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd4").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd5").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd5").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd5").ToString())), "!" & New String("@", 25)) &
      Format(rstD.Tables(0).Rows(0)("DState").ToString(), "00") & Format(rstD.Tables(0).Rows(0)("DPin").ToString(), "000000") &
      New String("0", 14) &
      Format(rstD.Tables(0).Rows(0)("AmtOfPay").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("DtOfPay").ToString(), "ddMMyyyy") &
      IIf(rstD.Tables(0).Rows(0)("IsBookEntry").ToString() = True, "B", "C") &
      Format(IIf(rstD.Tables(0).Rows(0)("RateOfTDS").ToString() >= 100, 0, rstD.Tables(0).Rows(0)("RateOfTDS").ToString()) * 100, "0000") & Space(1) &
      Format(rstD.Tables(0).Rows(0)("AmtOfTDS").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstD.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstD.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstD.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstD.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("CertificateDt").ToString()), Space(8), rstD.Tables(0).Rows(0)("CertificateDt").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("Reason").ToString()), Space(1), IIf(rstD.Tables(0).Rows(0)("Reason").ToString() = vbNullString, Space(1), rstD.Tables(0).Rows(0)("Reason").ToString())), "@") &
      New String("0", 14))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then RecNo = 1
            'rstD.MoveNext
        Loop
        TStrm.Close()
cleanup:
        'CLOSE THE FILE..

        'AND THE CONNECTIONS ALSO..
        'If rstC.State = adStateOpen Then 
        rstC.Dispose()
        'If rstD.State = adStateOpen Then 
        rstD.Dispose()
        'If rstRetn.State = adStateOpen Then 
        rstRetn.Dispose()
        'If rstCSum.State = adStateOpen Then 
        rstCSum.Dispose()
        'If rstDSum.State = adStateOpen Then 
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        filename = Nothing
        oCoMst = Nothing

    End Sub

    Public Sub Convert27ETxt(frmid As String, filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstC, rstRetn, rstCSum, rstDSum, rstD As New DataSet

        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        If frmid <> "F27E" Then
            MsgBox("Wrong Parameter, Call JAK Infosolutions P Ltd", vbCritical)
            GoTo cleanup
        Else
            'Parameter ok, check if return exists..
            rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = 31")
            If rstRetn.Tables(0).Rows.Count <= 0 Then
                Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
                GoTo cleanup
            End If
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        eFileName = eTDSPath & "\" & frmid & "-" & oCoMst.CoName & ".TXT"
        'Do the conversion
        ' TStrm = fs.CreateTextFile(filename, True)
        If System.IO.File.Exists(eFileName) = False Then
            System.IO.File.Create(eFileName).Dispose()
        End If
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())    'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM Deductee27E AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("retnID"))  'Deductee Records
        rstCSum = FetchDataSet("SELECT sum(Amt) as TotAmt,sum(Surcharges) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())    'Challan Records
        rstDSum = FetchDataSet("SELECT sum(AmtOfTDS) as TotD FROM Deductee27E WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())  'Deductee Records
        'START WRITING THE TEXT FILE NOW.
        Dim txtTANAppNo As Double
        'If oCoMst.CoTAN = "TANAPPLIED" Then
        '   txtTANAppNo = Format(InputBox("Please Enter your TAN Application No", "TANAPPLIED Number", 0), "00000000000000")
        'End If
        'FILE HEADER RECORD...COMMON FOR F26 AND F27
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            ' FILE TYPE CHANGED FROM XNS TO NS3 AS PER NEW FORMAT DT 22/4/2005
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            ' FILE TYPE CHANGED FROM ENS TO NS3 AS PER NEW FORMAT DT 22/4/2005
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            TStrm.WriteLine("000000001FHNS3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        If frmid = "F27E" Then
            Dim ChallanTotal As Double
            ChallanTotal = (IIf((rstCSum.Tables(0).Rows(0)("TotAmt").ToString()), 0, rstCSum.Tables(0).Rows(0)("TotAmt").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totsc").ToString()), 0, rstCSum.Tables(0).Rows(0)("totsc").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totecess").ToString()), 0, rstCSum.Tables(0).Rows(0)("totecess").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("totint").ToString()), 0, rstCSum.Tables(0).Rows(0)("totint").ToString()) +
                    IIf((rstCSum.Tables(0).Rows(0)("tototh").ToString()), 0, rstCSum.Tables(0).Rows(0)("tototh").ToString()))
            '    ChallanTotal = (rstCSumTables(0).Rows(0)("TotAmt + rstCSumTables(0).Rows(0)("totsc + rstCSumTables(0).Rows(0)("Totecess + rstCSumTables(0).Rows(0)("TotInt + rstCSumTables(0).Rows(0)("TotOth)
            If String.IsNullOrEmpty(ChallanTotal) = True Then
                SumC = 0
            Else
                SumC = CDbl(ChallanTotal * 100)
            End If
            If (rstDSum.Tables(0).Rows(0)("totd").ToString()) = True Then
                SumD = 0
            Else
                SumD = CDbl(rstDSum.Tables(0).Rows(0)("totd") * 100)
            End If
            'BATCH HEADER RECORD.

            TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
   Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format("27E", "!@@@@") & Space(8) &
   oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
   Format(oCoMst.CoName, "!" & New String("@", 75)) & Format(oCoMst.CoAdd1, "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("@", 25)) &
   Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
   IIf(oCoMst.IsCoAddChg = True, "Y", "N") & oCoMst.CoStatus &
   Format("Y", "!@@") & Format(oCoMst.PRName26, "!" & New String("@", 14)) &
   Format(oCoMst.PRDesg26, "!" & New String("@", 20)) &
   Format(SumC, New String("@", 14)) &
   Format(SumD, New String("@", 14)) & Format(0, New String("@", 14)) &
   Space(10) & Format(txtTANAppNo, New String("@", 14)) &
   Format(oldRRRNo, New String("0", 14)))   '' revised return number to be incrop...
        ElseIf frmid = "F27" Then
        End If
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            ' rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Dim CTotal As Long
        Do While Not rstC.Tables(0).Rows.Count
            CTotal = (CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()))) * 100
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") & Format(rstC.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()) * 100), New String("0", 14)) &
      Format(CTotal, New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), 0, rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      IIf(rstC.Tables(0).Rows(0)("IsBookEntry").ToString() = True, "Y", "N") & rstC.Tables(0).Rows(0)("CollCode").ToString())
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstD.Tables(0).Rows.Count
            ' new filler added after pin no as per new format dt 22/04/2005, isbookentry added..
            '
            TStrm.WriteLine(Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(rstD.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(IIf(rstD.Tables(0).Rows(0)("DType").ToString() = "O", 2, 1), "00") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(rstD.Tables(0).Rows(0)("DAdd1").ToString(), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd2").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd2").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd2").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd3").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd3").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd3").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd4").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd4").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd4").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd5").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd5").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd5").ToString())), "!" & New String("@", 25)) &
      Format(rstD.Tables(0).Rows(0)("DState").ToString(), "00") & Format(rstD.Tables(0).Rows(0)("DPin").ToString(), "000000") &
      Format(rstD.Tables(0).Rows(0)("PurchAmt").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("AmtOfPay").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("DtOfPay").ToString(), "ddMMyyyy") & Space(1) &
      Format(IIf(rstD.Tables(0).Rows(0)("RateOfTDS").ToString() >= 100, 0, rstD.Tables(0).Rows(0)("RateOfTDS").ToString()) * 100, "0000") & Space(1) &
      Format(rstD.Tables(0).Rows(0)("AmtOfTDS").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstD.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstD.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstD.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstD.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("CertificateDt").ToString()), Space(8), rstD.Tables(0).Rows(0)("CertificateDt").ToString()), "ddMMyyyy") &
      "X" & New String("0", 14))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then RecNo = 1
            ' rstD.MoveNext
        Loop
        TStrm.Close()
cleanup:
        'CLOSE THE FILE..

        'AND THE CONNECTIONS ALSO..
        'If rstC.State = adStateOpen Then 
        rstC.Dispose()
        'If rstD.State = adStateOpen Then 
        rstD.Dispose()
        'If rstRetn.State = adStateOpen Then 
        rstRetn.Dispose()
        'If rstCSum.State = adStateOpen Then 
        rstCSum.Dispose()
        'If rstDSum.State = adStateOpen Then 
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        filename = Nothing
        oCoMst = Nothing

    End Sub

    Public Sub Convert2Txt24(filename As String, revised As String, oldRRRNo As Double, TANApplNo As Double)
        Dim rstC, rstRetn, rstD, rstP, rstDSum, rstCSum As New DataSet

        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        Dim DRec As String
        Dim Total209 As Double, Total211 As Double, Total213 As Double, Total217 As Double
        Dim Total218 As Double, Total223 As Double, Total225 As Double
        Dim Total228 As Double, Total229 As Double

        'Parameter ok, check if return exists..
        rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = 24")
        If rstRetn.Tables(0).Rows.Count <= 0 Then
            Call MsgBox("There is no data for this return form.  Kindly create" _
         & vbCrLf & "this return by using Data entry option and then" _
         & vbCrLf & "create e-TDS file using this option." _
         , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
            GoTo cleanup
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        eFileName = eTDSPath & "\" & "24-" & oCoMst.CoName & ".TXT"
        If System.IO.File.Exists(eFileName) = False Then
            System.IO.File.Create(eFileName).Dispose()
        End If
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        'TStrm = fs.CreateTextFile(filename, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID"))  'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM DeducteeSAL AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("retnID"))  'Deductee Records

        '''for checking by namrata, fetches 2 records
        'even when retn id is not matching..eg. fetch 2 records of retn
        ''id = 13 when calling for retnid=2
        Dim sql As String

        sql = "SELECT P.*, D.* FROM RetnMst AS R, PerqSAL AS P, DeductMst as D " &
      "where  P.DId = D.DId and p.retnid=r.retnid and r.RetnID=" &
     rstRetn.Tables(0).Rows(0)("retnID")
        rstP = FetchDataSet(sql)   'Preq records..

        'rstP.Open "SELECT P.*, D.* FROM RetnMst AS R INNER JOIN " & _
        '    "(PerqSAL AS P INNER JOIN DeductMst AS D ON P.DId = D.DId) ON R.RetnID = P.RetnID" & _
        '    "WHERE (((R.RetnID)=" & rstRetn!RetnId & "));", Cnn

        rstCSum = FetchDataSet("SELECT sum(Amt) as TotAmt,sum(Surcharges) as TotSc, sum(Ecess) as TotEcess," &
               " Sum(Interest) as TotInt, sum(Others) as TotOth FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID"))     'Challan Records
        rstDSum = FetchDataSet("SELECT sum(TDSAmt) as TotAmt,sum(TDSEcess)as TotEcess,sum(TDSSurcharge) as TotSC FROM DeducteeSAL WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID")) 'Deductee Records

        'START WRITING THE TEXT FILE NOW.
        'FILE HEADER RECORD
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            ' FILE TYPE CHANGED TO SL3 from XSA as per new format dated 22/4/2005 by nitin
            TStrm.WriteLine("000000001FHSL3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            ' FILE TYPE CHANGED TO SL3 from ESA as per new format dated 22/4/2005 by nitin
            TStrm.WriteLine("000000001FHSL3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            'FILE TYPE CHANGED FROM SL1 TO SL3 AS PER NEW FORMATS DT. 22/04/2005. ON 03/05/2005 BY NITIN..
            TStrm.WriteLine("000000001FHSL3" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        Dim ChallanTotal As Double
        Dim DeducteeTotal As Double
        ChallanTotal = IIf((rstCSum.Tables(0).Rows(0)("TotAmt").ToString()), 0, (rstCSum.Tables(0).Rows(0)("TotAmt").ToString())) +
                 IIf((rstCSum.Tables(0).Rows(0)("totsc").ToString()), 0, (rstCSum.Tables(0).Rows(0)("totsc").ToString())) +
                 IIf((rstCSum.Tables(0).Rows(0)("totecess").ToString()), 0, (rstCSum.Tables(0).Rows(0)("totecess").ToString())) +
                 IIf((rstCSum.Tables(0).Rows(0)("totint").ToString()), 0, (rstCSum.Tables(0).Rows(0)("totint").ToString())) +
                 IIf((rstCSum.Tables(0).Rows(0)("tototh").ToString()), 0, (rstCSum.Tables(0).Rows(0)("tototh").ToString())) +
 DeducteeTotal = IIf((rstDSum.Tables(0).Rows(0)("TotAmt").ToString()), 0, (rstDSum.Tables(0).Rows(0)("TotAmt").ToString())) +
                 IIf((rstDSum.Tables(0).Rows(0)("totsc").ToString()), 0, (rstDSum.Tables(0).Rows(0)("totsc").ToString())) +
                 IIf((rstDSum.Tables(0).Rows(0)("totecess").ToString()), 0, (rstDSum.Tables(0).Rows(0)("totecess").ToString()))
        If String.IsNullOrEmpty(ChallanTotal) = True Then
            SumC = 0
        Else
            SumC = CLng(ChallanTotal * 100)
        End If
        If String.IsNullOrEmpty(DeducteeTotal) = True Then
            SumD = 0
        Else
            SumD = CLng(DeducteeTotal * 100)
        End If

        'BATCH HEADER RECORD.
        TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format(IIf(rstP.Tables(0).Rows.Count > 999999999, "999999999", rstP.Tables(0).Rows.Count), "000000000") &
Space(8) &
oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
Format(oCoMst.CoName, "!" & New String("@", 75)) &
IIf(oCoMst.CoStatus = "O", "0000000002", "0000000001") &
Format(oCoMst.CoAdd1, "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("@", 25)) &
Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
IIf(oCoMst.IsCoAddChg = True, "Y", "N") &
Format(oCoMst.PRName24, "!" & New String("@", 75)) &
Format(oCoMst.PRDesg24, "!" & New String("@", 20)) &
Format(oCoMst.PR24Add1, "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add2), Space(25), IIf(oCoMst.PR24Add2 = vbNullString, Space(25), oCoMst.PR24Add2)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add3), Space(25), IIf(oCoMst.PR24Add3 = vbNullString, Space(25), oCoMst.PR24Add3)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add4), Space(25), IIf(oCoMst.PR24Add4 = vbNullString, Space(25), oCoMst.PR24Add4)), "!" & New String("@", 25)) &
Format(IIf(String.IsNullOrEmpty(oCoMst.PR24Add5), Space(25), IIf(oCoMst.PR24Add5 = vbNullString, Space(25), oCoMst.PR24Add5)), "!" & New String("@", 25)) &
Format(oCoMst.PR24StateID, "00") & Format(oCoMst.PR24Pin, "000000") &
IIf(oCoMst.IsPR24AddChg = True, "Y", "N") &
Format(SumC, New String("0", 14)) &
Format(SumD, New String("0", 14)) & Format(0, New String("0", 14)) & Space(10) &
Format(IIf(TANApplNo = 0, 0, TANApplNo), New String("0", 14)) & Format(IIf(oldRRRNo = 0, 0, oldRRRNo), New String("0", 14)))

        'CHALLAN DETAIL RECORD, IMPORTANT NOTE: this is different from form 26,27,
        'section field is not used here...
        Dim CTotal As Long
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Do While Not rstC.Tables(0).Rows.Count - 1
            CTotal = (CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString())) + CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()))) * 100
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("amt").ToString()), 0, rstC.Tables(0).Rows(0)("amt").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Surcharges").ToString()), 0, rstC.Tables(0).Rows(0)("Surcharges").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("ECess").ToString()), 0, rstC.Tables(0).Rows(0)("ECess").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Interest").ToString()), 0, rstC.Tables(0).Rows(0)("Interest").ToString()) * 100), New String("0", 14)) &
      Format(CLng(IIf((rstC.Tables(0).Rows(0)("Others").ToString()), 0, rstC.Tables(0).Rows(0)("Others").ToString()) * 100), New String("0", 14)) &
      Format(CTotal, New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), 0, rstC.Tables(0).Rows(0)("ChqDDNo").ToString()), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      IIf(rstC.Tables(0).Rows(0)("IsBookEntry").ToString() = True, "Y", "N"))

            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        Dim rno As Long
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        rno = rstD.Tables(0).Rows.Count
        Dim esr() As String
        ReDim Preserve esr(rno)
        Do While Not rstD.Tables(0).Rows.Count - 1

            'Do some calcuations...
            Total209 = rstD.Tables(0).Rows(0)("TotalSal").ToString() + rstD.Tables(0).Rows(0)("TotalRent").ToString() + rstD.Tables(0).Rows(0)("PFAmt").ToString()
            Total211 = Total209 - rstD.Tables(0).Rows(0)("Deduction16").ToString()
            If rstD.Tables(0).Rows(0)("OtherIncomeFlg").ToString() = "P" Then
                Total213 = Total211 + rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString()
            Else
                Total213 = Total211 - rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString()
            End If
            Total217 = rstD.Tables(0).Rows(0)("80gAmt").ToString() + rstD.Tables(0).Rows(0)("80ggAmt").ToString() + rstD.Tables(0).Rows(0)("6aamt").ToString()
            Total218 = Total213 - Total217
            Total223 = rstD.Tables(0).Rows(0)("TaxOnIncome").ToString() - (rstD.Tables(0).Rows(0)("88Rebate").ToString() + rstD.Tables(0).Rows(0)("88BRebate").ToString() + rstD.Tables(0).Rows(0)("88crebate").ToString() + IIf((rstD.Tables(0).Rows(0)("88Drebate").ToString()), 0, rstD.Tables(0).Rows(0)("88Drebate").ToString())) + rstD.Tables(0).Rows(0)("SurchargeAmt").ToString()
            Total225 = Total223 - rstD.Tables(0).Rows(0)("89Relief").ToString()
            Total228 = rstD.Tables(0).Rows(0)("TDSAmt").ToString() + rstD.Tables(0).Rows(0)("TDSSurcharge").ToString() + IIf((rstD.Tables(0).Rows(0)("TDSECess").ToString()), 0, rstD.Tables(0).Rows(0)("TDSECess").ToString())
            Total229 = Math.Abs(Total225 - Total228)

            '---
            'esr(RecNo - 1,0) = RecNo
            'esr(RecNo - 1, 1) = rstD.Tables(0).Rows(0)("DName").ToString()

            'variable drec used, because of error during desing time,
            'too many lines to concatinate.
            'while converting vb6.0 to vb.net String(75,"@") can be written as New String("@",75)

            DRec = Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(RecNo, "000000000") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) &
      Format(rstD.Tables(0).Rows(0)("FromDt").ToString(), "ddMMyyyy") & Format(rstD.Tables(0).Rows(0)("todt").ToString(), "ddMMyyyy") &
      Format(rstD.Tables(0).Rows(0)("TotalSal").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("TotalRent").ToString() * 100, New String("0", 14)) & Format(rstD.Tables(0).Rows(0)("PFAmt").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("ExPerqAmt").ToString() * 100, New String("0", 14)) & Format((Total209 * 100), New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("Deduction16").ToString() * 100, New String("0", 14)) &
      Format((Total211) * 100, New String("0", 14)) &
      Format(IIf(rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString() = 0, "P", rstD.Tables(0).Rows(0)("OtherIncomeFlg").ToString()), "@") & Format(rstD.Tables(0).Rows(0)("OtherIncomeAmt").ToString() * 100, New String("0", 13)) &
      Format((Total213) * 100, New String("0", 14)) &
      Format((rstD.Tables(0).Rows(0)("80GAmt").ToString() * 100), New String("0", 14)) & Format((rstD.Tables(0).Rows(0)("80GGAmt").ToString() * 100), New String("0", 14)) &
      Format((rstD.Tables(0).Rows(0)("6AAmt").ToString() * 100), New String("0", 14)) &
      Format((Total217) * 100, New String("0", 14)) &
      Format((Total218) * 100, New String("0", 14)) &
      Format((rstD.Tables(0).Rows(0)("TaxOnIncome ").ToString() * 100), New String("0", 14)) & Format((rstD.Tables(0).Rows(0)("88rebate").ToString() * 100), New String("0", 14)) &
      Format((rstD.Tables(0).Rows(0)("88Brebate").ToString() * 100), New String("0", 14))
            DRec = DRec & Format((rstD.Tables(0).Rows(0)("88Crebate").ToString() * 100), New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("88drebate").ToString()), 0, (rstD.Tables(0).Rows(0)("88drebate").ToString()) * 100), New String("0", 14)) &
      Format(Total223 * 100, New String("0", 14)) &
      Format((rstD.Tables(0).Rows(0)("89Relief").ToString() * 100), New String("0", 14)) &
      Format(Total225 * 100, New String("0", 14)) &
      Format((rstD.Tables(0).Rows(0)("TDSAmt").ToString() * 100), New String("0", 14)) & Format(rstD.Tables(0).Rows(0)("TDSSurcharge").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("TDSECess").ToString()), 0, rstD.Tables(0).Rows(0)("TDSECess").ToString()) * 100, New String("0", 14)) &
      Format(Total228 * 100, New String("0", 14)) &
      Format(IIf(Total229 = 0, "P", rstD.Tables(0).Rows(0)("RefundFlag").ToString()), "@") &
      Format(Total229 * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("Remark").ToString()), Space(75), rstD.Tables(0).Rows(0)("Remark").ToString()), "!" & New String("@", 75))
            TStrm.WriteLine(DRec)

            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            ' rstD.MoveNext
        Loop

        'PREQ. DETAIL RECORD
        Dim PreqVal As Double, i As Long, psr As Long
        If rstP.Tables(0).Rows.Count > 0 Then
            'rstP.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstP.Tables(0).Rows.Count
            'variable drec used, because of error during desing time,
            'too many lines to concatinate.
            For i = 0 To UBound(esr)
                If UCase(Trim((rstP.Tables(0).Rows(0)("DName").ToString())) = UCase(Trim(esr(i)))) Then
                    psr = Len(Trim(esr(i)))
                    Exit For
                End If
            Next i
            DRec = Format(LNo, "000000000") & "PD" & "000000001" & Format(RecNo, "000000000") &
      Format(rstP.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(psr, "000000000") &
      Format((rstP.Tables(0).Rows(0)("UnFurnishAmt").ToString() * 100), New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString() * 100), New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 100), New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 0.1) * 100, New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString() + ((rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString() * 0.1)) * 100), New String("0", 14)) &
      Format((rstP.Tables(0).Rows(0)("RentAmt").ToString() * 100), New String("0", 14))
            If (rstP.Tables(0).Rows(0)("UnFurnishAmt").ToString()) > 0 Then
                PreqVal = (rstP.Tables(0).Rows(0)("UnFurnishAmt").ToString()) - (rstP.Tables(0).Rows(0)("RentAmt").ToString())
            Else
                PreqVal = ((rstP.Tables(0).Rows(0)("AsIfUnFurnishAmt").ToString()) + (rstP.Tables(0).Rows(0)("CostOfFurnture").ToString() * 0.1)) - (rstP.Tables(0).Rows(0)("RentAmt").ToString())
            End If
            DRec = DRec & Format(PreqVal * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("ConveyanceAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("SalaryForPersonal").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("TravellingAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("OtherAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("EmployerPFAmt").ToString() * 100, New String("0", 14)) &
      Format(rstP.Tables(0).Rows(0)("PFInterest").ToString() * 100, New String("0", 14)) &
      Format(((PreqVal + (rstP.Tables(0).Rows(0)("ConveyanceAmt").ToString()) + (rstP.Tables(0).Rows(0)("SalaryForPersonal").ToString()) + (rstP.Tables(0).Rows(0)("TravellingAmt").ToString()) +
     (rstP.Tables(0).Rows(0)("OtherAmt").ToString()) + (rstP.Tables(0).Rows(0)("EmployerPFAmt").ToString()) + rstP.Tables(0).Rows(0)("PFInterest").ToString()) * 100), New String("0", 14))

            TStrm.WriteLine(DRec)

            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstP.MoveNext
        Loop

        'CLOSE THE FILE..
        TStrm.Close()
cleanup:

        'AND THE CONNECTIONS ALSO..
        'If rstC.State = adStateOpen Then 
        rstC.Dispose()
        'If rstD.State = adStateOpen Then 
        rstD.Dispose()
        'If rstP.State = adStateOpen Then 
        rstP.Dispose()
        'If rstRetn.State = adStateOpen Then 
        rstRetn.Dispose()
        'If rstCSum.State = adStateOpen Then 
        rstCSum.Dispose()
        'If rstDSum.State = adStateOpen Then 
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstP = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        'fs = Nothing
        oCoMst = Nothing

    End Sub
    Public Sub Convert27Txt(frmid As String, filename As String, revised As String)
        Dim rstC, rstRetn, rstCSum, rstDSum, rstD As New DataSet
        Dim LNo As Long, RecNo As Long
        Dim SumC As Double, SumD As Double
        If frmid <> "F27" And frmid <> "F28" And frmid <> "F29" And frmid <> "F30" Then
            MsgBox("Wrong Parameter, Call JAK Infosolutions P Ltd", vbCritical)
            GoTo cleanup
        Else
            'Parameter ok, check if return exists..
            rstRetn = FetchDataSet("SELECT * FROM RetnMst WHERE CoID=" & selectedcoid & " AND FrmType = " & Right(frmid, 2))
            If rstRetn.Tables(0).Rows.Count <= 0 Then
                Call MsgBox("There is no data for this return form.  Kindly create" _
            & vbCrLf & "this return by using Data entry option and then" _
            & vbCrLf & "create e-TDS file using this option." _
            , vbExclamation + vbDefaultButton1, "RETURN NOT FOUND")
                GoTo cleanup
            End If
        End If
        oCoMst = oCoMst.FetchCo(selectedcoid)
        eTDSPath = Application.StartupPath & "\e-TDS Files"
        Dim FrmNo As String
        Select Case frmid
            Case "F27"
                FrmNo = "F27-Q1"
            Case "F28"
                FrmNo = "F27-Q2"
            Case "F29"
                FrmNo = "F27-Q3"
            Case "F30"
                FrmNo = "F27-Q4"
        End Select
        eFileName = eTDSPath & "\" & FrmNo & "-" & oCoMst.CoName & ".TXT"
        'Do the conversion
        If System.IO.File.Exists(eFileName) = False Then
            System.IO.File.Create(eFileName).Dispose()
        End If
        Dim TStrm As New System.IO.StreamWriter(eFileName, True)
        'Set TStrm = fs.CreateTextFile(filename, True)
        'Get the related data..
        rstC = FetchDataSet("SELECT * FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID"))    'Challan Records
        rstD = FetchDataSet("SELECT DT.*, DM.* FROM Deductee27 AS DT " &
      "INNER JOIN DeductMst AS DM ON DT.DId = DM.DId WHERE RetnID=" &
      rstRetn.Tables(0).Rows(0)("retnID"))   'Deductee Records
        rstCSum = FetchDataSet("SELECT sum(Amt) as TotC FROM Challan WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())      'Challan Records
        rstDSum = FetchDataSet("SELECT sum(AmtOfTDS) as TotD FROM Deductee27 WHERE RetnID=" & rstRetn.Tables(0).Rows(0)("retnID").ToString())  'Deductee Records
        'START WRITING THE TEXT FILE NOW.
        'FILE HEADER RECORD...COMMON FOR F26 AND F27
        If rstD.Tables(0).Rows.Count = 0 And rstC.Tables(0).Rows.Count = 0 Then
            'Header when deductee record count is zero..
            TStrm.WriteLine("000000001FHXNS" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        ElseIf rstD.Tables(0).Rows.Count > 999999999 Or rstC.Tables(0).Rows.Count > 999999999 Then
            'Header when deductee records exceeds 999999999
            TStrm.WriteLine("000000001FHENS" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        Else
            'Normal header
            TStrm.WriteLine("000000001FHNS1" & revised & Format(Today(), ("ddMMyyyy")) & "000000001" & oCoMst.CoTAN & "000000001")
        End If
        If String.IsNullOrEmpty(rstCSum.Tables(0).Rows(0)("totc").ToString()) = True Then
            SumC = 0
        Else
            SumC = CDbl(rstCSum.Tables(0).Rows(0)("totc") * 100)
        End If
        If rstDSum.Tables(0).Rows(0)("totd").ToString() = True Then
            SumD = 0
        Else
            SumD = CDbl(rstDSum.Tables(0).Rows(0)("totd") * 100)
        End If

        'BATCH HEADER RECORD.
        TStrm.WriteLine("000000002BH000000001" & Format(IIf(rstC.Tables(0).Rows.Count > 999999999, "999999999", rstC.Tables(0).Rows.Count), "000000000") &
   Format(IIf(rstD.Tables(0).Rows.Count > 999999999, "999999999", rstD.Tables(0).Rows.Count), "000000000") & Format("27", "!@@@@") & Space(8) &
   oCoMst.CoTAN & oCoMst.CoPAN & Left(AY, 4) & Right(AY, 2) & Left(FY, 4) & Right(FY, 2) &
   Format(oCoMst.CoName, "!" & New String("@", 75)) & Format(oCoMst.CoAdd1, "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd2), Space(25), IIf(oCoMst.CoAdd2 = vbNullString, Space(25), oCoMst.CoAdd2)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd3), Space(25), IIf(oCoMst.CoAdd3 = vbNullString, Space(25), oCoMst.CoAdd3)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd4), Space(25), IIf(oCoMst.CoAdd4 = vbNullString, Space(25), oCoMst.CoAdd4)), "!" & New String("@", 25)) &
   Format(IIf(String.IsNullOrEmpty(oCoMst.CoAdd5), Space(25), IIf(oCoMst.CoAdd5 = vbNullString, Space(25), oCoMst.CoAdd5)), "!" & New String("@", 25)) &
   Format(oCoMst.CoStateID, "00") & Format(oCoMst.CoPin, "000000") &
   IIf(oCoMst.IsCoAddChg = True, "Y", "N") & oCoMst.CoStatus &
   Format(Right(FrmNo, 2), "!@@") & Format(oCoMst.PRName26, "!" & New String("@", 75)) &
   Format(oCoMst.PRDesg26, "!" & New String("@", 20)) &
   Format(SumC, New String("0", 14)) &
   Format(SumD, New String("0", 14)) & Format(0, New String("0", 14)) & Space(10) & Format(0, New String("0", 14)))
        'CHALLAN DETAIL RECORD
        If rstC.Tables(0).Rows.Count > 0 Then
            'rstC.MoveFirst
        End If
        LNo = 3 : RecNo = 1
        Do While Not rstC.Tables(0).Rows.Count - 1
            TStrm.WriteLine(Format(LNo, "000000000") & "CD" & "000000001" & Format(RecNo, "000000000") & Format(rstC.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(CLng(rstC.Tables(0).Rows(0)("amt").ToString() * 100), New String("0", 14)) &
      Format(IIf((rstC.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstC.Tables(0).Rows(0)("ChallanNo ").ToString() = vbNullString, Space(9), rstC.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(rstC.Tables(0).Rows(0)("DtOfVoucher").ToString(), "ddMMyyyy") &
      Format(IIf((rstC.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstC.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstC.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@"))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            'rstC.MoveNext
        Loop
        'DEDUCTEE DETAIL RECORD
        If rstD.Tables(0).Rows.Count > 0 Then
            'rstD.MoveFirst
        End If
        RecNo = 1      'Line no not reset as it will continue from challan detail...
        Do While Not rstD.Tables(0).Rows.Count
            TStrm.WriteLine(Format(LNo, "000000000") & "DD" & "000000001" & Format(RecNo, "000000000") & Format(rstD.Tables(0).Rows(0)("Sec").ToString(), "!@@@@@") &
      Format(IIf(rstD.Tables(0).Rows(0)("DType").ToString() = "O", 2, 1), "00") &
      Format(IIf((rstD.Tables(0).Rows(0)("DPan").ToString()), Space(10), IIf(rstD.Tables(0).Rows(0)("DPan").ToString() = vbNullString, Space(10), rstD.Tables(0).Rows(0)("DPan").ToString())), "!@@@@@@@@@@") &
      Format(rstD.Tables(0).Rows(0)("DName").ToString(), "!" & New String("@", 75)) & Format(rstD.Tables(0).Rows(0)("DAdd1").ToString(), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd2").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd2").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd2").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd3").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd3").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd3").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd4").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd4").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd4").ToString())), "!" & New String("@", 25)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DAdd5").ToString()), Space(25), IIf(rstD.Tables(0).Rows(0)("DAdd5").ToString() = vbNullString, Space(25), rstD.Tables(0).Rows(0)("DAdd5").ToString())), "!" & New String("@", 25)) &
      Format(rstD.Tables(0).Rows(0)("DState").ToString(), "00") & Format(rstD.Tables(0).Rows(0)("DPin").ToString(), "000000") & Format(rstD.Tables(0).Rows(0)("AmtOfPay").ToString() * 100, New String("0", 14)) &
      Format(rstD.Tables(0).Rows(0)("DtOfPay").ToString(), "ddMMyyyy") & Format(rstD.Tables(0).Rows(0)("RateOfTDS").ToString() * 100, "0000") & "N" &
      Format(rstD.Tables(0).Rows(0)("AmtOfTDS").ToString() * 100, New String("0", 14)) &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDS").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), Space(8), rstD.Tables(0).Rows(0)("DtOfTDSPay").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("BankBrCode").ToString()), Space(7), IIf(rstD.Tables(0).Rows(0)("BankBrCode").ToString() = vbNullString, Space(7), rstD.Tables(0).Rows(0)("BankBrCode").ToString())), "!@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("ChallanNo").ToString()), Space(9), IIf(rstD.Tables(0).Rows(0)("ChallanNo").ToString() = vbNullString, Space(9), rstD.Tables(0).Rows(0)("ChallanNo").ToString())), "!@@@@@@@@@") &
      Format(IIf((rstD.Tables(0).Rows(0)("CertificateDt").ToString()), Space(8), rstD.Tables(0).Rows(0)("CertificateDt").ToString()), "ddMMyyyy") &
      Format(IIf((rstD.Tables(0).Rows(0)("Reason").ToString()), Space(1), IIf(rstD.Tables(0).Rows(0)("Reason").ToString() = vbNullString, Space(1), rstD.Tables(0).Rows(0)("Reason").ToString())), "@") &
     New String("0", 14))
            LNo = LNo + 1
            RecNo = RecNo + 1
            If RecNo > 999999999 Then
                RecNo = 1
            End If
            ' rstD.MoveNext
        Loop
        TStrm.Close()
cleanup:
        'CLOSE THE FILE..

        'AND THE CONNECTIONS ALSO..
        rstC.Dispose()
        rstD.Dispose()
        rstRetn.Dispose()
        rstCSum.Dispose()
        rstDSum.Dispose()
        rstC = Nothing
        rstD = Nothing
        rstRetn = Nothing
        rstCSum = Nothing
        rstDSum = Nothing
        filename = Nothing
        oCoMst = Nothing

    End Sub


End Module
