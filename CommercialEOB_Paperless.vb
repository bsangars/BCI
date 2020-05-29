
' These Page functions are to suppress build errors when page number globals are referenced inside a Tablix.
' However they only return the correct values when used in the Page Header and Page Footer Sections.
Public Function PageNumber() As Integer
    Return Report.Globals!PageNumber
End Function

Public Function TotalPages() As Integer
    Return Report.Globals!TotalPages
End Function

Public Function PageNOfM() As String
    Return "Page " & Report.Globals!PageNumber & " of " & Report.Globals!TotalPages
End Function


' These String Compare functions ignore trailing space and are case insensitive
Public Function StrEquals(val1 As Object, val2 As Object) As Boolean
    If IsNothing(val1) Or IsNothing(val2) Then Return False
    If val1.ToString().TrimEnd().ToUpper() = val2.ToString().TrimEnd().ToUpper() Then Return True
    Return False
End Function

Public Function StrCompare(val1 As Object, val2 As Object) As Integer
    If IsNothing(val1) Or IsNothing(val2) Then Return -1
    Return String.Compare(val1.ToString().TrimEnd().ToUpper(), val2.ToString().TrimEnd().ToUpper())
End Function

Public Function InArray(needle As Object, haystack() As Object) As Boolean
    Dim i As Integer
    If TypeOf needle Is String Then
        For i = 0 To UBound(haystack)
            If StrEquals(needle, haystack(i)) Then Return True
        Next i
        Return False
    End If
    For i = 0 To UBound(haystack)
        If needle = haystack(i) Then Return True
    Next i
    Return False
End Function

Public Function GroupName(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Object, sGroupName As String) As String

    If sGroupName = "Group1" Then
        Return Param1
    End If

    If sGroupName = "Group2" Then
        Return Param1
    End If

    If sGroupName = "Group3" Then
        Return Param1
    End If

    If sGroupName = "Group4" Then
        Return Param1
    End If

    If sGroupName = "Group5" Then
        Return Param1
    End If

    If sGroupName = "Group6" Then
        Return Param1
    End If

    If sGroupName = "Group7" Then
        Return Param1
    End If

    If sGroupName = "Daily" Or sGroupName = "Group8" Then
        Return Format(Param1, "d")
    End If

    If sGroupName = "Group9" Then
        Return Param1
    End If

    If sGroupName = "Group10" Then
        Return Param1
    End If

    If sGroupName = "Group11" Then
        Return Param1
    End If

    If sGroupName = "Group12" Then
        Return Param1
    End If

    If sGroupName = "Weekly" Or sGroupName = "Group13" Then
        Return Format(DateAdd("d", 1 - DatePart("w", Param1), Param1), "d")
    End If
    Return Param1
End Function

'Note: Accessor functions implemented like Crystal One based Arrays
Private _Assigned_Notes(0) As String
Public Function SetAssigned_Notes(ByVal idx As Integer, ByVal value As String) As String
    If idx > 0 Then
        _Assigned_Notes(idx - 1) = value
    End If
    Return value
End Function
Public Function SetAssigned_Notes(ByVal value() As String) As String()
    _Assigned_Notes = value
    Return _Assigned_Notes
End Function
Public Function JoinAssigned_Notes() As String
    Return Join(_Assigned_Notes, "~")
End Function
Public Function SetAssigned_Notes(ByVal value As String) As String()
    _Assigned_Notes = Split(value, "~")
    Return _Assigned_Notes
End Function
Public Function GetAssigned_Notes(ByVal idx As Integer) As String
    If idx > 0 And idx <= _Assigned_Notes.Length Then Return _Assigned_Notes(idx - 1)
    Return ""
End Function
Public Function GetAssigned_Notes() As String()
    Return _Assigned_Notes
End Function
Public Function CntAssigned_Notes() As Integer
    Return _Assigned_Notes.Length
End Function

Private _clcl_id As String
Public Function Setclcl_id(ByVal value As String) As String
    _clcl_id = value
    Return _clcl_id
End Function
Public Function Getclcl_id() As String
    Return _clcl_id
End Function

Private _bln_Suppress_Page As Boolean
Public Function Setbln_Suppress_Page(ByVal value As Boolean) As Boolean
    _bln_Suppress_Page = value
    Return _bln_Suppress_Page
End Function
Public Function Getbln_Suppress_Page() As Boolean
    Return _bln_Suppress_Page
End Function

Private _swap As String
Public Function Setswap(ByVal value As String) As String
    _swap = value
    Return _swap
End Function
Public Function Getswap() As String
    Return _swap
End Function

Private _reset As Boolean
Public Function Setreset(ByVal value As Boolean) As Boolean
    _reset = value
    Return _reset
End Function
Public Function Getreset() As Boolean
    Return _reset
End Function

Private _cleb_cur_year As DateTime
Public Function Setcleb_cur_year(ByVal value As DateTime) As DateTime
    _cleb_cur_year = value
    Return _cleb_cur_year
End Function
Public Function Getcleb_cur_year() As DateTime
    Return _cleb_cur_year
End Function

Private _eob As String
Public Function Seteob(ByVal value As String) As String
    _eob = value
    Return _eob
End Function
Public Function Geteob() As String
    Return _eob
End Function

Private _section2 As Boolean
Public Function Setsection2(ByVal value As Boolean) As Boolean
    _section2 = value
    Return _section2
End Function
Public Function Getsection2() As Boolean
    Return _section2
End Function

Private _bln_show_page_header As Boolean
Public Function Setbln_show_page_header(ByVal value As Boolean) As Boolean
    _bln_show_page_header = value
    Return _bln_show_page_header
End Function
Public Function Getbln_show_page_header() As Boolean
    Return _bln_show_page_header
End Function

Private _bln_force_new_page As Boolean
Public Function Setbln_force_new_page(ByVal value As Boolean) As Boolean
    _bln_force_new_page = value
    Return _bln_force_new_page
End Function
Public Function Getbln_force_new_page() As Boolean
    Return _bln_force_new_page
End Function

'Note: Accessor functions implemented like Crystal One based Arrays
Private _Segment_Array(0) As String
Public Function SetSegment_Array(ByVal idx As Integer, ByVal value As String) As String
    If idx > 0 Then
        _Segment_Array(idx - 1) = value
    End If
    Return value
End Function
Public Function SetSegment_Array(ByVal value() As String) As String()
    _Segment_Array = value
    Return _Segment_Array
End Function
Public Function JoinSegment_Array() As String
    Return Join(_Segment_Array, "~")
End Function
Public Function SetSegment_Array(ByVal value As String) As String()
    _Segment_Array = Split(value, "~")
    Return _Segment_Array
End Function
Public Function GetSegment_Array(ByVal idx As Integer) As String
    If idx > 0 And idx <= _Segment_Array.Length Then Return _Segment_Array(idx - 1)
    Return ""
End Function
Public Function GetSegment_Array() As String()
    Return _Segment_Array
End Function
Public Function CntSegment_Array() As Integer
    Return _Segment_Array.Length
End Function

Private _Counter As Double
Public Function SetCounter(ByVal value As Double) As Double
    _Counter = value
    Return _Counter
End Function
Public Function GetCounter() As Double
    Return _Counter
End Function
Public Function AddCounter(ByVal value As Double) As Double
    _Counter = _Counter + value
    Return _Counter
End Function

Public Function CRFadjusted_claims_notes_heading(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    If Fields!REPORT_SECTION.Value = 3 And Len(Fields!CLCL_ID_ADJ_FROM.Value) > 0 Then
        CRFadjusted_claims_notes_heading = "This claim was previously processed. This is the reprocessed claim. Please see the ADJUSTED CLAIMS section for the original information."
    End If
    If Fields!REPORT_SECTION.Value = 4 Then
        CRFadjusted_claims_notes_heading = "This claim was processed on " & Format(Fields!CLCL_PAID_DT.Value, "MM/dd/yy") & ".  " & "It is shown here for your reference.  Amounts shown are not included in the SUMMARY section.  " & "Please see Insurance Claim " & Fields!CLCL_ID_ADJ_TO.Value & " in the DETAIL section for " & "the reprocessed claim."
    End If
End Function

Public Function CRFamount_we_paid_notes(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As String
    If Fields!CLCL_PAY_PR_IND.Value = "P" Then
        CRFamount_we_paid_notes = "Payment made to Provider"
    End If
    If Fields!CLCL_PAY_PR_IND.Value = "S" And Param1 > 0 Then
		CRFamount_we_paid_notes = "Check Mailed Separately"
		'Checks are mailed seperately irrespective of Go Green Flag.
        'If UCase(RTrim(Fields!GOGREEN.Value)) = "Y" Then
           ' CRFamount_we_paid_notes = "Check Mailed Separately"
        'Else
            'CRFamount_we_paid_notes = "Check enclosed"
        'End If
    End If
End Function

Public Function CRFassign_notes_descriptions(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As Integer
    Dim i As Double
    Dim Assigned_Notes As Object
    If Getclcl_id() <> Fields!BCI_CL56_EOB_RPT_CLCL_ID.Value Then
        Setclcl_id(Fields!BCI_CL56_EOB_RPT_CLCL_ID.Value)
        For i = 1 To 15
            Assigned_Notes(i - 1) = ""
        Next i
    End If
    If Fields!NOTES.Value > 0 Then
        Assigned_Notes(Fields!NOTES.Value - 1) = Trim(Fields!NOTES_TEXT.Value)
    End If
    CRFassign_notes_descriptions = Fields!NOTES.Value
End Function

Public Function CRFcharges_accounting_for_duplicates(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As double
    Dim charges As double
    charges = Fields!CHARGES.Value
    If IsNothing(Fields!CL55_EOB_DUPLICATE_CLAIM.Value) Then
        CRFcharges_accounting_for_duplicates = charges
    ElseIf UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) <> "Y" Then
        CRFcharges_accounting_for_duplicates = charges
    ElseIf UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" Then
        CRFcharges_accounting_for_duplicates = 0
    End If
End Function

Public Function CRFcover_page(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim output As String
    Dim centerVert As String
    Dim i As Double
    Select Case Fields!CRYSTAL_SEGMENT.Value
        Case "SCME", "MCME", "MCSE", "SCSE"
            output = "EOBs with Checks" & Chr(13) & "(SCME, MCME, MCSE, SCSE)"
        Case "HMP"
            output = "Householding Multipage" & Chr(13) & "(HMP)"
        Case "SMP"
            output = "Single Multipage" & Chr(13) & "(SMP)"
        Case "SP"
            output = "Single Page" & Chr(13) & "(SP)"
        Case "OOB"
            output = "Out-of-Balance" & Chr(13) & "(OOB)"
        Case "COOB"
            output = "Out-of-Balance With Check" & Chr(13) & "(COOB)"
    End Select
    For i = 1 To 3
        centerVert = centerVert & Chr(13)
    Next i
    CRFcover_page = centerVert & "CL - EOB - BCI"
End Function

Public Function CRFdeductible_flag_assign(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As Boolean
    If Fields!REPORT_SECTION.Value = 2 Then
        Setsection2(True)
    End If
    CRFdeductible_flag_assign = Getsection2()
End Function

Public Function CRFdeductible_flag_initalize() As String
    Setsection2(False)
    CRFdeductible_flag_initalize = ""
End Function

Public Function CRFfam_deduction_status(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim start_dt As String
    Dim end_dt As String
    Dim used As String
    Dim deductible As String
    Dim deduct_desc As String
    start_dt = Format(Fields!CLEB_CUR_YEAR.Value, "MM/dd/yy")
    end_dt = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Fields!CLEB_CUR_YEAR.Value)), "MM/dd/yy")
    used = Format(Fields!FAM_DED_AMT.Value, "0.00")
    deductible = Format(Fields!FAM_DED_MAX.Value, "0.00")
    deduct_desc = Chr(149) & Space(1) & used & " of the " & deductible & Space(1) & " Family " & Fields!XREF_DESC.Value
    If Geteob() <> Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value Then
        Seteob(Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value)
        Setcleb_cur_year(CDate("01/01/1753"))
    End If
    If Getcleb_cur_year() <= Fields!CLEB_CUR_YEAR.Value Then
        Setcleb_cur_year(Fields!CLEB_CUR_YEAR.Value)
        If Fields!FAM_DED_MAX.Value >= 999999 And Fields!FAM_AGG_MAX.Value > 0 Then
            CRFfam_deduction_status = Chr(149) & Space(1) & Format(Fields!FAM_AGG_CNT.Value) & " of " & Format(Fields!FAM_AGG_MAX.Value) & " family members have met their Deductible."
        Else
            CRFfam_deduction_status = deduct_desc
        End If
    Else
        CRFfam_deduction_status = deduct_desc
    End If
End Function

private _paidflag as string 
Public Function SetPaidflag (ByVal value as String) as String 
_paidflag = value
return ""
End Function 

Public Function GetPaidflag() 
return _paidflag
end function


Public Function CRFfamlabel(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim start_dt As String
    Dim end_dt As String
    Dim used As String
    Dim deductible As String
    Dim deduct_desc As String
    start_dt = Format(Fields!CLEB_CUR_YEAR.Value, "MM/dd/yy")
    end_dt = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Fields!CLEB_CUR_YEAR.Value)), "MM/dd/yy")
    If Geteob() <> Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value Then
        Seteob(Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value)
        Setcleb_cur_year(CDate("01/01/1753"))
    End If
    If Getcleb_cur_year() <= Fields!CLEB_CUR_YEAR.Value Then
        Setcleb_cur_year(Fields!CLEB_CUR_YEAR.Value)
        If Fields!FAM_DED_MAX.Value >= 999999 Then
            CRFfamlabel = "For benefit period " & start_dt & "-" & end_dt & ", the following has been satisfied:"
        Else
            CRFfamlabel = "For benefit period " & start_dt & "-" & end_dt & ", the following has been satisfied:"
        End If
    End If
End Function

Public Function CRFflag_display_pagenumber_in_footer() As String
    If PageNumber() = 1 Then
        Setbln_Suppress_Page(False)
    Else
        Setbln_Suppress_Page(True)
    End If
    CRFflag_display_pagenumber_in_footer = ""
End Function

Public Function CRFflag_suppress_pagenumber_in_footer() As String
    Setbln_Suppress_Page(True)
    CRFflag_suppress_pagenumber_in_footer = ""
End Function

Public Shared Dim _strAddress(6) AS String

Public Function SetStrAddress(grgrid as String, addrname as String, addr1 as string, addr2 as string, addr3 as string, csz as string, mccyname as string)
	_strAddress(0)= grgrid
	_strAddress(1)= addrname
	_strAddress(2)= addr1
	_strAddress(3)= addr2
	_strAddress(4)= addr3
	_strAddress(5)= csz
	_strAddress(6)= mccyname
End Function


Public Function CRFgf_1_address_member_info(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim i As Integer
    Dim strOutput As String = ""
	
	SetStrAddress(Fields!GRGR_ID.Value, Fields!ADDR_NAME.Value, Fields!ADDR1.Value, Fields!ADDR2.Value, Fields!ADDR3.Value, Fields!CSZ.Value, Fields!MCCY_NAME.Value)
	
    For i = 1 To 7
        If Len(TRIM(_strAddress(i - 1))) > 0 Then
            strOutput = strOutput & _strAddress(i - 1) & vbcrlf
        End If
    Next i
    CRFgf_1_address_member_info = UCase(strOutput)
    strOutput = ""
	SetStrAddress("", "", "", "", "", "", "")
End Function

Public Function CRFgh_3_report_section_description(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer) As String
    Select Case Fields!REPORT_SECTION.Value
        Case 1
            CRFgh_3_report_section_description = "SUMMARY"
        Case 3
            If Param1 < 3 Then
                CRFgh_3_report_section_description = "DETAIL"
            End If
        Case 4
            If Param1 < 4 Then
                CRFgh_3_report_section_description = "ADJUSTED CLAIMS"
            End If
        Case Else
            CRFgh_3_report_section_description = Format(Fields!REPORT_SECTION.Value, "0")
    End Select
End Function

Public Function CRFgroup_id_name(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim GrpName As String
    If Len(Fields!GRGR_NAME.Value) > 0 Then
        GrpName = Trim(Fields!GRGR_ID.Value) & " - " & Trim(Fields!GRGR_NAME.Value)
    Else
        GrpName = Trim(Fields!GRGR_ID.Value)
    End If
    CRFgroup_id_name = GrpName
End Function

Public Function CRFind_deductible_status(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim start_dt As String
    Dim end_dt As String
    Dim used As String
    Dim deductible As String
    Dim deduct_desc As String
    start_dt = Format(Fields!CLEB_CUR_YEAR.Value, "MM/dd/yy")
    end_dt = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Fields!CLEB_CUR_YEAR.Value)), "MM/dd/yy")
    used = Format(Fields!INDIV_DED_AMT.Value, "0.00")
    deductible = Format(Fields!INDIV_DED_MAX.Value, "0.00")
    deduct_desc = Chr(149) & Space(1) & used & " of the " & deductible & Space(1) & " Individual " & Fields!XREF_DESC.Value
    If Geteob() <> Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value Then
        Seteob(Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value)
        Setcleb_cur_year(CDate("01/01/1753"))
    End If
    If Getcleb_cur_year() <= Fields!CLEB_CUR_YEAR.Value Then
        Setcleb_cur_year(Fields!CLEB_CUR_YEAR.Value)
        CRFind_deductible_status = deduct_desc
    Else
        CRFind_deductible_status = "       " & deduct_desc
    End If
End Function

Public Function CRFindlabel(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim start_dt As String
    Dim end_dt As String
    Dim used As String
    Dim deductible As String
    Dim deduct_desc As String
    start_dt = Format(Fields!CLEB_CUR_YEAR.Value, "MM/dd/yy")
    end_dt = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Fields!CLEB_CUR_YEAR.Value)), "MM/dd/yy")
    If Geteob() <> Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value Then
        Seteob(Fields!CRYSTAL_SEGMENT.Value & CStr(Fields!MEME_CK.Value) & Fields!PDPD_CARD_STOCK.Value)
        Setcleb_cur_year(CDate("01/01/1753"))
    End If
    If Getcleb_cur_year() <= Fields!CLEB_CUR_YEAR.Value Then
        Setcleb_cur_year(Fields!CLEB_CUR_YEAR.Value)
        CRFindlabel = "For benefit period " & start_dt & "-" & end_dt & ", the following has been satisfied:"
    End If
End Function

Public Function CRFmember_name(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As String, Param2 As String) As String
    Dim MemberName As String
    If Len(TRIM(Fields!MEME_MID_INIT.Value)) > 0 Then
        'Return Fields!MEME_FIRST_NAME.Value & " " & Trim(Param1) & ". " & Trim(Param2)
        MemberName = Fields!MEME_FIRST_NAME.Value & " " & Trim(Param1) & ". " & Trim(Param2)
    Else
        'Return Fields!MEME_FIRST_NAME.Value & " " & Trim(Param2)
        MemberName = Trim(Fields!MEME_FIRST_NAME.Value) & " " & Trim(Param2)
    End If
    CRFmember_name = MemberName
End Function

Public Function CRFmeme_sfx(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Return Format(Fields!MEME_SFX.Value, "00")
End Function

Public Function CRFnotes_1_display() As String
    Dim Assigned_Notes As String
    CRFnotes_1_display = Assigned_Notes(1 - 1)
End Function

Public Function CRFnotes_10_display() As String
    Dim Assigned_Notes As String
    CRFnotes_10_display = Assigned_Notes(10 - 1)
End Function

Public Function CRFnotes_11_display() As String
    Dim Assigned_Notes As String
    CRFnotes_11_display = Assigned_Notes(11 - 1)
End Function

Public Function CRFnotes_12_display() As String
    Dim Assigned_Notes As String
    CRFnotes_12_display = Assigned_Notes(12 - 1)
End Function

Public Function CRFnotes_13_display() As String
    Dim Assigned_Notes As String
    CRFnotes_13_display = Assigned_Notes(13 - 1)
End Function

Public Function CRFnotes_14_display() As String
    Dim Assigned_Notes As String
    CRFnotes_14_display = Assigned_Notes(14 - 1)
End Function

Public Function CRFnotes_15_display() As String
    Dim Assigned_Notes As String
    CRFnotes_15_display = Assigned_Notes(15 - 1)
End Function

Public Function CRFnotes_2_display() As String
    Dim Assigned_Notes As String
    CRFnotes_2_display = Assigned_Notes(2 - 1)
End Function

Public Function CRFnotes_3_display() As String
    Dim Assigned_Notes As String
    CRFnotes_3_display = Assigned_Notes(3 - 1)
End Function

Public Function CRFnotes_4_display() As String
    Dim Assigned_Notes As String
    CRFnotes_4_display = Assigned_Notes(4 - 1)
End Function

Public Function CRFnotes_5_display() As String
    Dim Assigned_Notes As String
    CRFnotes_5_display = Assigned_Notes(5 - 1)
End Function

Public Function CRFnotes_6_display() As String
    Dim Assigned_Notes As String
    CRFnotes_6_display = Assigned_Notes(6 - 1)
End Function

Public Function CRFnotes_7_display() As String
    Dim Assigned_Notes As String
    CRFnotes_7_display = Assigned_Notes(7 - 1)
End Function

Public Function CRFnotes_8_display() As String
    Dim Assigned_Notes As String
    CRFnotes_8_display = Assigned_Notes(8 - 1)
End Function

Public Function CRFnotes_9_display() As String
    Dim Assigned_Notes As String
    CRFnotes_9_display = Assigned_Notes(9 - 1)
End Function

Public Function CRFnotes_label(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim Assigned_Notes As String
    If Fields!REPORT_SECTION.Value = 4 Or Len(Assigned_Notes(1 - 1)) > 0 Or Len(Fields!CLCL_ID_ADJ_FROM.Value) > 0 Then
        CRFnotes_label = "Notes"
	Else 
		CRFnotes_label = "None"
    End If
End Function

Public Function CRFover_continued() As String
    If Not Getbln_Suppress_Page() Then
        If Report.Globals!PageNumber = 1 And Getreset() = False Then
            Setswap("(OVER)")
            Setreset(True)
        Else
            If Getswap() = "(OVER)" Then
                Setswap("(CONTINUED)")
            Else
                Setswap("(OVER)")
            End If
        End If
        CRFover_continued = Getswap()
    Else
        Setreset(False)
    End If
End Function

Public Function CRFpaid_per_contract_detl(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As String
    If UCase(RTrim(Fields!PAID_PER_CONTRACT_DETL.Value)) = "Y" Then
        Return "Paid Per Contract"
    End If
    Return Mid(Format(Param1, "###,###,###.##"), 1, Len(Format(Param1, "###,###,###.##")))
End Function

Public Function CRFpaid_per_contract_summ(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As String
    If UCase(RTrim(Fields!PAID_PER_CONTRACT_SUMM.Value)) = "Y" Then
        Return "Paid Per Contract"
    End If
    Return Param1
End Function

Public Function CRFph_pagenumber() As String
    Dim PgeNo As Double
    Dim PgeOf As Double
    Dim i As Double
    PgeNo = Round(Report.Globals!PageNumber / 2)
    PgeOf = Round(Report.Globals!TotalPages / 2)
    CRFph_pagenumber = "PAGE " & Format(PgeNo, "0") & " of " & Format(PgeOf, "0")
End Function

Public Function CRF_ph_pagenumber() As String
    Dim PgeNo As Double
    Dim PgeOf As Double
    Dim i As Double
    PgeNo = Round(PageNumber() / 2)
    PgeOf = Round(TotalPages() / 2)
    Return "PAGE " & Format(PgeNo, "0") & " of " & Format(PgeOf, "0")
End Function

Public Function CRFph_customer_service_phone_and_language(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    CRFph_customer_service_phone_and_language = "If you have a question about your" & Chr(13) & "claim, please call Customer Service at" & Chr(13) & Fields!CONTACT_INFO.Value
End Function

Public Function CRFpitbow_paid_per_contract_summ(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As String
    If UCase(RTrim(Fields!CRYSTAL_SEGMENT.Value)) = "SCME" Then
        Return Fields!CKPY_NET_AMTS_CONCAT.Value
    ElseIf UCase(RTrim(Fields!PAID_PER_CONTRACT_SUMM.Value)) = "Y" Then
        Return "Paid Per Contract"
    ElseIf Param1 < 1 Then
        Return Mid(CStr(Param1), 3, Len(CStr(Param1)))
    End If
    Return Mid(CStr(Param1), 2, Len(CStr(Param1)))
End Function

Public Function CRFsegment_array_return() As String
    Dim i As Double
    Dim output As String
    Dim Segment_Array As Object
    For i = 1 To 9
        If Len(Segment_Array(i - 1)) > 0 Then
            output = Segment_Array(i - 1) & "," & output
        End If
    Next i
    If Len(output) > 1 Then
        CRFsegment_array_return = Mid(output, 1, Len(output) - 1)
    End If
End Function

Public Function CRFsegment_array(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As String
    Dim i As Double
    Dim bln_Test As Boolean
    Dim Segment_Array As Object
    For i = 1 To 9
        If StrEquals(Segment_Array(i - 1), Trim(Fields!CRYSTAL_SEGMENT.Value)) Then
            bln_Test = True
            Exit For
        End If
    Next i
    If Not bln_Test Then
        For i = 1 To 9
            If Len(Segment_Array(i - 1)) = 0 Then
                Segment_Array(i - 1) = Trim(Fields!CRYSTAL_SEGMENT.Value)
                Exit For
            End If
        Next i
    End If
    CRFsegment_array = Segment_Array(1 - 1)
End Function

Public Function copaycoins1_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function copaycoins2_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function deductible1_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function deductible2_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function DetailSection2_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer, Param2 As Double) As Boolean
     
	IF ((Fields!REPORT_SECTION.Value <>3 AND Fields!REPORT_SECTION.Value <>4) Or Param1 <> Param2) then 
	Return True
	Else Return False
	END IF
End Function

Public Function DetailSection20_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer, Param2 As Double) As Boolean
    Dim Assigned_Notes As String
    If Param1 = Param2 Then
        If Fields!REPORT_SECTION.Value = 4 Or Len(Assigned_Notes(1 - 1)) > 0 Or Len(Fields!CLCL_ID_ADJ_FROM.Value) > 0 Then
            DetailSection20_Hidden = False
        Else
            DetailSection20_Hidden = True
        End If
    Else
        DetailSection20_Hidden = True
    End If
End Function

Public Function DetailSection21_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As Boolean
    If Fields!REPORT_SECTION.Value <> 2 Or Fields!DEDUCTIBLE_SECTION.Value <> 2 Then
        Return True
    ElseIf IsNothing(CRFfam_deduction_status(Fields)) Then
        Return True
    ElseIf UCase(RTrim(Fields!FAM_DED_MET.Value)) = "NA" Then
        Return True
    ElseIf Fields!FAM_DED_MAX.Value >= 999999 And Fields!FAM_AGG_MAX.Value = 0 And Fields!FAM_AGG_CNT.Value = 0 Then
        Return True
    End If
End Function

Public Function DetailSection22_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer) As Boolean
    If Fields!REPORT_SECTION.Value = 2 Then
        Return Not Param1 = 3
    End If
    Return True
End Function

Public Function DetailSection26_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As Boolean
    If IsNothing(CRFind_deductible_status(Fields)) Then
        Return True
    ElseIf Fields!INDIV_DED_MAX.Value > 999999 Then
        Return True
    ElseIf Fields!REPORT_SECTION.Value <> 2 Or Fields!DEDUCTIBLE_SECTION.Value <> 1 Then
        Return True
    End If
End Function

Public Function DetailSection3_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer, Param2 As Double) As Boolean
    Return Param1 <> Param2
End Function

Public Function FamLabel1_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) As Boolean
    Return AddCounter(1) > 1 And Not IsNothing(CRFind_deductible_status(Fields)) And Not (Fields!INDIV_DED_MAX.Value > 999999)
End Function

Public Function GroupHeaderSection20_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer) As Boolean
    If InArray(Fields!REPORT_SECTION.Value, New Object() {3, 4}) Then
        Return Not Param1 = 1
    End If
    Return True
End Function

Public Function GroupHeaderSection21_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer) As Boolean
    If InArray(Fields!REPORT_SECTION.Value, New Object() {3, 4}) Then
        Return Param1 = 1
    End If
    Return True
End Function

Public Function ChargesSum(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) as double 
	If Fields!REPORT_SECTION.Value =1 then 
	Return Fields!CHARGES.Value
	Else
	Return 0
	End If
End Function

Public Function SavingsSum(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) as double 
	If Fields!REPORT_SECTION.Value =1 then 
	Return Fields!SAVINGS.Value
	Else
	Return 0
	End If
End Function

Public Function PaidSum(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) as double 
	If Fields!REPORT_SECTION.Value =1 then 
	Return Fields!PAID.Value
	Else
	Return 0
	End If
End Function


Public Function GroupHeaderSection7_Hidden() As Boolean
    If PageNumber() = 1 Then
        GroupHeaderSection7_Hidden = False
    Else
        GroupHeaderSection7_Hidden = True
    End If
End Function

Public Function GroupHeaderSection8_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Fields!REPORT_SECTION.Value > 1 Or Param1 = 0
End Function

Public Function noncovered1_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function noncovered2_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function otherinsurance1_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function otherinsurance2_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Double) As Boolean
    Return Param1 = 0 Or (UCase(RTrim(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)) = "Y" And Fields!PAID.Value = 0)
End Function

Public Function PageHeaderSection1_Hidden() As Boolean
    If Report.Globals!PageNumber Mod 2 = 0 Or Report.Globals!PageNumber = 1 Then
        PageHeaderSection1_Hidden = True
    Else
        PageHeaderSection1_Hidden = False
    End If
End Function

Public Function PageHeaderSection2_Hidden() As Boolean
    If Report.Globals!PageNumber Mod 2 = 0 Or Report.Globals!PageNumber = 1 Then
        PageHeaderSection2_Hidden = True
    Else
        PageHeaderSection2_Hidden = False
    End If
End Function

Public Function PageHeaderSection3_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer, Param2 As String, Param3 As Integer) As Boolean
    If Report.Globals!PageNumber Mod 2 = 0 Or Report.Globals!PageNumber = 1 Then
        PageHeaderSection3_Hidden = True
    Else
        If Fields!REPORT_SECTION.Value = 3 Or Fields!REPORT_SECTION.Value = 4 Then
            If Param1 = 1 Then
                PageHeaderSection3_Hidden = True
            Else
                PageHeaderSection3_Hidden = False
            End If
        Else
            PageHeaderSection3_Hidden = True
        End If
        If Not StrEquals(Param2, Fields!PDPD_CARD_STOCK.Value) Or Param3 <> Fields!MEME_CK.Value Then
            PageHeaderSection3_Hidden = True
        End If
    End If
End Function

Public Function RTsavings1_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer, Param2 As Integer) As Boolean
    Return Param1 = Param2
End Function

Public Function Text67_Hidden(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, Param1 As Integer) As Boolean
    If Fields!REPORT_SECTION.Value = 2 Then
        Return Not Param1 = 1
    End If
    Return True
End Function

Public Function Text71_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text71_Hidden = False
    Else
        Text71_Hidden = True
    End If
End Function

Public Function Text75_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text75_Hidden = False
    Else
        Text75_Hidden = True
    End If
End Function

Public Function Text76_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text76_Hidden = False
    Else
        Text76_Hidden = True
    End If
End Function

Public Function Text77_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text77_Hidden = False
    Else
        Text77_Hidden = True
    End If
End Function

Public Function Text78_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text78_Hidden = True
    Else
        Text78_Hidden = False
    End If
End Function

Public Function Text79_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text79_Hidden = True
    Else
        Text79_Hidden = False
    End If
End Function

Public Function Text80_Hidden() As Boolean
    Dim section2 As Boolean
    If section2 Then
        Text80_Hidden = True
    Else
        Text80_Hidden = False
    End If
End Function

Public Function Text99_Hidden() As Boolean
    If Report.Globals!PageNumber = 1 Then
        If CRFover_continued() > "" Then
            Text99_Hidden = False
        Else
            Text99_Hidden = True
        End If
    Else
        Text99_Hidden = True
    End If
End Function

Public Function HideDuplicateDetail(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, deductible as double, copaycoins as double, concat as String ) As Boolean							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat)) Then 
	HideDuplicateDetail=True 
Else 
	HideDuplicateDetail=False 
End IF
return HideDuplicateDetail
End Function

Public Shared Dim _ChargeSum as Double

Public Function SetChargeSum(ByVal value as Double) as Double
	_ChargeSum = value
	return _ChargeSum
End Function 

Public Function HideDuplicateCharges(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, 
servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, 
deductible as double, copaycoins as double, concat as String ) As Double							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat)) Then 
	HideDuplicateCharges= 0.00 
Else 
	_ChargeSum = _ChargeSum+ Cdbl(Fields!CHARGES.Value)
	HideDuplicateCharges= Cdbl(Fields!CHARGES.Value)
End IF
return HideDuplicateCharges
End Function

Public Function GetChargeSum() as Double
	return _ChargeSum
End Function 

Public Shared Dim _SavingSum as Double

Public Function SetSavingSum(ByVal value as Double) as Double
	_SavingSum = value
	return _SavingSum
End Function 

Public Function HideDuplicateSavings(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, 
servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, 
deductible as double, copaycoins as double, concat as String ) As Double							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat)) Then 
	HideDuplicateSavings= 0.00 
Else 
	_SavingSum = _SavingSum+ Cdbl(Fields!SAVINGS.Value)
	HideDuplicateSavings= Cdbl(Fields!SAVINGS.Value)
End IF
return HideDuplicateSavings
End Function

Public Function GetSavingsSum() as Double
	return _SavingSum
End Function 

Public Shared Dim _PaidSum as Double

Public Function SetPaidSum(ByVal value as Double) as Double
	_PaidSum = value
	return _PaidSum
End Function 

Public Function HideDuplicatePaid(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, 
servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, 
deductible as double, copaycoins as double, concat as String, paid as double ) As Double							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat AND
	Cdbl(Fields!PAID.Value)= paid)) Then 
	HideDuplicatePaid= 0.00 
Else 
	_PaidSum = _PaidSum+ Cdbl(Fields!PAID.Value)
	HideDuplicatePaid= Cdbl(Fields!PAID.Value)
End IF
return HideDuplicatePaid
End Function

Public Function GetPaidSum() as Double
	return _PaidSum
End Function 

Public Shared Dim _PaidOtherIns as Double

Public Function SetPaidOtherIns(ByVal value as Double) as Double
	_PaidOtherIns = value
	return _PaidOtherIns
End Function 

Public Function HideDuplicateOtherIns(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, 
servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, 
deductible as double, copaycoins as double, concat as String, paid as double ) As Double							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat AND
	Cdbl(Fields!PAID.Value)= paid)) Then 
	HideDuplicateOtherIns= 0.00 
Else 
	_PaidOtherIns = _PaidOtherIns+ Cdbl(Fields!OTHER_INSURANCE.Value)
	HideDuplicateOtherIns= Cdbl(Fields!OTHER_INSURANCE.Value)
	
return HideDuplicateOtherIns
END IF
End Function

Public Function GetPaidOtherIns() as Double
	return _PaidOtherIns
End Function 


Public Shared Dim _PaidNonCov as Double

Public Function SetPaidNonCov(ByVal value as Double) as Double
	_PaidNonCov = value
	return _PaidNonCov
End Function 

Public Function HideDuplicateNonCov(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, 
servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, 
deductible as double, copaycoins as double, concat as String, paid as double ) As Double							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat AND
	Cdbl(Fields!PAID.Value)= paid)) Then 
	HideDuplicateNonCov= 0.00 
Else 
	_PaidNonCov = _PaidNonCov+ Cdbl(Fields!NONCOVERED.Value)
	HideDuplicateNonCov= Cdbl(Fields!NONCOVERED.Value)
	
return HideDuplicateNonCov
END IF
End Function

Public Function GetPaidNonCov() as Double
	return _PaidNonCov
End Function 

Public Shared Dim _PaidCopayins as Double

Public Function SetPaidCopayins(ByVal value as Double) as Double
	_PaidCopayins = value
	return _PaidCopayins
End Function 

Public Function HideDuplicateCopayins(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, 
servicedate as String, servicedesc as String, charges as double, savings as double, otherinsurance as double, noncovered as double, 
deductible as double, copaycoins as double, concat as String, paid as double ) As Double							
IF (Fields!REPORT_SECTION.Value <>3  OR ( 
	Cstr(Fields!BCI_CL56_EOB_RPT_SERVICE_DT.Value)= servicedate AND 
	Fields!SERVICE_DESC.Value = servicedesc AND
	Cdbl(Fields!CHARGES.Value) = charges AND
	Cdbl(Fields!SAVINGS.Value) = savings AND
	Cdbl(Fields!OTHER_INSURANCE.Value) = otherinsurance AND
	Cdbl(Fields!NONCOVERED.Value) = noncovered AND
	Cdbl(Fields!DEDUCTIBLE.Value) = deductible AND
	Cdbl(Fields!COPAY_COINS.Value) = copaycoins AND
	Fields!CONCAT.Value = concat AND
	Cdbl(Fields!PAID.Value)= paid)) Then 
	HideDuplicateCopayins= 0.00 
Else 
	_PaidCopayins = _PaidCopayins+ Cdbl(Fields!COPAY_COINS.Value)
	HideDuplicateCopayins= Cdbl(Fields!COPAY_COINS.Value)
	
return HideDuplicateCopayins
END IF
End Function

Public Function GetPaidCopayins() as Double
	return _PaidCopayins
End Function 

Public Shared dim _NotesArray(15) as String
Public Shared Dim _arrayindx as Integer
Public Shared Dim _CLCLid as String 
Public Function SetCLCLId(ByVal value as String) as Boolean
	_CLCLid = value
return _CLCLid = value
End Function

Public Shared Dim _note as Integer
Public Shared Dim _NoteString as String

Public Function SetNoteString() as Boolean 
	_NoteString =""
	return _NoteString = ""
End Function 

Public Function GetNoteString() as String 
	return _NoteString
End Function

Public Function Notesassign(Fields As Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, prevnotes as Integer) as Boolean
	Dim i as Double 
	If Fields!NOTES.Value > 0 AND CInt(Fields!NOTES.Value) <> prevnotes  Then
		_NoteString = _NoteString & CStr(Fields!NOTES.Value) & ". " & Fields!NOTES_TEXT.Value & "~"
	End IF 
	Notesassign = True
End Function	
 
Public Function Notesdesc(i as integer) as String 
	_NotesArray = Split(_NoteString,"~")
	
	IF  _NotesArray(i-1) <> "" AND Len(_NotesArray(i-1)) > 0 Then
		Notesdesc = Split(_NoteString,"~")(i-1) 	
	Else Notesdesc =""
	END IF 
	return Notesdesc
End Function


Public Function DupSeqvalues(Fields as Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, previousSeq as Integer, id as Integer ) as Double
	Dim fieldamt as Double
	IF  TRIM(Fields!CL55_EOB_DUPLICATE_CLAIM.Value)<>"Y" Then 
	Select Case id 
	Case 1 
	fieldamt = Fields!CHARGES.Value
	_ChargeSum = _ChargeSum+CDbl(Fields!CHARGES.Value)
	Case 2 
	fieldamt = Fields!SAVINGS.Value
	_SavingSum = _SavingSum + Cdbl(Fields!SAVINGS.Value)
	Case 3
	fieldamt = Fields!OTHER_INSURANCE.Value
	_PaidOtherIns = _PaidOtherIns + CDbl(Fields!OTHER_INSURANCE.Value)
	Case 4 
	fieldamt = Fields!NONCOVERED.Value
	_PaidNonCov = _PaidNonCov + Cdbl(Fields!NONCOVERED.Value)
	Case 5
	fieldamt = Fields!DEDUCTIBLE.Value
	Case 6 
	fieldamt = Fields!COPAY_COINS.Value
	_PaidCopayins = _PaidCopayins + Cdbl(Fields!COPAY_COINS.Value)
	Case 7
	fieldamt = Fields!PAID.Value
	_PaidSum = _PaidSum + CDbl(Fields!PAID.Value)
	End Select
	DupSeqvalues = fieldamt
	End IF 
End Function 
	
Public function Hidedups(Fields as Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields, seqno as Integer ,ClaimId as String) as boolean 
	Dim flag as Boolean
	IF Fields!BCI_CL56_EOB_RPT_CDML_SEQ_NO.Value = seqno AND TRIM(CStr(Fields!BCI_CL56_EOB_RPT_CLCL_ID.Value)) = TRIM(ClaimId) then 
	flag = True
	else 
	flag = False 
	END IF 
	Hidedups = flag
End Function

Public Shared Dim _contractflag as String
Public Function resetContratFlag() as Boolean 
	_contractflag = ""
	return True
End Function

Public Function SetPaidContractFlag(Fields as Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) as String 
	IF Ucase(TRIM(Fields!PAID_PER_CONTRACT.Value)) = "Y" then 
	_contractflag = "Y"
	End IF 
return ""
End Function

Public Function GetpaidContractFlag() as String 
return _contractflag
End Function 

Public Shared Dim _Count as Integer

Public Function SetCount()as Boolean
_Count = 0 
return _Count = 0 
End Function

Public Function UpdCount() as String 
_Count = _Count +1
return ""
End Function 

Public Function HideTotal(param1 as Integer)as Boolean
return not param1 = _Count
End Function 


Public Function FamilyDeductible(Fields as Microsoft.ReportingServices.ReportProcessing.ReportObjectModel.Fields) as String 
Dim Deductible as String
IF CInt(Fields!FAM_DED_MAX.Value)> 99999 then 
Deductible = CStr(Fields!FAM_AGG_CNT.Value) & " of the " & CStr(Fields!FAM_AGG_MAX.Value) & " family members have met their Deductible"
ELSE 
Deductible = CStr(Fields!FAM_DED_AMT.Value) & " of the " & CStr(Fields!FAM_DED_MAX.Value) & " Family Deductible"
End IF 
return Deductible
End Function


Public Shared dim _AscFlag as String 

Public Function SetAscFlag(ByVal value as string) as String 
_AscFlag = value
return ""
End Function 

Public Function GetAscFlag() as String 
return _AscFlag
End Function 