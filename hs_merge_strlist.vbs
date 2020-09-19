Function hs_merge_strlist(ByVal strList1, ByVal strList2, ByVal chr_join)
	'*** History ***********************************************************************************
	' 2020/09/19, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' 	Return merged string-list from two string-list with a separator 'chr_join'
	'
	'	Arguments
	'	<String> strList1, 1st String-List to be merged
	'	<String> strList2, 2nd String-list to be merged
	'	<String> chr_join, A string character that is used as a separator
	'
	'***********************************************************************************************

	On Error Resume Next
	hs_merge_strlist = ""

	'*** Initialization ****************************************************************************
	Dim strRes

	If len(CStr(chr_join)) = 0 Then
		chr_join = ";"
	End If

	If IsArray(strList1) Then
		strList1 = ""
	End If

	If IsArray(strList2) Then
		strList2 = ""
	End If

	'*** Operations ********************************************************************************
	'--- Merge two string-lists --------------------------------------------------------------------
	strRes = strList1 & chr_join & strList2

	If Left(strRes, 1) = chr_join Then
		strRes = Mid(strRes, 2)
	End If

	If Right(strRes, 1) = chr_join Then
		strRes = Left(strRes, len(strRes) - 1)
	End If

	'--- Remove any duplicate element --------------------------------------------------------------
	strRes = hs_arr_remove_duplicate(Split(strRes, chr_join))

	'--- Rebuild string-list then release ----------------------------------------------------------
	hs_merge_strlist = Join(strRes, chr_join)

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function