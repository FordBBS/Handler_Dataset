'*** History ***************************************************************************************
' 2020/08/30, BBS:	- First Release
' 					- Imported all mandatory materials
' 2020/09/19, BBS: 	- Updated 'IUser_filter_dataset', 'IBase_create_filter_combo', 'hs_arr_append'
' 					- Imported 'hs_merge_strlist'
'
'***************************************************************************************************

Function IUser_filter_dataset(ByVal inpDataset, ByVal inpFilter, ByVal flg_mode, ByVal flg_case)
	'*** History ***********************************************************************************
	' 2020/08/30, BBS:	- First Release
	' 2020/09/19, BBS: 	- Improved
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return an Array of valid element from 'inpDataset'
	'	
	'	Argument(s)
	'	<Array> inpDataset,	Array of Non-Array to be filtered
	'	<Array> inpFilter, 	Array of filter elements (Multi-Column filter is acceptable)
	'	<Int>	flg_mode,	0: All-Met True, 1: Some-Met True, 2: All-Met False, 3: Some-Met False
	'	<Bool>	flg_case,	False: Case doesn't matter, True: Case does matter
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IUser_filter_dataset = Array()

	'*** Pre-Validation ****************************************************************************
	If IsArray(inpDataset) and UBound(inpDataset) = 0 Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim RetVal, cnt1
	Dim arrDataset, arrFilter, arrFilterTemp, arrResult, thisElement

	If Not IsNumeric(flg_mode) Then
		flg_mode = 0
	Else
		flg_mode = CInt(flg_mode)
	End If
	If flg_mode < 0 or flg_mode > 3 Then flg_mode = 0

	If Not (LCase(CStr(flg_case)) = "true" or LCase(CStr(flg_case)) = "false") Then
		flg_case = True
	End If

	arrResult = Array()

	'*** Operations ********************************************************************************
	'--- Prepare 'arrDataset' ----------------------------------------------------------------------
	If IsArray(inpDataset) Then
		For Each thisElement in inpDataset
			If Not IsArray(thisElement) Then
				thisElement = CStr(thisElement)
			End If

			Call hs_arr_append(arrDataset, thisElement)
		Next
	Else
		arrDataset = Array(CStr(inpDataset))
	End If

	'--- Prepare 'arrFilter' -----------------------------------------------------------------------
	If IsArray(inpFilter) Then
		cnt1 = -1

		For Each thisElement in inpFilter
			If Not IsArray(thisElement) Then
				cnt1 = cnt1 + 1
				Call hs_arr_append(arrFilterTemp, Array(thisElement))
			Else
				Call hs_arr_append(arrFilterTemp, thisElement)
			End If
		Next

		'All elements in 'inpFilter' are not Array. Therefore 'inpFilter' is a single-column filter
		If cnt1 = UBound(inpFilter) Then
			arrFilter = Array(inpFilter)

		'Some element in 'inpFilter' is an Array. Therefore 'inpFilter' is a multi-column filter
		Else
			arrFilter = arrFilterTemp
		End If
	Else
		arrFilter = Array(Array(CStr(inpFilter)))
	End If

	arrFilter = IBase_create_filter_combo(arrFilter)

	'--- Filtering ---------------------------------------------------------------------------------
	For Each thisElement in arrDataset
		If Not IsArray(thisElement) Then
			RetVal = IBase_filter_single_element(thisElement, arrFilter, flg_mode, flg_case)
		Else
			RetVal = True
		End If

		If RetVal Then
			Call hs_arr_append(arrResult, thisElement)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	IUser_filter_dataset = arrResult

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IBase_filter_single_element(ByVal strInput, ByVal arrFilter, ByVal flg_mode, ByVal flg_case)
	'*** History ***********************************************************************************
	' 2020/08/30, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return filtering result of 'strInput' versus 'arrFilter' either True or False
	'	
	'	Argument(s)
	'	<String> strInput,	String to be filtered
	'	<Array>  arrFilter, Combination filter array (created by 'IBase_create_filter_combo')
	'	<Int>	 flg_mode,	0: All-Met True, 1: Some-Met True, 2: All-Met False, 3: Some-Met False
	'	<Int>	 flg_case,	0: Case doesn't matter, 1: Case does matter
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IBase_filter_single_element = False

	'*** Pre-Validation ****************************************************************************
	strInput = CStr(strInput)
	If len(strInput) = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim RetVal, cnt1, cnt2, flg_break, thisFilter

	If Not flg_case Then strInput = LCase(strInput)
	If flg_mode = 0 or flg_mode = 3 Then
		RetVal = True
	Else
		RetVal = False
	End If

	'*** Operations ********************************************************************************
	'--- Filtering ---------------------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrFilter(0))		'Each row
		flg_break = False

		For cnt2 = 0 to UBound(arrFilter)		'Each column
			thisFilter = arrFilter(cnt2)(cnt1)
			
			If Not flg_case Then thisFilter = LCase(thisFilter)
			If InStr(strInput, thisFilter) = 0 Then
				flg_break = True
				Exit For
			End If
		Next

		If flg_break Then 	'All-met is not possible any longer
			If flg_mode mod 2 = 0 Then
				If flg_mode = 0 Then
					RetVal = False
				Else
					RetVal = True
				End If

				Exit For
			End If
		Else 				'Some-met is fulfilled
			If flg_mode mod 2 = 1 Then
				If flg_mode = 1 Then
					RetVal = True
				Else
					RetVal = False
				End If

				Exit For
			End If
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	IBase_filter_single_element = RetVal

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IBase_create_filter_combo(ByVal arrFilter)
	'*** History ***********************************************************************************
	' 2020/08/30, BBS:	- First Release
	' 2020/09/19, BBS: 	- Improved
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return combination array of 'arrFilter' for filtering purpose
	'	
	'	Argument(s)
	'	<Array>  arrFilter, Array of filter elements (Multi-Column filter is acceptable)
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IBase_create_filter_combo = Array()

	'*** Pre-Validation ****************************************************************************
	If Not (IsArray(arrFilter) and UBound(arrFilter) > 0) Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim row_size, cnt1, cnt2, cnt3, cnt4, cnt_sw, thisCol
	Dim arrRes(), arrTmp()
	Redim Preserve arrTmp(0), arrRes(UBound(arrFilter))

	row_size = 1
	cnt_sw	 = 0

	'*** Operations ********************************************************************************
	'--- Get total amount of row -------------------------------------------------------------------
	For Each thisCol in arrFilter
		row_size = row_size*(UBound(thisCol) + 1)
	Next

	If row_size < 1 Then Exit Function
	row_size = row_size - 1

	'--- Combination creation ----------------------------------------------------------------------
	For cnt1 = UBound(arrFilter) to 0 Step -1
		cnt3 = 0
		cnt4 = 1
		Erase arrTmp
		Redim Preserve arrTmp(0)

		If cnt_sw = 0 Then
			cnt_sw = 1
		Else
			cnt_sw = cnt_sw*(1 + UBound(arrFilter(cnt1 + 1)))
		End If

		For cnt2 = 0 to row_size
			If cnt4 > cnt_sw Then
				cnt4 = 1
				cnt3 = cnt3 + 1
				If cnt3 > UBound(arrFilter(cnt1)) Then cnt3 = 0
			End If

			Call hs_arr_append(arrTmp, arrFilter(cnt1)(cnt3))
			cnt4 = cnt4 + 1
		Next

		arrRes(cnt1) = arrTmp
	Next

	'--- Release -----------------------------------------------------------------------------------
	IBase_create_filter_combo = arrRes

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function hs_arr_append(ByRef arrInput, ByVal tarValue)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First release
	' 2020/08/25, BBS:  - Implemented handler for Non-Array 'arrInput'
	' 2020/09/19, BBS: 	- Improved
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Append 'tarValue' to target array provided as 'arrInput', 'arrInput' can be only a single
	'	column array only
	'
	'	Argument(s)
	'	<Array>  arrInput, Base array to be appended 'tarValue'
	'	<Any> 	 tarValue, Desire value to be appended to 'arrInput'
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Initialization ****************************************************************************
	' Nothing to be initialized

	'*** Operations ********************************************************************************
	'--- Ensure 'arrInput' is Array type before doing appending ------------------------------------
	If Not IsArray(arrInput) Then
		arrInput = Array(arrInput)
	End If

	'--- Appending ---------------------------------------------------------------------------------
	If Not (UBound(arrInput) = 0 and LCase(TypeName(arrInput(0))) = "empty") Then
		Redim Preserve arrInput(UBound(arrInput) + 1)
	End If

	arrInput(UBound(arrInput)) = tarValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

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