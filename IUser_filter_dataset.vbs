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