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