Function IBase_create_filter_combo(ByVal arrFilter)
	'*** History ***********************************************************************************
	' 2020/08/30, BBS:	- First Release
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
	If Not (InStr(LCase(TypeName(arrFilter)), "variant") > 0 and UBound(arrFilter) > 0) Then
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