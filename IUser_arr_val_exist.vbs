Function IUser_arr_val_exist(ByVal arrInput, ByVal tarValue, ByVal flg_case)
	'*** History ***********************************************************************************
	' 2020/09/19, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return an index where 'tarValue' is found on 'arrInput'
	' 	-1 is returned if 'tarValue' doesn't exist on 'arrInput'
	'	
	'	Argument(s)
	'	<Array> arrInput,	Array to be searched
	'	<Str>	tarValue,	Target value in any format but Array, it will be converted to String anyway
	'	<Bool>	flg_case,	False: Case doesn't matter, True: Case does matter
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IUser_arr_val_exist = -1

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(arrInput) or len(CStr(tarValue)) = 0 Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim idx

	tarValue = CStr(tarValue)

	If Not (LCase(CStr(flg_case)) = "true" or LCase(CStr(flg_case)) = "false") Then
		flg_case = True
	End If

	If Not flg_case Then
		tarValue = LCase(tarValue)
	End If

	'*** Operations ********************************************************************************
	For idx = 0 to UBound(arrInput)
		If Not IsArray(arrInput(idx)) Then
			If CStr(arrInput(idx)) = tarValue or _
			 	(Not flg_case and LCase(CStr(arrInput(idx))) = tarValue) Then
			 	
			 	IUser_arr_val_exist = idx
			 	Exit For
			 End If
		End If
	Next

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
