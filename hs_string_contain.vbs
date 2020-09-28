Function hs_string_contain(ByVal strBase, ByVal arrString, ByVal flg_truecase, ByVal flg_case)
	'### hs_string_contain, BBS ####################################################################
	'*** History ***********************************************************************************
	' 2020/07/31, BBS:	- First release
	'
	'***********************************************************************************************

	On Error Resume Next

	Dim myName: myName = "hs_string_contain: "
	Dim flg_msg, flg_msg_storage

	flg_msg 		= False
	flg_msg_storage = 0
	hs_string_contain = False

	'*** Initialization ****************************************************************************
	'--- Flag, Case of return True: 0= All met, 1= Some met ----------------------------------------
	If IsNumeric(flg_truecase) Then
		flg_truecase = CInt(flg_truecase)
	Else
		flg_truecase = 0
	End If

	'--- Flag, Case-sensitive: True= Case is sensitive, False= Vice versa --------------------------
	If LCase(CStr(flg_case)) = "false" or LCase(CStr(flg_case)) = "0" Then
		flg_case = False
	Else
		flg_case = True
	End If

	'--- General -----------------------------------------------------------------------------------
	Dim idx, curString, arrReg

	strBase = CStr(strBase)
	If Not flg_case Then strBase = LCase(CStr(strBase))

	'*** Operations ********************************************************************************
	'--- Ensure base string has something ----------------------------------------------------------
	If len(strBase) = 0 Then
		Exit Function
	End If

	'--- Condition target strings to be searched ---------------------------------------------------
	If IsArray(arrString) Then
		arrReg = arrString
	Else
		arrReg = Split(arrString, ";")
	End If

	'--- Filtering ---------------------------------------------------------------------------------
	For idx = 0 to UBound(arrReg)
		curString = CStr(arrReg(idx))
		If Not flg_case Then curString = LCase(curString)
		
		hs_string_contain = InStr(strBase, curString) > 0
		If hs_string_contain and flg_truecase = 1 Then Exit Function
		If Not hs_string_contain and flg_truecase = 0 Then Exit Function
	Next

	'*** Error Handler *****************************************************************************
	If Err.Number <> 0 Then
		PoiSendMessage 4, myName & "finished with Error number " & CStr(Err.Number), flg_msg_storage
		Err.Clear
	End If
End Function
