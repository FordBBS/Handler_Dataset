####################################################################################################
#                                                                                                  #
# Python, Dataset handler, BBS					   						                           #
#                                                                                                  #
####################################################################################################

#*** History ***************************************************************************************
# 2020/08/08, BBS:	- Move to BBS modules
#
#***************************************************************************************************

#*** Function Group List ***************************************************************************
# - Helper
# - Mathematics
# - Data Analysis
# - Sizing
# - Arrangement



#*** Library Import ********************************************************************************
# Operating system
import  os



#*** Function Group: Helper ************************************************************************
def hs_prep_AnyList(inpItem):
	if not isinstance(inpItem, list): return [inpItem]
	else: return IBase_list_remove_duplicate(inpItem, True)

def hs_prep_StrList(inpItem):
	if not (isinstance(inpItem, str) or isinstance(inpItem, list)): return []
	elif isinstance(inpItem, str):
		if inpItem == "": return []
		else: return [inpItem]
	else: return IBase_list_remove_duplicate(inpItem, True)

def hs_prep_RawFilter(listFilter):
	#*** Type of 'listFilter' validation *********************************************************** 
	if not isinstance(listFilter, list):
		listFilter = str(listFilter)
		if len(listFilter) > 0: return [[listFilter]]
		else: return []
	elif sum([not isinstance(x, list) for x in listFilter]) == len(listFilter): listFilter = [listFilter]

	#*** Remove all empty elements from each column ************************************************
	for idx, eachCol in enumerate(listFilter):
		eachCol 		= [str(x) for x in eachCol]
		listValid 	    = [1 if len(x) > 0 else 0 for x in eachCol]
		listFilter[idx] = IBase_get_reduced_list(eachCol, listValid)
	return listFilter

def hs_prep_convert_all_to_str(listInput, flg_case):
	#*** Input Validation **************************************************************************
	if not isinstance(listInput, list): return listInput

	#*** Initialization ****************************************************************************
	if not isinstance(flg_case, int): flg_case = 0
	if not flg_case in range(0, 3):   flg_case = 0 			# 0: no convert, 1:Upper, 2:Lower

	listRes = []

	#*** Operations ********************************************************************************
	for eachCol in listInput:
		if isinstance(eachCol, list):
			listRes.append([])
			listRes[len(listRes) - 1] = hs_prep_convert_all_to_str(eachCol, flg_case)
		elif flg_case == 1: listRes.append(str(eachCol).upper())
		elif flg_case == 2: listRes.append(str(eachCol).lower())
		else: listRes.append(str(eachCol))
	return listRes

def hs_convert_StrCase(inpItem, flg_case):
	# flg_case: 0: Convert to lower case, 1: Convert to upper case
	if not (isinstance(inpItem, str) or isinstance(inpItem, list)): return inpItem

	if not isinstance(flg_case, int): flg_case = 0
	if flg_case < 0 or flg_case > 1:  flg_case = 0

	if isinstance(inpItem, str):
		if flg_case == 0: return inpItem.lower()
		else: return inpItem.upper()

	inpItem = inpItem.copy()
	for idx, eachCol in enumerate(inpItem):
		if isinstance(eachCol, list): inpItem[idx] = hs_convert_StrCase(eachCol, flg_case)
		elif isinstance(eachCol, str):
			if flg_case == 0: inpItem[idx] = eachCol.lower()
			else: inpItem[idx] = eachCol.upper()

	return inpItem

def getconst_chr_path():
	RetVal = [chr(47), chr(92)] 		# ["/", "\"]
	return RetVal

def hs_fill_string_with_chr(strInput, chr_target, intStrLen):
	#*** Input Validation **************************************************************************
	chr_target = str(chr_target)

	if len(chr_target) == 0: return strInput
	try: intStrLen = int(intStrLen)
	except: return strInput

	#*** Initialization ****************************************************************************
	strResult = str(strInput)

	#*** Operations ********************************************************************************
	for idx in range(0, intStrLen - len(strResult)): strResult = strResult + chr_target
	return strResult



#*** Function Group: Mathematics *******************************************************************
def IUser_math_product(dataset):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Multiply each numerical element in a single column list 'dataset'
		Return False if 'dataset' is not a list or int
	
	[list] dataset, a numerical dataset to be obtained its product

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(dataset, list) or isinstance(dataset, int)): return False
	if isinstance(dataset, int): return dataset

	#*** Initialization ****************************************************************************
	intRes = 1

	#*** Operations ********************************************************************************
	for eachNum in dataset:
		if isinstance(eachNum, int): intRes = intRes*eachNum
	return intRes



#*** Function Group: Data Analysis *****************************************************************
def IBase_get_rootpath_of_list(listPath):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Return a root path which is the deepst folder among all paths in 'listPath'
		e.g. "ema/dataprocessing/Formulas/JAPAN/Blue_Book-2018_HD"
			 "ema/dataprocessing/Formulas/GLOBAL/UN-ECE_R49_Rev5-2008_74_EC"
			 "ema/dataprocessing/Formulas/INDIA/AIS137_Part7-A1_Draft_Dec2018"

			 return "ema/dataprocessing/Formulas"

	[list] listPath, A list of path to be analyzed

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listPath, list): return listPath

	#*** Initialization ****************************************************************************
	chr_path = getconst_chr_path()[1]
	listVal  = []
	shrtElem = "" 			# shortest string in 'listPath', used as a trial path
	strRes   = ""

	#*** Operations ********************************************************************************
	#--- Remove non-string element out from 'listPath' ---------------------------------------------
	for eachElement in listPath:
		if isinstance(eachElement, str):
			listVal.append(1)
			if shrtElem == "": shrtElem = eachElement
			elif len(shrtElem) > len(eachElement): shrtElem = eachElement
		else: listVal.append(0)

	listPath = IBase_get_reduced_list(listPath, listVal)

	#--- Path Analysis -----------------------------------------------------------------------------
	while shrtElem != "" and strRes == "":
		if sum([shrtElem in x for x in listPath]) == len(listPath): strRes = shrtElem
		else: shrtElem = os.path.split(shrtElem)[0]

	#--- Release -----------------------------------------------------------------------------------
	return strRes

def IBase_get_rootfolder_of_list(listPath):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Return a deepst common folder among all paths in 'listPath'
		e.g. "ema/dataprocessing/Scripts/EMA_Applications/ReportCreator/GenerateReport.csf"
			 "ema/dataprocessing_cus/Scripts/EMA_Applications/ReportCreator/GenerateReport.csf"

			 return "Scripts"

	[list] listPath, A list of path to be analyzed

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listPath, list): return listPath

	#*** Initialization ****************************************************************************
	chr_path = getconst_chr_path()[1]
	listVal  = []
	shrtElem = "" 			# shortest string in 'listPath', used as a trial path
	strRes   = ""

	#*** Operations ********************************************************************************
	#--- Remove non-string element out from 'listPath' ---------------------------------------------
	for eachElement in listPath:
		if isinstance(eachElement, str):
			listVal.append(1)
			if shrtElem == "": shrtElem = eachElement
			elif len(shrtElem) > len(eachElement): shrtElem = eachElement
		else: listVal.append(0)

	listPath = IBase_get_reduced_list(listPath, listVal)

	#--- Path Analysis -----------------------------------------------------------------------------
	shrtElem = os.path.split(shrtElem)[0]
	
	while shrtElem != "" and strRes == "":
		curFolder = os.path.split(shrtElem)[1]
		if sum([curFolder in x for x in listPath]) == len(listPath): strRes = curFolder
		else: shrtElem = os.path.split(shrtElem)[0]

	#--- Release -----------------------------------------------------------------------------------
	return strRes

def IBase_create_combination_fromlist(listInput):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Return a combination list of multiple columns input list 'listInput'
		e.g. [[1, 2], [3], [4, 5]] -> [[1, 1, 2, 2], [3, 3, 3, 3], [4, 5, 4, 5]]

	[list] listInput, A list to be combined

	'''

	#*** Input Validation **************************************************************************
	# Nothing to be validated

	#*** Initialization ****************************************************************************
	listInput  = hs_prep_AnyList(listInput)
	listWork   = []
	listRes    = []
	flg_single = False

	#*** Operations ********************************************************************************
	#--- Single List protection --------------------------------------------------------------------
	if sum([not isinstance(x, list) for x in listInput]) == len(listInput): listInput  = [listInput]

	#--- Optimzies 'listInput' ---------------------------------------------------------------------
	for idx in range(0, len(listInput)):
		curCol = listInput[idx]
		if not isinstance(curCol, list): curCol = [curCol]

		curCol = IBase_list_remove_duplicate(curCol, True)
		listWork.append(IBase_get_sorted_list([curCol], [0], [1])[0])
		listRes.append([])

	#--- Combination Creation ----------------------------------------------------------------------
	row_amnt    = IUser_math_product([len(x) for x in listWork])
	amnt_repeat = 0 		# Amount of continuous filling on each element per loop

	for idx in range(len(listRes)-1, -1, -1):
		if amnt_repeat == 0: amnt_repeat = 1
		else: amnt_repeat = amnt_repeat*len(listWork[idx + 1])

		amnt_loop = int(row_amnt/amnt_repeat/len(listWork[idx]))

		for cnt_loop in range(0, amnt_loop):
			for eachVal in listWork[idx]: listRes[idx].extend([eachVal]*amnt_repeat)

	#--- Release -----------------------------------------------------------------------------------
	return listRes

def IBase_list_remove_duplicate(listInput, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Return a list of all unique items in 'listInput'. Multi-column list is acceptable.

	[list] listInput, A list to be processed
	[bool] flg_case,  True: string case is considered differently, False: string case is ignored

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listInput, list) or len(listInput) <= 1: return listInput
	if not isinstance(flg_case, bool):  flg_case = True

	#*** Initialization ****************************************************************************
	flg_single = False
	listWork   = []
	listRes    = []
	listAppend = []

	#*** Operations ********************************************************************************
	#--- Checks whether 'listInput' is a single column list ----------------------------------------
	if sum([not isinstance(x, list) for x in listInput]) == len(listInput):
		flg_single = True
		listWork  = [listInput]
	else:
		for eachItem in listInput:
			if isinstance(eachItem, list): listWork.append(eachItem)
			else: listWork.append([eachItem])

	listWork = IBase_get_filled_list(listWork)

	#--- Duplicate item removal --------------------------------------------------------------------
	listVal = [0]*len(listWork[0])

	for idx in range(0, len(listWork[0])):
		tmpStr = ""

		for col in range(0, len(listWork)):
			curItem = listWork[col][idx]

			if curItem == None: curItem = ""
			else: curItem = str(curItem)
			if not flg_case: curItem = curItem.lower()

			tmpStr = tmpStr + "%" + curItem

		if not tmpStr in listAppend:
			listAppend.append(tmpStr)
			listVal[idx] = 1

	#--- Release -----------------------------------------------------------------------------------
	return  IBase_get_reduced_list(listWork, listVal)

def IBase_filter_combo_creation(listFilter):
	listFilter = hs_prep_AnyList(listFilter)
	effFilter  = []

	for idx in range(0, len(listFilter)):
		curCol = listFilter[idx]
		if not isinstance(curCol, list): curCol = [curCol]

		if sum(["all" in x.lower() for x in curCol]) == 0:
			curCol = IBase_list_remove_duplicate(curCol, True)
			effFilter.append(IBase_get_sorted_list([curCol], [0], [1])[0])

	effFilter = IBase_create_combination_fromlist(effFilter)
	return effFilter

def IBase_filter_str_singlelist(strInput, listFilter, flg_filt_type, flg_skip_all, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Checks 'strInput' with filter strings provided in 'listFilter'. Method of check depends
		on 'flg_filt_type' flag. An immediate return 'Passed' when 'all' is found can be selected.

	[str]  strInput,      A string to be checked
	[list] listFilter,    A list of filter string to be used as a constraint
	[int]  flg_filt_type, 0: All-Met returns True, 1: Some-Met returns True
	[bool] flg_skip_all,  True: Return True immediately when "all" is found in 'listFilter'
						  False: Disabled "all" Bypass
	[bool] flg_case, 	  True: string case is considered differently, False: string case is ignored

	'''

	#*** Input Validation **************************************************************************
	try:
		strInput = str(strInput)
		if len(strInput) == 0: return False
	except: return 101

	#*** Initialization ****************************************************************************
	if not (isinstance(flg_filt_type, int) and flg_filt_type in range(0, 2)): flg_filt_type = 0
	if not (isinstance(flg_skip_all, bool)): flg_skip_all = True
	if not (isinstance(flg_case, bool)):	 flg_case 	  = True

	listFilter = hs_prep_StrList(listFilter)
	listFilter = [str(x) for x in listFilter]
	kw_skip    = "all"
	flg_met    = 0

	#*** Operations ********************************************************************************
	#--- Converts all strings to lower case --------------------------------------------------------
	if not flg_case:
		strInput   = strInput.lower()
		listFilter = [x.lower() for x in listFilter]

	#--- Filter ------------------------------------------------------------------------------------
	for eachFilt in listFilter:
		if eachFilt.lower() == kw_skip and flg_skip_all: return True
		if eachFilt in strInput: flg_met = flg_met + 1
		if flg_met > 0 and flg_filt_type == 1: return True

	if flg_met == len(listFilter): return True
	else: return False

def IBase_filter_str_multilist(strInput, listFilter, flg_filt_type, flg_skip_all, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Checks 'strInput' with multiple columns filter provided in 'listFilter'. Method of check
		depends on 'flg_filt_type' flag. An immediate return 'Passed' when 'all' is found can be selected.

		Each element in each column provided in 'listFilter' will be created as combination tree.
		A filtering considers 'Met' when 'strInput' meets all columns in a tree.

	[str]  strInput,      A string to be checked
	[list] listFilter,    A list of filter string to be used as a constraint
	[int]  flg_filt_type, 0: All-Met returns True, 1: Some-Met returns True
	[bool] flg_skip_all,  True: Return True immediately when "all" is found in 'listFilter'
						  False: Disabled "all" Bypass
	[bool] flg_case, 	  True: string case is considered differently, False: string case is ignored

	'''

	#*** Input Validation **************************************************************************
	try:
		strInput = str(strInput)
		if len(strInput) == 0: return False
	except: return 101
	
	#*** Initialization ****************************************************************************
	if not (isinstance(flg_filt_type, int) and flg_filt_type in range(0, 2)): flg_filt_type = 0
	if not isinstance(flg_skip_all, bool): 	flg_skip_all = True
	if not isinstance(flg_case, bool):	 	flg_case 	 = True

	listFilter = hs_prep_RawFilter(listFilter)
	if len(listFilter) == 0: return False  		# no filter provided, no chance to return True

	effFilter  = []

	#*** Operations ********************************************************************************
	#--- Character Case Handler --------------------------------------------------------------------
	if not flg_case:
		listFilter = hs_convert_StrCase(listFilter, 0)
		strInput   = hs_convert_StrCase(strInput, 0)

	#--- Pre-Filtering -----------------------------------------------------------------------------
	for eachCol in listFilter:
		listVal = []
		flg_all = False
		
		if not isinstance(eachCol, list): eachCol = [eachCol]
		if len(eachCol) == 0: listVal.append(0)
		else:
			for eachElement in eachCol:
				if eachElement.lower() == "all":
					flg_all = True

					if flg_skip_all == True: break
					else: listVal.append(0)
				elif eachElement in strInput: listVal.append(1)
				else: listVal.append(0)
			
			else:
				tmpReduced = IBase_get_reduced_list(eachCol, listVal)
				if len(tmpReduced) > 0: effFilter.append(tmpReduced)
				elif not (len(eachCol) == 1 and flg_all): return False
				# If there is at least one empty column, it means 'strInput' can't meet any element
				# in that column!

	#--- Release if there is no any filter left ----------------------------------------------------
	if len(effFilter) == 0: return flg_all

	#--- Create Tree-Combination filter of effective filter elements -------------------------------
	effFilter = IBase_create_combination_fromlist(effFilter)

	#--- Post-Filtering ----------------------------------------------------------------------------
	for idx in range(0, len(effFilter[0])):
		for col in range(0, len(effFilter)):
			if not effFilter[col][idx] in strInput:
				if flg_filt_type == 0: return False
				break
		else:
			if flg_filt_type == 1: return True

	if flg_filt_type == 1: return False
	else: return True

def IUser_filter_get_listvalid(listInput, listFilter, flg_filt_type, flg_skip_all, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Return a 'listValid' consists of 0 or 1 which represents the filter result of each element.

	[list] listInput,     Either String or List of base data, If multi-column list is provided, 
							first column is always used for filtering.
	[list] listFilter,    A list of filter string to be used as a constraint
	[int]  flg_filt_type, 0: All-Met returns True, 1: Some-Met returns True
	[bool] flg_skip_all,  True: Return True immediately when "all" is found in 'listFilter'
						  False: Disabled "all" Bypass
	[bool] flg_case, 	  True: string case is considered differently, False: string case is ignored

	'''

	#*** Input Validation **************************************************************************
	if (isinstance(listInput, str) or isinstance(listInput, list)) and len(listInput) == 0:
		return []

	#*** Initialization ****************************************************************************
	listInput = hs_prep_AnyList(listInput)

	#*** Operations ********************************************************************************
	#--- Checks whether 'listInput' is a multi-column list -----------------------------------------
	if sum([isinstance(x, list) for x in listInput]) == len(listInput): listInput = listInput[0]

	#--- Filtering ---------------------------------------------------------------------------------
	listValid = [0]*len(listInput)

	for idx, eachStr in enumerate(listInput):
		if IBase_filter_str_multilist(eachStr, listFilter, flg_filt_type, flg_skip_all, flg_case):
			listValid[idx] = 1
	return listValid



#*** Function Group: Sizing ************************************************************************
def IBase_remove_target_from_list(listDataset, listTargetValue):
	#*** Documentation *****************************************************************************
	'''Help,

		Remove any value provided in 'listTargetvalue' out from every column or index of 'listDataset'
		Target value shall be provided correctly in term of Upper/Lower cases. Each target value will
		be converted to string for the proper comparison.

	[list] listDataset,	    A list of dataset to be used.
	[list] listTargetValue, A list of value to be removed from 'listDataset'
	
	'''

	#*** Input Validation **************************************************************************
	# Nothing to be initialized

	#*** Initialization ****************************************************************************
	listTargetValue = hs_prep_AnyList(listTargetValue)
	listRes 		= []

	#*** Operations ********************************************************************************
	#--- 'listTargetValue' preparation -------------------------------------------------------------
	listTargetValue = hs_prep_convert_all_to_str(listTargetValue, 0)

	#--- Case: 'listDataset' is not a List ---------------------------------------------------------
	if not isinstance(listDataset, list):
		if str(listDataset) in listTargetValue: return ""
		else: return listDataset

	#--- Case: 'listDataset' is List ---------------------------------------------------------------
	for eachElement in listDataset:
		if isinstance(eachElement, list):
			listRes.append([])
			listRes[len(listRes) - 1] = IBase_remove_target_from_list(eachElement, listTargetValue)
		elif not str(eachElement) in listTargetValue: listRes.append(eachElement)

	#--- Release -----------------------------------------------------------------------------------
	return listRes

def IBase_remove_empty_list(listDataset):
	#*** Documentation *****************************************************************************
	'''Help,

		Remove any empty list in 'listDataset'

	[list] listDataset,	    A list of dataset to be used.
	
	'''

	#*** Input Validation **************************************************************************
	# Nothing to be initialized

	#*** Initialization ****************************************************************************
	listRes = []

	#*** Operations ********************************************************************************
	for eachElement in listDataset:
		if isinstance(eachElement, list):
			if len(eachElement) > 0:
				curRes = IBase_remove_empty_list(eachElement)
				if len(curRes) > 0:
					listRes.append([])
					listRes[len(listRes) - 1] = curRes
		else: listRes.append(eachElement)

	#--- Release -----------------------------------------------------------------------------------
	return listRes

def IBase_get_reduced_list(listDataset, listFlagValid):
	#*** Documentation *****************************************************************************
	'''Help,

		return a reduced list of 'listDataset' based on referecne valid flag in 'listFlagValid'
		Any element that has 0, False, or empty flag value will be removed from 'listDataset'

	[list] listDataset,	  A list of dataset to be reduced. Multi-dimension is acceptable
						  If each column in 'listDataset' has different length, the maximum length
						  found will be used and 'None' will be assumed to all missing index in
						  other columns
						  e.g. [["A", "B"], [1], ["AB", "AC", "AD"]]
						  compiled to [["A", "B", ""], [1, "", ""], ["AB", "AC", "AD"]]

	[list] listFlagValid, A list of validity flag. If this list has a different dimension.
						  Auto-Alignment will be adapted with "True" as default flag value for thos
						  position.
	
	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listDataset, list) and len(listDataset) > 0 \
			and isinstance(listFlagValid, list)):
		return listDataset

	#*** Initialization ****************************************************************************
	if not isinstance(listDataset[0], list): listDataset = [listDataset]
	listFlg = []
	listRes = [[]]*len(listDataset)
	row_max = 0

	#*** Operations ********************************************************************************
	#--- Target Dataset Conditioning ---------------------------------------------------------------
	row_max = max([len(eachCol) for eachCol in listDataset])

	for idx in range(0, len(listDataset)):
		listEmpty = [None]*(row_max - len(listDataset[idx]))
		listDataset[idx].extend(listEmpty)

	#--- Reduce Flag List preparation --------------------------------------------------------------
	for idx in range(0, len(listDataset[0])):
		curFlg = True

		if idx < len(listFlagValid):
			curVal = listFlagValid[idx]
			if isinstance(curVal, int) and curVal == 0: curFlg = False
			elif isinstance(curVal, bool): curFlg = curVal

		listFlg.append(curFlg)

	#--- Reduce Dataset ----------------------------------------------------------------------------
	for idx in range(0, len(listRes)):
		listTmp = [val for val, flagValid in zip(listDataset[idx], listFlg) if flagValid]
		listRes[idx] = listTmp

	#--- Release -----------------------------------------------------------------------------------
	if len(listRes) == 1: return listRes[0]
	else: return listRes

def IBase_get_filled_list(listDataset):
	#*** Documentation *****************************************************************************
	'''Help,

		return a filled list of 'listDataset' by searching for the first-found longest column in 
		'listDataset' and fill other columns with their own information. If other columns has more
		than one unique value originally, the last row information will be used for filling.

	[list] listDataset,	  A list of dataset to be filled.

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listDataset, list): return [listDataset]

	#*** Initialization ****************************************************************************
	flg_valid = False
	row_max   = 0
	listRes   = []

	#*** Operations ********************************************************************************
	#--- Checks whether there is any list in 'listDataset' -----------------------------------------
	for eachItem in listDataset:
		if isinstance(eachItem, list): flg_valid = True
		else: eachItem = [eachItem]

		if row_max < len(eachItem): row_max = len(eachItem)
		listRes.append(eachItem)

	if not flg_valid: return [[x] for x in listDataset]

	#--- Fill information --------------------------------------------------------------------------
	for idx, eachCol in enumerate(listRes):
		col_len = len(eachCol)
		if col_len > 0: 
			usedVal = eachCol[col_len-1]
			listRes[idx].extend([usedVal]*(row_max - col_len))

	#--- Release -----------------------------------------------------------------------------------
	return listRes



#*** Function Group: Arrangement *******************************************************************
def IBase_get_sorted_list(listInput, sortColumn, sortDirection):
	#*** Documentation *****************************************************************************
	'''Help,

		return a sorted 'listInput'
		sorting column is referred to 'sortColumn' in order
		sorting directory of each column is referred to 'sortDirection'
	
	[list] listInput,  		a list to be sorted
	[list] sortColumn, 		a list contains a column index to be sorted in order. A length of this
							list must be less than or equal to listInput.
							If 'sortColumn' is empty or not provided, every column in 'listInput'
							will be sorted with 'sortDirection' or default direction
	[list] sortDirection, 	a list contains a flag of sorting direction for each column index
							provided in 'sortColumn'. A length of 'sortDirection' must be less than
							or equal to a length of 'sortColumn'. Default dir is "Ascending"
							1: "Ascending", 2: "Descending"
	
	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listInput, list) and len(listInput) > 0): return listInput

	#*** Initialization ****************************************************************************
	#--- Preparation: sortColumn -------------------------------------------------------------------
	if not (isinstance(sortColumn, list)):
		if isinstance(sortColumn, int) and sortColumn < len(listInput): sortColumn = [sortColumn]
		else: sortColumn = []

	if len(sortColumn) == 0: sortColumn = list(x for x in range(0, len(listInput)))

	listTmp = []
	for eachIndex in sortColumn:
		if eachIndex < len(listInput): listTmp.append(eachIndex)

	sortColumn = list(listTmp)
	if len(sortColumn) >= len(listInput): sortColumn = sortColumn[:len(listInput)] 

	#--- Preparation: sortDirection ----------------------------------------------------------------
	if not (isinstance(sortDirection, list)):
		if isinstance(sortDirection, int): sortDirection = [sortDirection]
		else: sortDirection = []

	nDir = len(sortDirection)
	nCol = len(sortColumn)
	
	if nDir == 0: 	sortDirection = list(1 for x in range(0, nCol))
	if nDir > nCol: sortDirection = sortDirection[:nCol]
	if nDir < nCol: sortDirection.extend([1]*(nCol - nDir))

	sortDirection = list(x != 1 for x in sortDirection)

	#--- General -----------------------------------------------------------------------------------
	listRes = listInput.copy()
	listIdx = [x for x in range(0, len(listRes[0]))] 		# 0, 1, 2, ..., n

	#*** Operations ********************************************************************************
	for idx, idxCol in enumerate(sortColumn):
		# Group to (x, y) then Sort by x; where x is Value and y is Latest Index
		tmpList = list([x,y] for x, y in zip(listRes[idxCol], listIdx))
		tmpList = sorted(tmpList, reverse=sortDirection[idx])

		# Extract current Index list into 'sortIdx' for applying proper result on listRes
		sortIdx = [x[1] for x in tmpList]

		# Correct indices of duplicate values if reverse option was enabled
		# BBS: I always want to have an ascending sort method when it comes to an index sorting
		if sortDirection[idx]:
			sortVal  = [x[0] for x in tmpList]
			flg_corr = True
			flg_dupl = 0
			lcl_cnt  = 0

			while flg_corr:
				if lcl_cnt == 0: preVal = sortVal[lcl_cnt]
				else:
					curVal = sortVal[lcl_cnt]

					if curVal == preVal:
		 				if flg_dupl == 0: idx_sta = lcl_cnt - 1
		 				idx_end  = lcl_cnt
		 				flg_dupl = 1

					elif flg_dupl > 0:
		 				tmpCorrIdx = sortIdx[idx_sta:idx_end + 1]
		 				tmpCorrIdx = sorted(tmpCorrIdx)

		 				for cnt_corr in range(0, len(tmpCorrIdx)):
		 					sortIdx[cnt_corr + idx_sta] = tmpCorrIdx[cnt_corr]

					preVal = curVal

				lcl_cnt = lcl_cnt + 1
				if lcl_cnt >= len(sortVal): flg_corr = False

		# Compile sorting result to each column of 'listRes'
		for resIdx, resCol in enumerate(listRes):
			listRes[resIdx] = [resCol[sortIdx[x]] for x in range(0, len(resCol))]

	return listRes

def IBase_get_arranged_list(listInput, value_missing):
	#*** Documentation *****************************************************************************
	'''Help,

		Return rearranged list of 'listInput' if there is any missing value, 'value_missing' will
		be filled. If 'value_missing' is not provided, "" will be used

		e.g. listInput     = [[1,2], [2, 4], [3, 6], 5]
			 value_missing = 0
			 Result: [[1,2,3,5], [2,4,6,0]]

			 listInput     = [[1,2,3], [2,4,6], [3,6,9], 5]
			 value_missing = ""
			 Result: [[1,2,3,5], [2,4,6,""], [3,6,9,""]]
	
	[list] listInput,    A list to be rearranged
	[any] value_missing, Default value for filling missing value. It can be any type. If list type
						 is provided, the first value in that list is used

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listInput, list) and len(listInput) > 0): return listInput

	#*** Initialization ****************************************************************************
	if isinstance(value_missing, list): value_missing = value_missing[0]
	if len(str(value_missing)) == 0: 	value_missing = ""

	listRes = []
	n_list  = 0

	#*** Operations ********************************************************************************
	#--- Determine longest length among 'listInput' ------------------------------------------------
	for eachElement in listInput:
		if isinstance(eachElement, list) and len(eachElement) > n_list: n_list = len(eachElement)

	#--- Create 'listRes' --------------------------------------------------------------------------
	for idx in range(0, n_list): listRes.append([])

	#--- Rearrange ---------------------------------------------------------------------------------
	for eachElement in listInput:
		if not isinstance(eachElement, list): eachElement = [eachElement]
		if len(eachElement) < n_list: eachElement.extend([value_missing]*(n_list - len(eachElement)))
		
		for col in range(0, n_list): listRes[col].append(eachElement[col])

	#--- Release -----------------------------------------------------------------------------------
	return listRes


