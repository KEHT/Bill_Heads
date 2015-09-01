#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=bill_heads_8eA_icon.ico
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=Application producing Bill Head cover sheets
#AutoIt3Wrapper_Res_Fileversion=1.0.1.0
#AutoIt3Wrapper_Res_ProductVersion=1.0
#AutoIt3Wrapper_Res_LegalCopyright=U.S. GPO
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=ProductName|Bill Heads Generator
#AutoIt3Wrapper_Res_Field=OriginalFilename|Bill_Heads.exe
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ 07/01/2014 - sjohnson@gpo.gov - Alpha version (0.90) BILL_HEADS to process speeches
;~ 07/31/2014 - sjohnson@gpo.gov - Beta version (0.99) BILL_HEADS to process HOR daily activities is ready
;~ 09/17/2014 - sjohnson@gpo.gov - Beta 2 version (0.9.9.3) BILL_HEADS: Fixed long titles in Excel, em dash problems and S/O
;~ 09/19/2014 - sjohnson@gpo.gov - Beta 2 version (0.9.10.1) BILL_HEADS: Created intermediate screen to choose bills for Excel
;~ 09/23/2014 - sjohnson@gpo.gov - Beta 2 version (0.9.10.5) BILL_HEADS: Fixed Combobox scrolling selection problem, blank names, etc.
;~ 10/01/2014 - sjohnson@gpo.gov - Release 1 version (1.0.0.0) BILL_HEADS: Most likely fixed regex for Committee of the Whole DOC files
;~ 07/02/2015 - sjohnson@gpo.gov - Bug Fix version (1.0.0.1) BILL_HEADS: Fixing congressmen adding to the list of COTW (unknown bug, may re-appear)
;~ 08/24/2015 - sjohnson@gpo.gov - Congressmen addition to COTW bug is squashed, hopefully.
;~ 09/01/2105 - sjohnson@gpo.gov - Removed counts for bills and added COTW bills to the Excel Spreadsheet
#include <file.au3>
#include <ClipBoard.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <DateTimeConstants.au3>
#include <ProgressConstants.au3>
#include <word.au3>
#include <StringConstants.au3>
#include <objDictonary.au3>
#include <Excel.au3>
#include <FontConstants.au3>
#include <GUIListBox.au3>
#include <GuiComboBox.au3>

Opt("GUIOnEventMode", 1)

Dim $yorno = 7
Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $cInputFileName, $cInputFile, $sInputFileText
Global $hProcHeadsButton, $hMrChairButton, $hDefault_Button, $hApply_Button, $hWholeCommButton, $hGenLeaveButton, $hMainGUI, $hCombo, $hListbox, $hPrintDocButton, _
		$hCancelDocButton, $hBillGUI = 9999, $hBillButton = 9999, $hCongGUI = 8888, $bSelected = False ; Predeclare the variables with dummy values to prevent firing the Case statements

Global $cInputFolderDefault = "\\alpha3\E\CR\FM"
Global $cOutputFolderDefault = "\\alpha3\E\RECSCAN\TofA"
Global $cInputFolder, $cOutputFolder

Global $tipmsg = "PLEASE WAIT..."

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate, $ProcHeadsButton, $MrChairButton, $hWCbillsNum, $progressbar, $inputFolder, $outputFolder, $SelectedBill, $Radios[0], _
		$CheckBoxes[0][2]
Dim $toWholeCommittee, $toGenLeave

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()

	$hMainGUI = GUICreate("Bill Heads Program v" & _GetVersion(), 350, 300)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked
	$tab = GUICtrlCreateTab(5, 5, 340, 290)

	; tab 0
	$tab0 = GUICtrlCreateTabItem("Main")
	GUICtrlCreateLabel("Choose Date Below:", 15, 40, 300)
	$LocalDate = _DateAdd('d', -1, _NowCalcDate())
	$Date = GUICtrlCreateMonthCal($LocalDate, 65, 70, 220, 140, $MCS_NOTODAY)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$DateSelected = GUICtrlCreateLabel("Date Selected: " & $LocalDate, 15, 220, 300)
	$hProcHeadsButton = GUICtrlCreateButton("Process Heads", 35, 240, 120)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hWCbillsNum = GUICtrlCreateLabel("", 230, 220, 50, 30, $SS_NOTIFY)
	$hMrChairButton = GUICtrlCreateButton("Print Mr. Chair DOC", 195, 240, 120)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hMrChairButton, $GUI_DISABLE)
	$progressbar = GUICtrlCreateProgress(35, 275, 280, 10, $PBS_SMOOTH)

	; tab 1
	$tab1 = GUICtrlCreateTabItem("Settings")
	GUICtrlCreateLabel("Input Folder:", 15, 40, 300)
	$inputFolder = GUICtrlCreateInput("", 15, 65, 320, 20)
	$cInputFolder = fuGetInputOutput("input", $cInputFolderDefault)
	GUICtrlSetData($inputFolder, $cInputFolder)

	$hOutputFolderLabel = GUICtrlCreateLabel("Output Folder:", 15, 100, 300)
	$outputFolder = GUICtrlCreateInput("", 15, 125, 320, 20)
	$cOutputFolder = fuGetInputOutput("output", $cOutputFolderDefault)
	GUICtrlSetData($outputFolder, $cOutputFolder)
	GUICtrlSetState($outputFolder, $GUI_HIDE)
	GUICtrlSetState($hOutputFolderLabel, $GUI_HIDE)

	$hDefault_Button = GUICtrlCreateButton("Default", 15, 160, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hApply_Button = GUICtrlCreateButton("Apply", 100, 160, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function


	GUICtrlCreateTabItem("") ; end tabitem definition

	GUISetState()

	; Run the GUI until the dialog is closed
	While 1
		Sleep(10)
	WEnd

EndFunc   ;==>fuMainGUI

Func fuChooseBillHeadsGUI()
	Local $iBillCount = _ObjDictCount($toGenLeave)
	$iBillCount += _ObjDictCount($toWholeCommittee)
	$hBillGUI = GUICreate("Choose Quantities & Bills", 300, $iBillCount * 20 + 100, Default, Default, Default, Default, $hMainGUI)
	$hWordingLabel = GUICtrlCreateLabel("You have " & $iBillCount & " Bill Head(s). Which ones do you want to print?", 0, 8, 295, 30, $SS_CENTER)
	GUICtrlSetFont($hWordingLabel, Default, $FW_NORMAL, $GUI_FONTITALIC, "Arial")
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the secondary GUI [X] is clicked

	ReDim $CheckBoxes[$iBillCount][1]
	;_ObjDictList($toGenLeave)
	Local $aP = 50, $aX = 0
	For $myHR In $toGenLeave
		$CheckBoxes[$aX][0] = GUICtrlCreateCheckbox( _ObjDictGetValue($toGenLeave, $myHR), 60, $aP, 150, 17)
		$aX += 1
		$aP += 16
	Next

	For $myHR In $toWholeCommittee.Keys
		$CheckBoxes[$aX][0] = GUICtrlCreateCheckbox($myHR, 60, $aP, 150, 17)
		$aX += 1
		$aP += 16
	Next
	$hGenLeaveButton = GUICtrlCreateButton("Select Bills", 35, $aP + 16, 90)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($CheckBoxes[0][0], $GUI_CHECKED)
	GUISetState(@SW_SHOW)
EndFunc   ;==>fuChooseBillHeadsGUI


Func fuWholeCommGUI()
	Local $iBillCount = _ObjDictCount($toWholeCommittee)
	$hBillGUI = GUICreate("Choose a Bill", 300, $iBillCount * 20 + 100, Default, Default, Default, Default, $hMainGUI)
	$hWordingLabel = GUICtrlCreateLabel("You have " & $iBillCount & " Word Doc(s). Which one do you want to print?", 5, 8, 295, 25, $SS_CENTER)
	GUICtrlSetFont($hWordingLabel, Default, $FW_NORMAL, $GUI_FONTITALIC, "Arial")
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the secondary GUI [X] is clicked
	ReDim $Radios[$iBillCount]

	Local $aP = 50, $aX = 0
	For $myHR In $toWholeCommittee
		$Radios[$aX] = GUICtrlCreateRadio( _ObjDictGetValue($toWholeCommittee, $myHR)[2], 50, $aP, 150, 17)
		$aX += 1
		$aP += 16
	Next
	$hWholeCommButton = GUICtrlCreateButton("Select Bill", 35, $aP + 16, 90)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($Radios[0], $GUI_CHECKED)
	GUISetState(@SW_SHOW)
EndFunc   ;==>fuWholeCommGUI


Func fuCongressPickerGUI($sThisBill)
	$hCongGUI = GUICreate("Select House Members", 350, 500, Default, Default, Default, Default, $hBillGUI)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the secondary GUI [X] is clicked
	Local $hSelectMembersHeadLabel = GUICtrlCreateLabel("Select Members:", 15, 7, 320, 20, $SS_LEFT)
	GUICtrlSetFont($hSelectMembersHeadLabel, Default, $FW_BOLD)
	Local $hSelectMembersTextLabel = GUICtrlCreateLabel("Select Members from the 'Select Member' combobox below.  As you click on a Member that Member will be added to the listbox below.  " & _
			"When you have finished selecting Members click on the Print Word Docs button to print Word Docs.", 10, 27, 325, 75, $SS_LEFT)
	Local $hSelectMemberComboHeaderLabel = GUICtrlCreateLabel("Select Member", 15, 95, 320, 16, $SS_CENTER)
	; Get a list of House members from MS Word doc and convert it into an array
	Local $oWord = _Word_Create(False, Default)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_Create House Members", "Error creating a new Word instance." & _
			@CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $sDocument = "\\alpha3\MARKUP\SenateHouseMembers\House.Doc"
	Local $oWordDoc = _Word_DocOpen($oWord, $sDocument, Default, Default, True)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocOpen Example 1", "Error opening '.\Extras\Test.doc'." & _
			@CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $TextDoc = $oWordDoc.Content.Text ; Ask to Receive the Text from Contents Object of the Object Document
	_Word_DocClose($oWordDoc)
	_Word_Quit($oWord)
	Local $aHouseMem = StringSplit($TextDoc, @CR)

	; And here we get the elements into a list
	Local $sMemList = ""
	For $i = 3 To UBound($aHouseMem) - 3
		$sMemList &= "|" & StringSplit($aHouseMem[$i], "   � ", $STR_ENTIRESPLIT)[1]
	Next
	; Create the combo
	$hCombo = GUICtrlCreateCombo("", 15, 120, 320, 20)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
	; And fill it
	GUICtrlSetData($hCombo, $sMemList)

	Local $hSelectMembersListLabel = GUICtrlCreateLabel("The below Members will print.  To delete a Member in the listbox below click on their name and then press the DELETE key.", 25, 150, 300, 60, $SS_LEFT)
	Local $hBillNumberLabel = GUICtrlCreateLabel("Bill Number: " & $sThisBill, 25, 190, 300, 16, $SS_CENTER)
	GUICtrlSetFont($hBillNumberLabel, Default, $FW_BOLD)
	Local $hListboxLabel = GUICtrlCreateLabel("LISTBOX", 25, 230, 300, 16, $SS_CENTER)
	HotKeySet("{DELETE}", "HotKeyPressed")

	$hListbox = GUICtrlCreateList("", 25, 250, 300, 200, BitOR($WS_BORDER, $WS_VSCROLL))
	$hPrintDocButton = GUICtrlCreateButton("Print Word Docs", 25, 450, 125, 40)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hCancelDocButton = GUICtrlCreateButton("Close / Cancel", 200, 450, 125, 40)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUISetState(@SW_SHOW)
EndFunc   ;==>fuCongressPickerGUI

;~ Captures certain keys activation
Func HotKeyPressed()
	Switch @HotKeyPressed ; The last hotkey pressed
		Case "{DELETE}" ; String is the {DELETE} hotkey
			Local $sName = GUICtrlRead($hListbox)
			_GUICtrlListBox_DeleteString($hListbox, _GUICtrlListBox_GetCaretIndex($hListbox))
	EndSwitch
EndFunc   ;==>HotKeyPressed

;~ Captures an On_Click event
Func On_Click()
	Switch @GUI_CtrlId ; See wich item sent a message
		Case $Date
			$ValidDate = _DateIsValid(GUICtrlRead($Date))
			If $ValidDate Then
				GUICtrlSetData($DateSelected, "Date Selected: " & GUICtrlRead($Date))
			EndIf
		Case $hProcHeadsButton
			_ObjDictDeleteKey($toGenLeave)
			_ObjDictDeleteKey($toWholeCommittee)
			GUICtrlSetState($hMrChairButton, $GUI_DISABLE)
			GUICtrlSetData($hWCbillsNum, "")
			GUICtrlSetBkColor($hWCbillsNum, $GUI_BKCOLOR_TRANSPARENT)
			fuProcHeads()
		Case $hGenLeaveButton
			Local $iPos
			Local $SelectedGenBills[0][2]
			For $iX = 0 To UBound($CheckBoxes) - 1
				If BitAND(GUICtrlRead($CheckBoxes[$iX][0]), $GUI_CHECKED) = $GUI_CHECKED Then
					$iPos = UBound($SelectedGenBills) + 1
					ReDim $SelectedGenBills[$iPos][1]
					$SelectedGenBills[$iPos - 1][0] = GUICtrlRead($CheckBoxes[$iX][0], 1)
;~ 					$SelectedGenBills[$iPos - 1][1] = GUICtrlRead($CheckBoxes[$iX][1])
				EndIf
			Next
			fuCreateGenLeaveExcelSheet($toGenLeave, $toWholeCommittee, $SelectedGenBills)
		Case $hMrChairButton
			GUISetState(@SW_DISABLE, $hMainGUI)
			fuWholeCommGUI()
		Case $hDefault_Button
			$cInputFolder = $cInputFolderDefault
			GUICtrlSetData($inputFolder, $cInputFolder)
			$cOutputFolder = $cOutputFolderDefault
			GUICtrlSetData($outputFolder, $cOutputFolder)
		Case $hApply_Button
			$cInputFolder = GUICtrlRead($inputFolder)
			$cInputFolder = StringRegExpReplace($cInputFolder, '\\* *$', '') ; strip trailing \ and spaces
			If Not FileExists($cInputFolder) Then
				MsgBox(16, "Input folder invalid", "Input folder does not exists. Enter a valid input folder.")
			Else
				If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\BillHeads", "input", "REG_SZ", $cInputFolder) Then
					MsgBox(16, "Input folder could not be saved", "The input folder could not be saved, Error #" & @error)
				EndIf
			EndIf
			GUICtrlSetData($inputFolder, $cInputFolder)

			$cOutputFolder = GUICtrlRead($outputFolder)
			$cOutputFolder = StringRegExpReplace($cOutputFolder, '\\* *$', '') ; strip trailing \ and spaces
			If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\BillHeads", "output", "REG_SZ", $cOutputFolder) Then
				MsgBox(16, "Output folder could not be saved", "The output folder could not be saved, Error #" & @error)
			EndIf
			GUICtrlSetData($outputFolder, $cOutputFolder)
		Case $hWholeCommButton
			For $x = 0 To UBound($Radios) - 1
				If BitAND(GUICtrlRead($Radios[$x]), $GUI_CHECKED) = $GUI_CHECKED Then
					$SelectedBill = GUICtrlRead($Radios[$x], 1)
				EndIf
			Next
			GUISetState(@SW_DISABLE, $hBillGUI)
			fuCongressPickerGUI($SelectedBill)
		Case $hCancelDocButton
			GUIDelete($hCongGUI) ; If it was this GUI - we just delete the GUI <<<<<<<<<<<<<<<
			GUISetState(@SW_ENABLE, $hBillGUI)
		Case $hPrintDocButton
			Local $hLB = GUICtrlGetHandle($hListbox)
			Local $iNameCount = _GUICtrlListBox_GetCount($hLB)
			Local $asMembers[0]
			For $i = 0 To $iNameCount - 1
				_ArrayAdd($asMembers, _GUICtrlListBox_GetText($hLB, $i))
			Next
			fuCreateComWholeCoverDoc($SelectedBill, $asMembers)
	EndSwitch
EndFunc   ;==>On_Click

Func fuProcHeads()
	Dim $aMonths[13] = ["00", "JA", "FE", "MR", "AP", "MY", "JN", "JY", "AU", "SE", "OC", "NO", "DE"]
	Dim $cDay = GUICtrlRead($Date)
	Dim $nMonth = Number(StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$2'))
	Dim $cTempDay = StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$3')
	Dim $cCaptureFileName = "*" & $cTempDay & $aMonths[$nMonth] & "7.*"
	$cInputFolder = StringRegExpReplace($cInputFolder, '\\$', '')
	$cInputFileName = $cInputFolder & "\" & $cCaptureFileName
	Local $aFileList = _FileListToArray($cInputFolder, $cCaptureFileName, $FLTA_FILES, True)

	If $aFileList <> 0 Then

		; concatenate all files for that day together into one string
		$sInputFileText = ''
		For $i = 1 To $aFileList[0]
			$sInputFileText &= FileRead($aFileList[$i])
		Next

		Local $aWholeCommitteeBuckets = StringRegExp($sInputFileText, '(?sm)I81\w(?:(?!I66F).)*?(?i:I89In the Committee of the Whole).*?I66F', $STR_REGEXPARRAYGLOBALMATCH)
		Local $aGeneralLeaveBuckets = StringRegExp($sInputFileText, '(?sm)I81\w(?:(?!I66F).)*?(?i:I89General Leave)(?:(?!I89In the Committee of the Whole).)*?I66F', $STR_REGEXPARRAYGLOBALMATCH)

		; Run the routines
		$toGenLeave = fuPopulateGeneralLeaveHash($aGeneralLeaveBuckets)


		$toWholeCommittee = fuPopulateWholeCommitteeHash($aWholeCommitteeBuckets)
		Local $iWholeCount = _ObjDictCount($toWholeCommittee)

		If $iWholeCount > 0 Then
			GUICtrlSetState($hMrChairButton, $GUI_ENABLE)
			GUICtrlSetData($hWCbillsNum, $iWholeCount & " bills")
			GUICtrlSetBkColor($hWCbillsNum, 0x00FF00)
		EndIf

		If _ObjDictCount($toGenLeave) = 0 And $iWholeCount = 0 Then
			MsgBox(48, "No Bills Found", 'There are no General Leave Bills for ' & $cDay)
		Else
			GUISetState(@SW_DISABLE, $hMainGUI)
			fuChooseBillHeadsGUI()
		EndIf

		GUICtrlSetData($progressbar, 100)
		Sleep(2000)
		GUICtrlSetData($progressbar, 0)

	Else
		MsgBox(16, "Bills do not exist", 'There are no Bills for ' & $cDay & '. Try selecting another date.')
	EndIf
EndFunc   ;==>fuProcHeads


Func On_Close()
	Switch @GUI_WinHandle ; See which GUI sent the CLOSE message
		Case $hMainGUI
			Exit ; If it was this GUI - we exit <<<<<<<<<<<<<<<
		Case $hBillGUI
			GUIDelete($hBillGUI) ; If it was this GUI - we just delete the GUI <<<<<<<<<<<<<<<
			GUIDelete($hCongGUI) ; Also delete a child GUI <<<<<<<<<<<<<<<
			GUISetState(@SW_ENABLE, $hMainGUI)
		Case $hCongGUI
			GUIDelete($hCongGUI) ; If it was this GUI - we just delete the GUI <<<<<<<<<<<<<<<
			GUISetState(@SW_ENABLE, $hBillGUI)
	EndSwitch
EndFunc   ;==>On_Close


; function to get input or output values from registry if they exist
Func fuGetInputOutput($IorO, $DefaultFolder)

	Dim $inputreg, $outputreg

	If $IorO = "input" Then
		$inputreg = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\BillHeads", "input")
		If $inputreg = "" Then
			RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\BillHeads", "input", "REG_SZ", $DefaultFolder)
			Return $DefaultFolder
		Else
			Return $inputreg
		EndIf
	Else
		$outputreg = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\BillHeads", "output")
		If $outputreg = "" Then
			RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\BillHeads", "output", "REG_SZ", $DefaultFolder)
			Return $DefaultFolder
		Else
			Return $outputreg
		EndIf
	EndIf

EndFunc   ;==>fuGetInputOutput

; function to populate General Leave object dictionary
Func fuPopulateGeneralLeaveHash($asBills)
	Dim $to_dict = _ObjDictCreate()
	Local $sHeader = Null
	Local $sBillNo = Null
	For $i = 0 To UBound($asBills) - 1
		$sHeader = StringRegExp($asBills[$i], '(?sm)(?<=I81)(\w.*?)(?=\s*\n)', $STR_REGEXPARRAYMATCH)
		$sHeader[0] = StringRegExpReplace($sHeader[0], '_', '—') ; Replace underscore with em dash
		$sHeader[0] = StringRegExpReplace($sHeader[0], 'ÿ09', '–') ; Replace ÿ09 with en dash
		$sHeader[0] = wrap_text($sHeader[0], 55) ; Text wrap at around 35 characters to fit nicely into Excell cell
		$sBillNo = StringRegExp($asBills[$i], '(?sm)\(([H|S]\..*?)\)', $STR_REGEXPARRAYMATCH)
		If @error == 0 Then
			_ObjDictAdd($to_dict, $sHeader[0], $sBillNo[0])
		Else
			$sBillNo = StringRegExp($asBills[$i], '(House.*?Resolution\s*?\d+|Senate.*?Resolution\s*?\d+)', $STR_REGEXPARRAYMATCH)
			If @error == 0 Then
				_ObjDictAdd($to_dict, $sHeader[0], $sBillNo[0])
			Else
				$sBillNo = StringRegExp($asBills[$i], 'Special Order', $STR_REGEXPARRAYMATCH)
				If @error == 0 Then
					_ObjDictAdd($to_dict, $sHeader[0], "s/o")
				Else
					_ObjDictAdd($to_dict, $sHeader[0], "XXX")
				EndIf
			EndIf
		EndIf
	Next
	Return $to_dict
EndFunc   ;==>fuPopulateGeneralLeaveHash

; function to populate "In the Committee of the Whole" object dictionary
Func fuPopulateWholeCommitteeHash($asWholeBills)
	Local $to_dict_whole = _ObjDictCreate()
	Local $sHeader = Null
	For $i = 0 To UBound($asWholeBills) - 1
		$sHeader = StringRegExp($asWholeBills[$i], '(?sm)(?<=I81)(\w.*?)(?=\s*\n).*?(bill\s*\(([H|S]\..*?)\).*)(?=,\swith\sM)', $STR_REGEXPARRAYMATCH)
		If UBound($sHeader) = 3 Then
			_ObjDictAdd($to_dict_whole, $sHeader[2], $sHeader)
		Else
			$sHeader = StringRegExp($asWholeBills[$i], '(?sm)(?<=I81)(\w.*?)(?=\s*\n).*?(concurrent\sresolution\s*\(([H|S]\..*?)\).*)(?=,\swith\sM)', $STR_REGEXPARRAYMATCH)
			If UBound($sHeader) = 3 Then
				_ObjDictAdd($to_dict_whole, $sHeader[2], $sHeader)
			Else
				MsgBox($MB_ICONWARNING, "Irregular Wording", "The Committee of the Whole House text does not conform to usual pattern''." _
						 & "You must process this entry manually.")
			EndIf
		EndIf

	Next
	Return $to_dict_whole
EndFunc   ;==>fuPopulateWholeCommitteeHash

; function to produce Excel sheet of General Leave
Func fuCreateGenLeaveExcelSheet($toActs, $toWhole, $asChosenBills = '')
	Local $asActs[1][3] = [["", "", ""]]
	Local $iRowIndex = 0
	Local $asBillData
	Local $i = 1
	For $myBillName In $toActs
		$this_var = _ObjDictGetValue($toActs, $myBillName)
		If @error = 5 Then
			$asBillData = _ObjDictGetValue($toWhole, $myBillName)
			$asBillData[0] = StringRegExpReplace($asBillData[0], '_', '—') ; Replace underscore with em dash
			$this_var = $asBillData[0]
			$asActs[$i][2] = $asBillData[1]
		EndIf

		$iRowIndex = _ArraySearch($asChosenBills, $this_var, 0, 0, 0, 0, 1, 0)
		If $iRowIndex >= 0 Then
			ReDim $asActs[UBound($asActs) + 1][3]
			$asActs[0][0] = $asActs[0][0] + 1
			$asActs[$i][0] = $myBillName
			$asActs[$i][1] = $this_var
			$i += 1
		EndIf
	Next

	For $myBillName In $toWhole
		$asBillData = _ObjDictGetValue($toWhole, $myBillName)
		$asBillData[0] = StringRegExpReplace($asBillData[0], '_', '—') ; Replace underscore with em dash
		$this_var = $asBillData[2]

		$iRowIndex = _ArraySearch($asChosenBills, $this_var, 0, 0, 0, 0, 1, 0)
		If $iRowIndex >= 0 Then
			ReDim $asActs[UBound($asActs) + 1][3]
			$asActs[0][0] = $asActs[0][0] + 1
			$asActs[$i][0] = wrap_text($asBillData[0], 55)
			$asActs[$i][1] = $this_var
			$asActs[$i][2] = wrap_text($asBillData[1], 55)
			$i += 1
		EndIf
	Next

	Local $oExcel1 = _Excel_Open()
	If @error Then Exit MsgBox($MB_ICONERROR, "Excel UDF: _Excel_Open General Leave", "Error creating a new Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $oWorkbook = _Excel_BookNew($oExcel1)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF:_Excel_BookNew General Leave", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel1)
		Exit
	EndIf

	$oExcel1.ActiveSheet.Columns("A:C").ColumnWidth = 60
	If @error Then MsgBox(64, "Excel Bill Sheet", "Error " & @error & " returned by function '_ExcelRowHeightSet' on line " & @ScriptLineNumber)
	$asActs[0][0] = fuCreateCorrectDate() & @CRLF

	_Excel_RangeWrite($oWorkbook, $oExcel1.ActiveSheet, $asActs, "A1", Default, True)
	If @error Then Exit MsgBox($MB_ICONERROR, "Excel UDF: _Excel_RangeWrite General Leave", "Error writing to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	$oExcel1.ActiveSheet.Range("A1").Font.Size = 24
	$oExcel1.ActiveSheet.Range("A1").Font.Bold = True
	$oExcel1.ActiveSheet.Range("A:C").VerticalAlignment = -4160
	$oExcel1.ActiveSheet.Columns("B:B").AutoFit
	Return
EndFunc   ;==>fuCreateGenLeaveExcelSheet

; function to print Committee of the Whole covers
Func fuCreateComWholeCoverDoc($sRecordKey, $asMemberNames)
	Local $asBillData = _ObjDictGetValue($toWholeCommittee, $sRecordKey)
	$asBillData[0] = StringRegExpReplace($asBillData[0], '_', '—') ; Replace underscore with em dash
	; Put bill abstract on a clipboard
	ClipPut($asBillData[1] & ":")
	$oWordApp = _Word_Create(0, True)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_Create Template Doc", "Error creating a new Word instance." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	Dim $progpercent = 10
	Dim $progincrement = Round(UBound($asMemberNames) / $progpercent)
	GUICtrlSetData($progressbar, 0)

	Local $oDoc = 0
	Local $asNamesState[0]
	Local $sNameString = "", $sStateString = ""
	For $sMemberName In $asMemberNames
		$asNamesState = StringSplit($sMemberName, ", ", $STR_ENTIRESPLIT)
		If $asNamesState[0] = 3 Then
			$sNameString = "HON. " & (StringStripWS($asNamesState[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)) _
					 & " " & (StringStripWS($asNamesState[1], $STR_STRIPLEADING + $STR_STRIPTRAILING))
			$sStateString = "of " & StringStripWS(StringRegExp($asNamesState[3], "(?s)[^\(]*", $STR_REGEXPARRAYMATCH)[0], $STR_STRIPLEADING + $STR_STRIPTRAILING)
		ElseIf $asNamesState[0] = 4 Then
			$sNameString = "HON. " & (StringStripWS($asNamesState[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)) _
					 & " " & (StringStripWS($asNamesState[1], $STR_STRIPLEADING + $STR_STRIPTRAILING)) & ", " _
					 & (StringStripWS($asNamesState[3], $STR_STRIPLEADING + $STR_STRIPTRAILING))
			$sStateString = "of " & StringStripWS(StringRegExp($asNamesState[4], "(?s)[^\(]*", $STR_REGEXPARRAYMATCH)[0], $STR_STRIPLEADING + $STR_STRIPTRAILING)
		EndIf

		$oDoc = _Word_DocAdd($oWordApp, $wdNewBlankDocument, @ScriptDir & "\COTWTemplate.doc")
		If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Template", "Error creating a new Word document from template." _
				 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<TITLE>", $asBillData[0])
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Bill Title", _
				"Error replacing text in the document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<aCtWoRdS>", "^c")
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Bill Summary", _
				"Error replacing text in the document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<DateOfBill>", fuCreateCorrectDate())
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Date", _
				"Error replacing text in the document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<StateOf>", $sStateString)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace State", _
				"Error replacing text in the document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<MemberName>", $sNameString)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Member Name", _
				"Error replacing text in the document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		GUICtrlSetData($progressbar, (100 - ($progincrement * UBound($asMemberNames))))
	Next

	$oWordApp.Visible = True

	Return
EndFunc   ;==>fuCreateComWholeCoverDoc

; Function to wrap long lines of text to insert into Excel sheet
Func wrap_text($txt, $col = 80)
	$txt = StringRegExpReplace($txt, '(.{1,' & $col & '})( +|$)\n?|(.{' & $col & '})', '$1$3' & @CRLF)
	Return $txt
EndFunc   ;==>wrap_text

;~ Function to create a date string without leading zeros
Func fuCreateCorrectDate()
	Dim $cDay = GUICtrlRead($Date)
	Local $aMyDate, $aMyTime
	_DateTimeSplit($cDay, $aMyDate, $aMyTime)
	Local $sLongDayName = _DateDayOfWeek(_DateToDayOfWeek($aMyDate[1], $aMyDate[2], $aMyDate[3]))
	Local $sLongMonthName = _DateToMonth($aMyDate[2])
	Local $sCompleteDateValue = $sLongDayName & ', ' & $sLongMonthName & ' ' & Number($aMyDate[3]) & ', ' & $aMyDate[1]
	Return $sCompleteDateValue
EndFunc   ;==>fuCreateCorrectDate

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)
	$nNotifyCode = BitShift($iwParam, 16)
	$nID = BitAND($iwParam, 0x0000FFFF)
	$hCtrl = $ilParam
	If $nID = $hCombo Then
		Switch $nNotifyCode
			Case $CBN_SELCHANGE
				Local $sSelectedMemberName = StringStripWS(GUICtrlRead($hCombo), $STR_STRIPLEADING + $STR_STRIPTRAILING)
				If $sSelectedMemberName <> '' Then _XferNames($sSelectedMemberName)
		EndSwitch
	EndIf
EndFunc   ;==>WM_COMMAND

Func _GetVersion()
	If @Compiled Then
		Return FileGetVersion(@AutoItExe)
	Else
		Return IniRead(@ScriptFullPath, "FileVersion", "#AutoIt3Wrapper_Res_Fileversion", "0.0.0.0")
	EndIf
EndFunc   ;==>_GetVersion

Func _XferNames($sSelectedMemberName)
	Local $hLB = GUICtrlGetHandle($hListbox)
	_GUICtrlListBox_BeginUpdate($hLB)
	_GUICtrlListBox_AddString($hLB, $sSelectedMemberName)
	_GUICtrlListBox_EndUpdate($hLB)
EndFunc   ;==>_XferNames
