;~ 07/01/2014 - sjohnson@gpo.gov - Alpha version (0.90) BILL_HEADS to process speeches
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

Dim $yorno = 7
Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $cInputFileName, $cInputFile, $sInputFileText

;~ Global $cInputFolderDefault = "U:\Constitutional Heads\L Files"
Global $cInputFolderDefault = "E:\CR\FM"
;~ Global $cOutputFolderDefault = "U:\Constitutional Heads\Output"
Global $cOutputFolderDefault = "E:\RECSCAN\TofA"
Global $cInputFolder, $cOutputFolder

Global $tipmsg = "PLEASE WAIT..."

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate, $ProcHeadsButton, $MrChairButton, $hWCbillsNum, $progressbar, $inputFolder, $outputFolder
Dim $toWholeCommittee
; create GUI and tabs

GUICreate("Bill Heads Program v0.9", 350, 300)
$tab = GUICtrlCreateTab(5, 5, 340, 290)

; tab 0
$tab0 = GUICtrlCreateTabItem("Main")
GUICtrlCreateLabel("Choose Date Below:", 15, 40, 300)
$LocalDate = _DateAdd('d', -1, _NowCalcDate())
$Date = GUICtrlCreateMonthCal($LocalDate, 65, 70, 220, 140, $MCS_NOTODAY)
$DateSelected = GUICtrlCreateLabel("Date Selected: " & $LocalDate, 15, 220, 300)
$ProcHeadsButton = GUICtrlCreateButton("Process Heads", 35, 240, 120)
$hWCbillsNum = GUICtrlCreateLabel("", 230, 220, 50, 30, $SS_NOTIFY)
$MrChairButton = GUICtrlCreateButton("Print Mr. Chair DOC", 195, 240, 120)
GUICtrlSetState($MrChairButton, $GUI_DISABLE)
$progressbar = GUICtrlCreateProgress(35, 275, 280, 10, $PBS_SMOOTH)

; tab 1
$tab1 = GUICtrlCreateTabItem("Settings")
GUICtrlCreateLabel("Input Folder:", 15, 40, 300)
$inputFolder = GUICtrlCreateInput("", 15, 65, 320, 20)
$cInputFolder = fuGetInputOutput("input", $cInputFolderDefault)
GUICtrlSetData($inputFolder, $cInputFolder)

GUICtrlCreateLabel("Output Folder:", 15, 100, 300)
$outputFolder = GUICtrlCreateInput("", 15, 125, 320, 20)
$cOutputFolder = fuGetInputOutput("output", $cOutputFolderDefault)
GUICtrlSetData($outputFolder, $cOutputFolder)

$Default_Button = GUICtrlCreateButton("Default", 15, 160, 75)
$Apply_Button = GUICtrlCreateButton("Apply", 100, 160, 75)


GUICtrlCreateTabItem("") ; end tabitem definition

GUISetState()

; Run the GUI until the dialog is closed



While 1
	$msg = GUIGetMsg()
	Switch $msg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Default_Button
			$cInputFolder = $cInputFolderDefault
			GUICtrlSetData($inputFolder, $cInputFolder)
			$cOutputFolder = $cOutputFolderDefault
			GUICtrlSetData($outputFolder, $cOutputFolder)
		Case $Apply_Button
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
		Case $ProcHeadsButton
			Dim $aMonths[13] = ["00", "JA", "FE", "MR", "AP", "MY", "JN", "JY", "AU", "SE", "OC", "NO", "DE"]
			Dim $cDay = GUICtrlRead($Date)
			Dim $nMonth = Number(StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$2'))
			Dim $cTempDay = StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$3')
			Dim $cCaptureFileName = "*" & $cTempDay & $aMonths[$nMonth] & "7.*"
			$cInputFolder = StringRegExpReplace($cInputFolder, '\\$', '')
			$cInputFileName = $cInputFolder & "\" & $cCaptureFileName
			Local $aFileList = _FileListToArray($cInputFolder, $cCaptureFileName, $FLTA_FILESFOLDERS, True)

			If $aFileList <> 0 Then

				; concatenate all files for that day together into one string
				For $i = 0 To $aFileList[0]
					$sInputFileText &= FileRead($aFileList[$i])
				Next

				; preprocessing

;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '\r?\n', @CRLF) ; make end of line consistent
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '~', ChrW(0x07)) ; precedence code
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '\x{1A}', '') ; remove any stray end of files
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '(?<!\x{0A})\x{07}(I|S|F)', @CRLF & ChrW(0x07) & '$1') ; make sure each bell I, bell S and bell F is on its own line
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '\x{AE}MD[0-9A-Z]{2,2}\x{AF}', '') ; strip Xywrite modes
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '\x{AE}IP.*?\x{AF}', '') ; strip Xywrite modes
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, '\x{AE}PT.*?\x{AF}', '') ; strip Xywrite modes
;~ 				$sInputFileText = StringRegExpReplace($sInputFileText, "(\r\n)(\1)+", "\1") ; strip double newline

;~ 				   ConsoleWrite($sInputFileText)
				; split the variable into an array of lines

;~ 				Dim $aRecords = StringSplit($sInputFileText, @CRLF, $STR_ENTIRESPLIT)
;~ 				   _ArrayDisplay($aRecords)


				; temp location, change this folder in final script

;~ 				If Not FileExists($cOutputFolder) Then
;~ 					DirCreate($cOutputFolder)
;~ 				EndIf
;~ 				ConsoleWrite($sInputFileText)
				Local $aI81buckets = StringRegExp($sInputFileText, '(?sm)I81\w(?:(?!I66F).)*?I89General Leave.*?I66F', $STR_REGEXPARRAYGLOBALMATCH)

				Local $aWholeCommitteeBuckets = StringRegExp($sInputFileText, '(?sm)I81\w(?:(?!I66F).)*?(?i:I89In the Committee of the Whole).*?I66F', $STR_REGEXPARRAYGLOBALMATCH)
				Local $aGeneralLeaveBuckets = StringRegExp($sInputFileText, '(?sm)I81\w(?:(?!I66F).)*?(?i:I89General Leave)(?:(?!I89In the Committee of the Whole).)*?I66F', $STR_REGEXPARRAYGLOBALMATCH)
;~ 				_ArrayDisplay($aGeneralLeaveBuckets)
				; Run the routines
				Local $toGenLeave = fuPopulateGeneralLeaveHash($aGeneralLeaveBuckets)
;~ 				For $myBillName In $toGenLeave
;~ 					$this_var = _ObjDictGetValue($toGenLeave, $myBillName)
;~ 					ConsoleWrite("Bill Name: " & $myBillName & "|||||| Bill No: " & $this_var & @CRLF)
;~ 				Next
				fuCreateGenLeaveExcelSheet($toGenLeave)

				$toWholeCommittee = fuPopulateWholeCommitteeHash($aWholeCommitteeBuckets)


				If _ObjDictCount($toWholeCommittee) > 0 Then
					GUICtrlSetState($MrChairButton, $GUI_ENABLE)
					GUICtrlSetData($hWCbillsNum, _ObjDictCount($toWholeCommittee) & " bills")
					GUICtrlSetBkColor($hWCbillsNum, 0x00FF00)
				EndIf


				GUICtrlSetData($progressbar, 100)
				Sleep(2000)
				GUICtrlSetData($progressbar, 0)

			Else
				MsgBox(16, "Bills do not exist", 'There are no Bills for ' & $LocalDate & '. Try selecting another date.')
			EndIf
		Case $MrChairButton
			Switch MsgBox($MB_YESNOCANCEL + $MB_ICONQUESTION, "Print Options", "Do you want to print name and state?")
				Case $IDYES

				Case $IDNO
					fuCreateComWholeCoverDoc($toWholeCommittee)
			EndSwitch

		Case $GUI_EVENT_PRIMARYUP
			$ValidDate = _DateIsValid(GUICtrlRead($Date))
			If $ValidDate Then
				GUICtrlSetData($DateSelected, "Date Selected: " & GUICtrlRead($Date))
			EndIf
	EndSwitch
WEnd





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
		$sBillNo = StringRegExp($asBills[$i], '(?sm)\(([H|S]\..*?)\)', $STR_REGEXPARRAYMATCH)
		If @error == 0 Then
			_ObjDictAdd($to_dict, $sHeader[0], $sBillNo[0] )
		Else
			$sBillNo = StringRegExp($asBills[$i], '(House.*?Resolution\s*?\d+|Senate.*?Resolution\s*?\d+)', $STR_REGEXPARRAYMATCH)
			If @error == 0 Then
				_ObjDictAdd($to_dict, $sHeader[0], $sBillNo[0] )
			Else
				_ObjDictAdd($to_dict, $sHeader[0], Null)
			EndIf
		EndIf
	Next
;~ 	_ObjDictList($to_dict)
	Return $to_dict
EndFunc   ;==>fuPopulateGeneralLeaveHash

; function to populate "In the Committee of the Whole" object dictionary
Func fuPopulateWholeCommitteeHash($asWholeBills)
	Local $to_dict = _ObjDictCreate();[UBound($asBills)]
	Local $sHeader = Null
	For $i = 0 To UBound($asWholeBills) - 1
		$sHeader = StringRegExp($asWholeBills[$i], '(?sm)(?<=I81)(\w.*?)(?=\s*\n).*?(\w+\s\w+\s*\(([H|S]\..*?)\).*?other\spurposes)', $STR_REGEXPARRAYMATCH)
		_ObjDictAdd($to_dict, $sHeader[0], $sHeader)
	Next
	Return $to_dict
EndFunc   ;==>fuPopulateWholeCommitteeHash

; function to produce Excel sheet of General Leave
Func fuCreateGenLeaveExcelSheet($toActs)
	Local $asActs[1][2] = [["", ""]]
	$i = 1
	For $myBillName In $toActs
		ReDim $asActs[UBound($asActs) + 1][2]
		$this_var = _ObjDictGetValue($toActs, $myBillName)
		$asActs[0][0] = $asActs[0][0] + 1
		$asActs[$i][0] = $myBillName
		$asActs[$i][1] = $this_var
		$i += 1
	Next
;~ 	_ArrayDisplay($asActs)

	Local $oExcel1 = _Excel_Open()
	If @error Then Exit MsgBox($MB_ICONERROR, "Excel UDF: _Excel_Open General Leave", "Error creating a new Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $oWorkbook = _Excel_BookNew($oExcel1)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF:_Excel_BookNew General Leave", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel1)
		Exit
	EndIf
;~ 	Local $oSheet = _Excel_SheetAdd($oWorkbook)
;~ 	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_SheetAdd General Leave", "Error adding sheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_ArrayDelete($asActs, 0)
	_Excel_RangeWrite($oWorkbook, $oExcel1.ActiveSheet, $asActs, "A1")
	If @error Then Exit MsgBox($MB_ICONERROR, "Excel UDF: _Excel_RangeWrite General Leave", "Error writing to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	$oExcel1.ActiveSheet.Columns("A:B").AutoFit
	Return
EndFunc   ;==>fuCreateGenLeaveExcelSheet

; function to print no-name Committee of the Whole covers
Func fuCreateComWholeCoverDoc($recordHash)
	$oWordApp = _Word_Create(0, True)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_Create Template Doc", "Error creating a new Word instance." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	$oFinalWordApp = _Word_Create(0, True)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_Create Final Doc", "Error creating a new Word instance" _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	$oDoc_2 = _Word_DocAdd($oFinalWordApp)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Final Doc", "Error creating a new Word document." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	Dim $progpercent = 10
	Dim $progincrement = Round(_ObjDictCount($recordHash) / $progpercent)
	GUICtrlSetData($progressbar, 0)

	Local $oDoc = 0
	Local $asBill = Null
	For $myHR In $recordHash
		_ArrayDisplay(_ObjDictGetValue($recordHash, $myHR))
;~ 		ConsoleWrite("My key: "&$myHR&" My value: "&_ObjDictGetValue($recordHash, $myHR)&@CRLF)
		$asBill = _ObjDictGetValue($recordHash, $myHR)
		$oDoc = _Word_DocAdd($oWordApp, $wdNewBlankDocument, @ScriptDir & "\COTWTemplate.doc")
		If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Template", "Error creating a new Word document from template." _
				 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<TITLE>", $asBill[0])
		_Word_DocFindReplace($oDoc, "<act>", $asBill[1])

		$oRange = _Word_DocRangeSet($oDoc, -1, Default, Default, $wdStory, 1)
		$oRange.Copy

		$oFinalRange = _Word_DocRangeSet($oDoc_2, 0, $wdStory, 1, Default, Default)
		$oFinalRange.PasteAndFormat(16)
		_ObjDictDeleteKey($recordHash, $myHR)
		If _ObjDictCount($recordHash) <> 0 Then
			$oFinalRange = _Word_DocRangeSet($oDoc_2, $oFinalRange, $wdStory, 1, -1, Default)
			$oFinalRange.InsertBreak()
		EndIf
		GUICtrlSetData($progressbar, (100 - ($progincrement * _ObjDictCount($recordHash))))
	Next
	_Word_DocClose($oDoc)

	$oFinalWordApp.Visible = True
;~ 	_Word_DocSaveAs($oDoc_2, @ScriptDir & "\_Word_Test2.doc")
;~ 	_Word_DocClose($oDoc_2)
	Return
EndFunc
