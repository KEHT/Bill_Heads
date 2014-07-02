;~ 07/01/2014 - sjohnson@gpo.gov - Alpha version (0.90) BILL_HEADS to process speeches

#include <file.au3>
#include <ClipBoard.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <DateTimeConstants.au3>
#include <ProgressConstants.au3>
#include <objDictonary.au3>
#include <word.au3>

Dim $yorno = 7
Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $cInputFileName, $cInputFile, $cInputFileText

;~ Global $cInputFolderDefault = "U:\Constitutional Heads\L Files"
Global $cInputFolderDefault = "E:\CR\OC"
;~ Global $cOutputFolderDefault = "U:\Constitutional Heads\Output"
Global $cOutputFolderDefault = "E:\RECSCAN\TofA"
Global $cInputFolder, $cOutputFolder

Global $tipmsg = "PLEASE WAIT..."

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate, $Button_1, $progressbar, $inputFolder, $outputFolder

; create GUI and tabs

GUICreate("Constitutional Heads Program v0.9", 350, 300)
$tab = GUICtrlCreateTab(5, 5, 340, 290)

; tab 0
$tab0 = GUICtrlCreateTabItem("Main")
GUICtrlCreateLabel("Choose Date Below:", 15, 40, 300)
$LocalDate = _Date_Time_GetLocalTime()
$Date = GUICtrlCreateMonthCal(_Date_Time_SystemTimeToDateStr($LocalDate, 1), 65, 70, 220, 140, $MCS_NOTODAY)
$DateSelected = GUICtrlCreateLabel("Date Selected: " & _Date_Time_SystemTimeToDateStr($LocalDate, 1), 15, 220, 300)
$Button_1 = GUICtrlCreateButton("Process Heads", 115, 240, 120)
$progressbar = GUICtrlCreateProgress(70, 275, 210, 10, $PBS_SMOOTH)

; tab 1
$tab1 = GUICtrlCreateTabItem("Settings")
GUICtrlCreateLabel("Input Folder:", 15, 40, 300)
$inputFolder = GUICtrlCreateInput("", 15, 65, 320, 20)
$cInputFolder = GetInputOutput("input", $cInputFolderDefault)
GUICtrlSetData($inputFolder, $cInputFolder)

GUICtrlCreateLabel("Output Folder:", 15, 100, 300)
$outputFolder = GUICtrlCreateInput("", 15, 125, 320, 20)
$cOutputFolder = GetInputOutput("output", $cOutputFolderDefault)
GUICtrlSetData($outputFolder, $cOutputFolder)

$Default_Button = GUICtrlCreateButton("Default", 15, 160, 75)
$Apply_Button = GUICtrlCreateButton("Apply", 100, 160, 75)


GUICtrlCreateTabItem("") ; end tabitem definition

GUISetState()

; Run the GUI until the dialog is closed