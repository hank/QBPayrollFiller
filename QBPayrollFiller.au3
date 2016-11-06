#include <MsgBoxConstants.au3>
#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <GuiButton.au3>
#include <WindowsConstants.au3>
#include "Tesseract.au3"

Global $g_idBtn, $g_hGUI, $g_done

; Configuration
; Tesseract top adjustment
Global $g_tess_top_adj = 5
Global $g_tess_left_adj = 40
; Tesseract multiplier/divisor
Global $g_tess_mul_div = 3

Example()

Func Example()
	; === TESSERACT ===
	Local $tesseract_exe = IniRead(@ScriptDir & "\Settings.ini", "General", "Tesseract", "")
	_TesseractTempPathSet(@TempDir)
	_TesseractExePathSet(FileGetShortName($tesseract_exe))
	Local $xlsx = IniRead(@ScriptDir & "\Settings.ini", "General", "PayrollSpreadsheet", "")
	Local $fedWithPos[2], $ssERPos[2]

	Local $hQBPaycheck = WinGetHandle("Preview Paycheck","&Do not accrue sick/")

	If 0 = $hQBPaycheck Then
		MsgBox(0, "Failed to find window", "Failed to find QB Preview Paycheck window, open it in QB")
		Exit
	EndIf

	ConsoleWrite("QB Window: " & $hQBPaycheck & @LF)

	Local $out = _TesseractWinCapture($hQBPaycheck, "", 0, "T", 0, $g_tess_mul_div, 0, 0, 0, 0, 0)
	if IsArray($out) Then
;~ 		level	page_num	block_num	par_num	line_num	word_num	left	top	width	height	conf	text
;~ 		_ArrayDisplay($out, "Tesseract Text Capture")
		for $i = 0 to (UBound($out)-1)
			; Find Federal Withholding
			If $out[$i][12] = "Federal" And $out[$i+1][12] = "Withholding" Then
				$fedWithPos[0] = _AdjustTessCoordX($out[$i][7])
				$fedWithPos[1] = _AdjustTessCoordY($out[$i][8])
			EndIf
		Next
		for $i = 0 to (UBound($out)-1)
			; Find Federal Withholding
			If $out[$i][12] = "Social" And $out[$i+1][12] = "Security" And $out[$i+2][12] = "Company" Then
				$ssERPos[0] = _AdjustTessCoordX($out[$i][7])
				$ssERPos[1] = _AdjustTessCoordY($out[$i][8])
			EndIf
		Next
	EndIf

	; === EXCEL ===
	Local $oExcel = _Excel_Open()
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeCopy Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	; Open Workbook
	Local $oWorkbook = _Excel_BookOpen($oExcel, $xlsx, True)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeCopy Example", "Error opening workbook '" & @ScriptDir & "\Extras\_Excel1.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel)
		Exit
	EndIf

	Opt("GUIOnEventMode", 1) ; Change to OnEvent mode
	Local $aWorkSheets = _Excel_SheetList($oWorkbook)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_SheetList Example 1", "Error listing Worksheets." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	; _ArrayDisplay($aWorkSheets, "Excel UDF: _Excel_SheetList Example 1")

	; Create a GUI
	$g_hGUI = GUICreate("Select Workbook", 280, 100)
	GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEButton")

	GUICtrlCreateLabel("Select a workbook", 30, 10)
	; Create the combo
	$hCombo = GUICtrlCreateCombo("", 30, 30, 200, 20)
	; And fill it
	$sList = ""
	For $i = 0 To UBound($aWorkSheets) - 1
		$sList &= "|" & $aWorkSheets[$i][0]
	Next
	GuiCtrlSetData($hCombo, $sList)
	Local $iOKButton = GUICtrlCreateButton("OK", 70, 60, 60)
	GUICtrlSetOnEvent($iOKButton, "OKButton")

	GUISetState(@SW_SHOW, $g_hGUI)

	Local $sheetSelection
	While 1
		If $g_done == 1 Then
			; Get the selected workbook
			$sheetSelection = GuiCtrlRead($hCombo)
			ConsoleWrite("Selected sheet: " & $sheetSelection & @LF)
			GuiDelete($g_hGUI)
			ExitLoop
		EndIf
		Sleep(10)
	WEnd

	; Activate selected sheet
	$oWorkbook.Sheets($sheetSelection).Activate

	; Grab the Federal Withholding from the spreadsheet
	Local $fedWithVal = $oWorkbook.ActiveSheet.Range("C6").text
	Local $ssEEVal = $oWorkbook.ActiveSheet.Range("C7").text
	Local $medEEVal = $oWorkbook.ActiveSheet.Range("C8").text
	Local $mdWithVal = $oWorkbook.ActiveSheet.Range("C9").text
	Local $401kEEVal = $oWorkbook.ActiveSheet.Range("C10").text
	Local $ssERVal = $oWorkbook.ActiveSheet.Range("D7").text
	Local $medERVal = $oWorkbook.ActiveSheet.Range("D8").text
	Local $401kERVal = $oWorkbook.ActiveSheet.Range("D14").text

	; ConsoleWrite("Fed With: " & $fedWithVal & @LF)

	; Click Federal Withholding
	WinActivate($hQBPaycheck)
	Local $winpos = WinGetPos("Preview Paycheck")
	ConsoleWrite("fedWithPos: " & $fedWithPos[0] & " " & $fedWithPos[1] & @LF)
	MouseClick("right",$fedWithPos[0] + $winpos[0], $fedWithPos[1] + $winpos[1], 1)
	; Enter Federal Withholding and move down
	Send($fedWithVal & "{TAB}")
	; Enter Social Security and move down
	Send($ssEEVal & "{TAB}")
	; Enter Medicare and move down
	Send($medEEVal & "{TAB}")
	; Enter MD Withholding
	Send($mdWithVal)

	ConsoleWrite("ssERPos: " & $ssERPos[0] & " " & $ssERPos[1] & @LF)
	MouseClick("right",$ssERPos[0] + $winpos[0], $ssERPos[1] + $winpos[1], 1)
	; Enter SS ER and move down
	Send($ssERVal & "{TAB}")
	; Enter Medicare ER and move down
	Send($medERVal & "{TAB}")

	; Close excel
	_Excel_Close($oExcel)

EndFunc   ;==>Example

Func OKButton()
    ; Note: At this point @GUI_CtrlId would equal $iOKButton,
    ; and @GUI_WinHandle would equal $hMainGUI
    ConsoleWrite("You selected OK!" & @LF)
	$g_done = 1
EndFunc   ;==>OKButton

Func CLOSEButton()
    ; Note: At this point @GUI_CtrlId would equal $GUI_EVENT_CLOSE,
    ; and @GUI_WinHandle would equal $hMainGUI
    ConsoleWrite("You selected CLOSE! Exiting..." & @LF)
    Exit
EndFunc   ;==>CLOSEButton

Func _AdjustTessCoordX($in)
	return Floor($in/$g_tess_mul_div) + $g_tess_left_adj
EndFunc

Func _AdjustTessCoordY($in)
	return Floor($in/$g_tess_mul_div) + $g_tess_top_adj
EndFunc