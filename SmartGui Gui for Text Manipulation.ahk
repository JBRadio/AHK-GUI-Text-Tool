; AHK GUI for Text Manipulation
; Author: James E Burns
;
; Types of Text Input
;  - Line of Text (no CRLF)
;  - Lines of Text (separated by CRLF)
;  - 2 Sets of Lines of Text (comparing, combining, etc.) -- Pending review.
;  - Conversion field elements (Kilograms to US Pounds, etc.) -- Pending review.
;
; Special Data Formats
;  General
;   Plain text
;   Number Patterns (Telephone, Currency, 
;   Text Patterns (US Address, URL, Delimiter separated data, Email, ...
;
;  Development
;   WEB: JSON, XML, HTML, CSS
;   JavaScript
;   PHP
;
;
; Ideas for GUI/Usage/User Feedback Improvement
;  - Remove "Done." as the default status text for functions. Allow functions to return "Done." or something else.
;  - ButtonUndo: Save up to the last 10 actions and undo one at a time.
;  - ButtonRecent: Process the most recent command
;  - ButtonSave: Save output to a text file
;
; Ideas for Tools
;  Programming (General)
;   - Code Alignment/Formatting
;   - Count matching characters (in attempt to find a missing [], (), {}, ', ", <> and so on..)
;   - Generate Lorem Ipsum
;   - Generate GUID
;  Programming (Web: HTML, CSS, XML, JSON, JavaScript)
;   - Color Picker (by Name/Hex)
;   - Markdown to HTML
;   - CSV/*SV to HTML Table
;   - Lint (JS, CSS, HTML, etc.)
;   - Minify (JS, CSS, HTML, etc.)
;   - Preview HTML file in the browser
;  PowerShell/Command Prompt user
;  Text
;   Lists/Lines
;    - Combine (allow duplicates)
;    - Merge (preventing duplicates)
;    - Compare (showing differences)
;    - Join (on one line, with/out separator)
;    - Indent (decrease/increase preceeding tab)
;    - Tabs to Spaces (user chooses how many spaces per tab)
;   Lists,CSV / *SV
;    - Swap 1 CSV column position with another column
;
;#NoTrayIcon

/*
CHANGE LOG
----------
 2019.09.08:
			GUI/ / Application
			 * Changed font to fixed-width font (Consolas) for input/output text boxes
			 * Changed font to bold for Text labels (default font is system, or MS Shell Dlg for me)
			 * Changed Win+4 toggle logic to no longer use a state variable / check for active window title and class instead
			 * Removed minimized box to prevent invalid GUI placement on restore after being minimized when Windows Taskbar is vertical; ESC, Close button, and Win+4 hides the GUI
			 
 2019.09.06:
			GUI / Application
			 * Win+4 now toggles the GUI show/hide. Escape and GuiClose only hides when the GUI is on and then toggles state.
			Function RemoveBlankLines
			 * Changed the condition to break out of loop when removing all double CRLF and double LF
             
 2019.08.30: ;Gui, Margin, 8, -10 - Commented out line so ListView/Tab would have perfect margins
			Function: HelperExtractColumnDataForAnySV
			 * Should work with any delimiter entered now
			Function: HelperAskUserForDelim() - NEW
			* When *SV function was used, `t or \t didn't resolve to a tab character, fixed now. - ExtractColumnDataForAnySV, SwapColumnForAnySV, FilterColumnDataForAnySV, SortDataByColumnForAnySV, SpacedColumnForAnySV		
			
 2019.08.25: Added # directive (#NoEnv), SendMode, SetWorkingDir at the top
             Moved MVC files to 'MVC' folder
			 Moved 3rd party files to '3rd Party' folder
*/

#SingleInstance Force

; #Warn						; Enable warnings to assist with detecting common errors.
#NoEnv						; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  			; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%	; Ensures a consistent starting directory.

; 3rd Party Dependencies
#Include %A_ScriptDir%\3rd Party\Text-to-Speech.ahk
#Include %A_ScriptDir%\3rd Party\JSON_FromObj.ahk
#Include %A_ScriptDir%\3rd Party\JSON_Beautify.ahk
#Include %A_ScriptDir%\3rd Party\FGP - FileGetProperties.ahk

; My Library Dependencies
#Include MyDSVTool.ahk		; Delimiter-Separated-Values Tool (CSV, TSV, etc.)
global myDSVTool := New MyDSVTool

; Main Program
; Gui, New	; Breaks Gui hide/show ; This is a reminder
;Gui, Margin, 8, -10
Gui, Add, Tab3, y+5 vTabVar, General|Programmer

Gui, Tab, General,, Exact
#Include %A_ScriptDir%\MVC\Tab General VIEW.ahk
;Gui, Add, Button, Hidden Default, OK

Gui, Tab, Programmer,, Exact
#Include %A_ScriptDir%\MVC\Tab Programmer VIEW.ahk
;Gui, Add, Button, Hidden Default, OKProg

Gui, Tab  ; i.e. subsequently-added controls will not belong to the tab control.
Gui, Add, Button, Hidden Default, OK	; Allows for enter key on ListView to submit choice

;Gui, Margin	; Reset Margin

; Shared controls
Gui, Font,,Consolas ; Effort to use fixed-width font for better visual effect (especially in Output)
Gui, Add, Edit, vEditInput x17 y205 w330 h70
Gui, Add, Edit, vEditOutput x17 y315 w330 h70 ReadOnly, Readonly processed results appear here. (Can be selected and copied)

Gui, Font, bold
Gui, Add, Text, vGuiTextUserInput x17 y185 w80 h20, User Input 	; Additional 10 for y
Gui, Add, Text, vGuiTextOutput x17 y295 w80 h20, Ouput		 	; Additional 10 for y

Gui, Font, norm ; remove bold
Gui, Font,,MS Shell Dlg
Gui, Add, Button, vGuiButtonClear x367 y205 w100 h30, Clear		; Additional 5 for y ; + 10 width
Gui, Add, Button, vGuiButtonPaste x367 y240 w100 h30, Paste
Gui, Add, Button, vGuiButtonProcess x367 y275 w100 h30, Process
Gui, Add, Button, vGuiButtonCopy x367 y315 w90 h30 , Copy
Gui, Add, Button, vGuiButtonMoveUp x367 y355 w90 h30 , Move Up

; Status Bar
Gui, Add, StatusBar,, Status text will appear here	; Usage of StatusBar: SB_SetText("There are " . RowCount . " rows selected.")

; Define Windows dimensions for GUI
hMaxSize := A_ScreenHeight
wMaxSize := A_ScreenWidth / 2

Gui, +Resize +MinSize +MaxSize%wMaxSize%x%hMaxSize% -MinimizeBox	; MinSize with no suffix to use the window's current size as the limit
;Gui, +MinSize +Resize 									; MinSize with no suffix to use the window's current size as the limit
Gui, Show, w480 h425,Text Manipulation	                ; Application window appears centered on load
;Gui, Show, Hide w480 h425,Text Manipulation

; Force focus to first ListView initially
;~ GuiControl, Focus, MyListView
;~ LV_Modify(1, "Select")

return

; When GUI size changes:
GuiSize:
	if ( A_GuiHeight == 425 and A_GuiWidth == 480 )
	{
		; Normal size
		GuiControl, Move, EditInput, x17 y205 w330 h70
		GuiControl, Move, EditOutput, x17 y315 w330 h70
		GuiControl, Move, GuiButtonClear, x367 y205 w100 h30 ; + 10 width
		GuiControl, Move, GuiButtonPaste, x367 y240 w100 h30
		GuiControl, Move, GuiButtonProcess, x367 y275 w100 h30
		GuiControl, Move, GuiButtonCopy, x367 y315 w100 h30
		GuiControl, Move, GuiButtonMoveUp, x367 y355 w100 h30
		GuiControl, Move, GuiTextUserInput, x17 y185 w80 h20
		GuiControl, Move, GuiTextOutput, x17 y295 w80 h20
		GuiControl, Move, MyListView, w440 ;x17 y30 w440 h130
		GuiControl, Move, MyListViewProg, w440 ;x17 y30 w440 h130
		GuiControl, Move, TabVar, w460 ; x17 y30 w440 h130
		;Gui, Show, AutoSize ; Redraw
	}
	else
	{
		; Input		x17 y205 w330 h70 (
		; Output	x17 y315 w330 h70
		
		hMaxEditInput := (A_ScreenHeight / 3) - 5
		;MsgBox % hMaxEditInput
		
		;hEditInput := (A_GuiHeight-355) / 1.5 < 150 ? (A_GuiHeight-355) / 1.5 : 150 ; Max Height = 580
		hEditInput := (A_GuiHeight-355) / 1.5 < hMaxEditInput ? (A_GuiHeight-355) / 1.5 : hMaxEditInput ; Max Height = 580
		hEditInput := hEditInput < 70 ? 70 : hEditInput
		
		yEditOutput := 245 + hEditInput ; EditInputy + EditInputh = 275 - 315 (difference between EditInput and EditOutput = 40)
		yTextOutput := 225 + hEditInput
		yButtonMoveUp := yEditOutput + 40 ; 355 - 315
		
		GuiControl, MoveDraw, EditInput, % "w" A_GuiWidth-150 "h" hEditInput ;A_GuiHeight-355
		GuiControl, MoveDraw, EditOutput, % "y" yEditOutput "w" A_GuiWidth-150 "h" hEditInput ; "y" A_GuiHeight-110  "h" A_GuiHeight-355
		GuiControl, MoveDraw, GuiButtonClear, % "x" A_GuiWidth-113 ; x367 y170 w90 h30
		GuiControl, MoveDraw, GuiButtonPaste, % "x" A_GuiWidth-113 ; x367 y205 w90 h30
		GuiControl, MoveDraw, GuiButtonProcess, % "x" A_GuiWidth-113 ; x367 y245 w90 h30
		GuiControl, MoveDraw, GuiButtonCopy, % "x" A_GuiWidth-113 "y" yEditOutput ; x367 y315 w90 h30
		GuiControl, MoveDraw, GuiButtonMoveUp, % "x" A_GuiWidth-113 "y" yButtonMoveUp ; x367 y355 w90 h30
		;GuiControl, MoveDraw, GuiTextUserInput, x17 y185 w80 h20
		GuiControl, MoveDraw, GuiTextOutput, % "y" yTextOutput ; x17 y295 w80 h20
		GuiControl, MoveDraw, MyListView, % "w" A_GuiWidth-43 ; x17 y30 w440 h130
		GuiControl, MoveDraw, MyListViewProg, % "w" A_GuiWidth-43 ; "w" A_GuiWidth-23 ; x17 y30 w440 h130
		GuiControl, MoveDraw, TabVar, % "w" A_GuiWidth-23 ; x17 y30 w440 h130
	}
return

#4::
	; If GUI is showing but not active - Activate it / Show it
	; If GUI is showing and active - Hide it
	; If GUI is not showing - Show it
	
	WinGetTitle, vTitle, A
	WinGetClass, vClass, A
	
	if ( (InStr(vTitle, "Text Manipulation") != 1) or (vClass != "AutoHotkeyGUI") )
	{
		; Our window name starts with "Text Manipulation" and the active window doesn't start off the same, show GUI
		; If the active window is not class AutoHotkeyGUI, we know we aren't the active window
		Gui, Show ; Activate - whether it's under another window or toggled off
		
	} else {
		Gui, Hide
	}
return

; When GUI closes
GuiClose:
    Gui, Hide
return

#IfWinActive ahk_class AutoHotkeyGUI ; Do not block Escape key usage for non-AHK GUI Applications

; When user hits escape while AutoHotkey GUI is active
ESC::
    Gui, Hide
return

; Workaround to execute code when pressing enter in ListView, when there are multiple ListView and multiple Tabs
; https://autohotkey.com/board/topic/55167-using-enter-with-multiple-listviews-in-one-script/
#IfWinActive, Text Manipulation
Enter::
	GuiControlGet, OutputVarFocusedVariable, FocusV
	if ( OutputVarFocusedVariable == "MyListView" )
	{
		Gui, ListView, MyListView
		EnterKeyPressedOnListView()
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg" )
	{
		Gui, ListView, MyListViewProg
		EnterKeyPressedOnListView()
	}
	else
	{
		SendInput, {Enter}	; Allow default behavior
	}
return

EnterKeyPressedOnListView()
{
	selRow := LV_GetNext(0, "Focused")
	LV_GetText(OutputVarFunctionName, selRow, 4)
	LV_GetText(OutputVarInput, selRow, 2)
	ProcessFunction(OutputVarFunctionName, OutputVarInput)
}

; Workaround to scroll when multiple ListViews and/or tabs with listviews are present
; https://www.autohotkey.com/boards/viewtopic.php?t=678
WheelDown::
	GuiControlGet, OutputVarFocusedVariable, FocusV
	GuiControlGet, Current_Tab,,TabVar
	;~ if ( OutputVarFocusedVariable == "MyListView" or Current_Tab == "General" )
	;~else if ( OutputVarFocusedVariable == "MyListViewProg" or Current_Tab == "Programmer")
	; Prevents scrolling in textbox
	if ( OutputVarFocusedVariable == "MyListView")
	{
		Gui, ListView, MyListView
		GuiControl,, MyListView
		sendmessage, 0x115, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg")
	{
		Gui, ListView, MyListViewProg
		GuiControl,, MyListViewProg
		sendmessage, 0x115, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
	}
	else
	{
		SendInput, {WheelDown}	; Allow default behavior
	}
return

; Workaround to scroll when multiple ListViews and/or tabs with listviews are present
; https://www.autohotkey.com/boards/viewtopic.php?t=678
WheelUp::
	GuiControlGet, OutputVarFocusedVariable, FocusV
	;GuiControlGet, Current_Tab,,TabVar
	;if ( OutputVarFocusedVariable == "MyListView" or Current_Tab == "General" )
	;else if ( OutputVarFocusedVariable == "MyListViewProg" or Current_Tab == "Programmer" )
	; Prevents scrolling in textboxes
	
	if ( OutputVarFocusedVariable == "MyListView" )
	{
		;~ Gui, ListView, MyListView
		;~ GuiControl,, MyListView
		sendmessage, 0x115, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg" )
	{
		sendmessage, 0x115, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
	}
	else
	{
		SendInput, {WheelUp}	; Allow default behavior
	}
return

; Shift + Wheel for horizontal scrolling for ListView
; https://gist.github.com/cheeaun/160999
+WheelDown::
	; WheelRight
	; Scroll slowly for maximum effect
	GuiControlGet, OutputVarFocusedVariable, FocusV
	;GuiControlGet, Current_Tab,,TabVar
	;if ( OutputVarFocusedVariable == "MyListView" or Current_Tab == "General" )
	;else if ( OutputVarFocusedVariable == "MyListViewProg" or Current_Tab == "Programmer" )
	if ( OutputVarFocusedVariable == "MyListView" )
	{
		Gui, ListView, MyListView
		GuiControl,, MyListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 1, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg" )
	{
		Gui, ListView, MyListViewProg
		GuiControl,, MyListViewProg
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 1, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
	}
	else
	{
		;SendInput, {WheelDown}	; Allow default behavior
	}
return

+WheelUp::
	;WheelLeft
	; Scroll slowly for maximum effect
	GuiControlGet, OutputVarFocusedVariable, FocusV
	GuiControlGet, Current_Tab,,TabVar
	if ( OutputVarFocusedVariable == "MyListView"  or Current_Tab == "General" )
	{
		Gui, ListView, MyListView
		GuiControl,, MyListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		sendmessage, 0x114, 0, 0,, ahk_id %MyListView% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg" or Current_Tab == "Programmer" )
	{
		Gui, ListView, MyListViewProg
		GuiControl,, MyListViewProg
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
		sendmessage, 0x114, 0, 0,, ahk_id %MyListViewProg% ; Works with hwndhlv on Gui, Add, ListView
	}
	else
	{
		;SendInput, {WheelDown}	; Allow default behavior
	}
return

#IfWinActive ; Restore hotkeys to work from any environment

#Include %A_ScriptDir%\MVC\Tab General CONTROLLER.ahk		; General Tab's Controller - button logic
#Include %A_ScriptDir%\MVC\Tab General MODEL.ahk 			; Includes all the functions called upon by the General Tab's Functions

#Include %A_ScriptDir%\MVC\Tab Programmer CONTROLLER.ahk	; Programmer Tab's Controller - button logic, listview initiator
#Include %A_ScriptDir%\MVC\Tab Programmer MODEL.ahk		; Includes all the functions called upon by the Programmer Tab's Functions (unless shared with an existing function)