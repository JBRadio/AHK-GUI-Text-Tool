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
 2021.07.03:
			SMARTGUI TEXT MANIPULATION.AHK
			Added the app main menu bar (File, Edit, Input, Output, Options, Help)
			 * Add "File" Menu - Open Folder, Reload, Close, ExitApp
			 * Add "Input" Menu - Clear, Copy, Paste, Process
			 * Add "Output" Menu - Copy, Move Up
			 * Add "Help" Menu - About This Tool
 2021.06.27:
			TAB GENERAL MODEL.AHK
			Function: HelperExtractColumnDataForAnySV
			 * Changed how the sample row, to choose which column to extract, is presented. Now lists each column in a new line with it's number choice.
 2021.03.07:
			TAB GENERAL MODEL.AHK
			Function: TextPrefixSuffixRemoveEachLine
			 * Removed Regex and moved it to its own function. This function now does a simple replace. Prefix removal may not work will with beginning indentation.
			
			Function: TextPrefixSuffixRemoveEachLineRegex *NEW*
			 * Same code as previous version of TextPrefixSuffixRemoveEachLine
			
			TAB GENERAL CONTROLLER.AHK
			ListViewInit:
			 * Made distinct options ( "Prefix/Suffix Remove (Text)"and "Prefix/Suffix Remove (Regex)" )
			 
 2020.11.07:
            TAB GENERAL MODEL.AHK
			Function: HelperGetSVDataIntoRowsArray
			 * Now trims each cell value ([row][colum])
 
 2020.08.14:
			TAB GENERAL MODEL.AHK
			Function: HelperTextJoinLinesWithChar
			 * Updated the join to take place after the first item. Previously, the joining/separating character was being applied before the list instead of between list items.
 
 2020.06.21:
			TAB GENERAL MODEL.AHK
			Function: TextCharLineUp
			 * Removed quotations on prompt box
 2020.05.09:
			TAB GENERAL MODEL.AHK
			Function: TextJoinBrokenLines - Join Broken Lines w/ w/o Char", "1+ Lines", "Joins every x lines as user defined to a single line by user entered separator"
 
 2020.01.26:
			Main file
			 * Removed Script Hotkey Ctrl+Alt+R - Created "My AHK Script Reloader" application to handle the reloading and manage AHK scripts

 2020.01.25:
            Main file
			Made all #Include paths relative to %A_ScriptDir%\, so that the script can be reloaded by another script, if necessary
			
 2020.01.11:
			Main file
			Hotkey to Reload this script (CTRL+ALT+R) - Remove in EXE version
			Hotkey to Edit this script (CTRL+ALT+E) - Remove in EXE version
			Hotkey to Open folder of Source code (CTRL+ALT+F) - Remove in EXE version
			
			TAB GENERAL MODEL.AHK
			Function: TextMakeWindowsFileFriendly - "Replace (Windows)" *** Function shared to Programmer tab/controller
			 * Results returned was incorrect due to incorrect SubStr() arguments.
			
 2019.12.01:
			Main file
			 * Minor screen change, "Ouput" to "Output" above text box
			Function: sendMessageMouseWheel() - created to reduce redundant code
			
			AHK
			 * Updated to version 1.1.32.0

 2019.11.22:
			TAB GENERAL MODEL.AHK
			Function: SpacedColumnVerticalForAnySV - *SV (Spaced Columns V), also renamed existing function to *SV (Spaced Columns H)
			 * Prints row data vertically, rather than horizontally
			 * Dependencies: HelperSpacedColumnVerticallyForAnySV, HelperFindRowMaxWidth

 2019.11.15:
			TAB GENERAL MODEL.AHK
			Function: RemoveWhitespaceEachLine *NEW*
			 * Removes duplicate whitespace by removing preceeding/trailing whitespace (trim) first and then duplicate whitespace between text

 2019.11.09:
			Main Program
			DebuggerMode (Ctrl+Alt+Shit+d)
			 * Allow for simple debugger info by MsgBox (Wheelup/down when ListView doesn't scroll with wheel)
			 
 2019.11.01-02:
			TAB GENERAL MODEL.AHK
			Function RemoveWhitespace
			 * Reduces whitespace - whenever there is 2+ in a row, it is reduced to one
			
 2019.10.25-27:
			TAB GENERAL MODEL.AHK
			Function: TextReplaceStringReplace "Replace (AHK StringReplace)"
			 * Use AHK's StringReplace to find and replace all matching text patterns
			Function: TextReplaceStringReplaceEachLine "Replace (AHK StringReplace Each Line)"
			 * Use AHK's StringReplace function to find and replace text patterns line by line
			Function: TextPrefixOrdinalNumberRenumberIndented "List lines with ordinal numbers, renumbered for indented line items"
			 * 1 -> 2 -> 3 .... 1 -> 1.1 -> 1.2 -> 2 -> 2.1 -> 2.2 -> 3

 2019.10.19-20:
			TAB GENERAL MODEL.AHK
			Function: TextPrefixNestedList
			 * Take a List (lines of text) and converts it into a level/nested-oriented list
			 
 2019.09.20:
            TAB GENERAL MODEL.AHK
             * Removing commented out superfluous status bar message functions - replaced by SB_SetText()
             * Removed unnecessary comments and commented out code
             * Reviewing changes to reduce redundant code
 
 2019.09.15:
			GENERAL TAB
			Function: HelperFilterColumnDataForAnySV
			 * Created new reusable functions and trimmed down size of the function
				HelperDynamicallyFindHeightForInputBox(vPromptText) - Function to determine height for AHK InputBox based on number of new lines (`n) found in the Prompt message
				HelperBuildPickAColumn(vRowFirst) - Returns "Pick a number between 1 and {Max Columns}"
				HelperBuildVerticalSampleRow(vRowFirst) - Returns the first row of data, numbered and separated by new lines (`n)
			Function: HelperAskUserForDelim
			 * Now uses HelperDynamicallyFindHeightForInputBox for its prompt message
				Prompt := "What character separates the data?"
				Height := HelperDynamicallyFindHeightForInputBox(Prompt) ; (vPromptText)

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
			Function GUtilityResultsLhsToFe
			 * Added Line: RemoveBlankLines() ; Remove any blank lines - which can occur from a MG Workstation ScreenGrab on GUtility Results - Save the user a step - Function from General Tab
			
 2019.08.30: ;Gui, Margin, 8, -10 - Commented out line so ListView/Tab would have perfect margins
			Function: HelperExtractColumnDataForAnySV
			 * Should work with any delimiter entered now
			Function: HelperAskUserForDelim() - NEW
			* When *SV function was used, `t or \t didn't resolve to a tab character, fixed now. - ExtractColumnDataForAnySV, SwapColumnForAnySV, FilterColumnDataForAnySV, SortDataByColumnForAnySV, SpacedColumnForAnySV
			Function: SirtTableStrip() - NEW - SIRT TSV (Clean Print)			
			
 2019.08.25: Added # directive (#NoEnv), SendMode, SetWorkingDir at the top
             Moved MVC files to 'MVC' folder
			 Moved 3rd party files to '3rd Party' folder
*/

#SingleInstance Force

; #Warn						; Enable warnings to assist with detecting common errors.
#NoEnv						; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  			; Recommended for new scripts due to its superior speed and reliability.
;SetWorkingDir %A_ScriptDir%	; Ensures a consistent starting directory.

; 3rd Party Dependencies
#Include %A_ScriptDir%\3rd Party\Text-to-Speech.ahk
#Include %A_ScriptDir%\3rd Party\JSON_FromObj.ahk
#Include %A_ScriptDir%\3rd Party\JSON_Beautify.ahk
#Include %A_ScriptDir%\3rd Party\FGP - FileGetProperties.ahk

; My Library Dependencies
#Include %A_ScriptDir%\MyDSVTool.ahk		; Delimiter-Separated-Values Tool (CSV, TSV, etc.)

global myDSVTool := New MyDSVTool
global appVersion := "ver. 1.0"

; Main Program

; Add "File" Menu - Open Folder, Reload, Close, ExitApp
Menu, AppMenuFileSubMenu, Add, Open Folder, FileMenuHandler
Menu, AppMenuFileSubMenu, Add, Reload, FileMenuHandler
Menu, AppMenuFileSubMenu, Add, Close, FileMenuHandler
Menu, AppMenuFileSubMenu, Add, Exit App, FileMenuHandler

; Add "Input" Menu - Clear, Copy, Paste, Process
Menu, AppMenuInputSubMenu, Add, Clear, InputMenuHandler
Menu, AppMenuInputSubMenu, Add, Copy, InputMenuHandler
Menu, AppMenuInputSubmenu, Add, Paste, InputMenuHandler
Menu, AppMenuInputSubmenu, Add, Process, InputMenuHandler

; Add "Output" Menu - Copy, Move Up
Menu, AppMenuOutputSubmenu, Add, Copy, OutputMenuHandler
Menu, AppMenuOutputSubmenu, Add, Move Up, OutputMenuHandler

; Add "Help" Menu - About This Tool
Menu, AppMenuHelpSubmenu, Add, About Text Manipulation Tool, HelpMenuHandler

; Add main menu - File, Edit, Input, Output, Options, About
Menu, AppMenu, Add, File, :AppMenuFileSubMenu
Menu, AppMenu, Add, Edit, AppMenu
Menu, AppMenu, Add, Input, :AppMenuInputSubmenu
Menu, AppMenu, Add, Output, :AppMenuOutputSubmenu
Menu, AppMenu, Add, Options, AppMenu
Menu, AppMenu, Add, Help, :AppMenuHelpSubmenu

Gui, Menu, AppMenu

; Gui, New	; Breaks Gui hide/show ; This is a reminder
;Gui, Margin, 8, -10
Gui, Add, Tab3, y+5 vTabVar, General|Programmer
Gui, Tab, General,, Exact
#Include %A_ScriptDir%\MVC\Tab General VIEW.ahk
;Gui, Add, Button, Hidden Default, OK ; Uncomment if General is the only tab

Gui, Tab, Programmer,, Exact
#Include %A_ScriptDir%\MVC\Tab Programmer VIEW.ahk
;Gui, Add, Button, Hidden Default, OKProg ; Uncomment if Programmer is the only tab

Gui, Tab  ; i.e. subsequently-added controls will not belong to the tab control.
Gui, Add, Button, Hidden Default, OK	; Allows for enter key on ListView to submit choice

;Gui, Margin	; Reset Margin

; Shared controls
Gui, Font,,Consolas ; Effort to use fixed-width font for better visual effect (especially in Output)
Gui, Add, Edit, vEditInput x17 y205 w330 h70
Gui, Add, Edit, vEditOutput x17 y315 w330 h70 ReadOnly, Readonly processed results appear here. (Can be selected and copied)

Gui, Font, bold
Gui, Add, Text, vGuiTextUserInput x17 y185 w80 h20, User Input 	; Additional 10 for y
Gui, Add, Text, vGuiTextOutput x17 y295 w80 h20, Output		 	; Additional 10 for y

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
;Gui, +MinSize +Resize 									            ; MinSize with no suffix to use the window's current size as the limit
Gui, Show, w480 h425,Text Manipulation		                        ; Application window appears centered on load
;Gui, Show, Hide w480 h425, Text Manipulation

; Force focus to first ListView initially
;~ GuiControl, Focus, MyListView
;~ LV_Modify(1, "Select")

return

AppMenu:
return

FileMenuHandler:
	;MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	
	switch A_ThisMenuItem
	{
		case "Open Folder":
			vRunCommand := "explorer.exe /select,""" . A_ScriptFullPath . """"
			Run, %vRunCommand%
		return
		
		case "Reload":
			Reload
		return
		
		case "Close":
			Gui, Hide
		return
		
		case "Exit App":
			ExitApp
		return
	}
return

InputMenuHandler:	
	switch A_ThisMenuItem
	{
		case "Clear":
			GuiControl,, EditInput,												; Sets input text box to nothing (clears it)
			SB_SetText("Cleared all input text box.")
		return
			
		case "Copy":
			GuiControlGet, OutputVarOutput,,EditInput,Edit						; Gets output text box content
			OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")	; The "Edit" controls only have `n so need to add `r`n on "Copy"
			Clipboard := OutputVarStrReplace									; Sets Windows Clipboard to updated output text box content
			SB_SetText("Copied input to Windows Clipboard.")
		return
			
		case "Paste":
			if IsLabel("ButtonPaste")
				gosub, ButtonPaste
		return
			
		case "Process":
			if IsLabel("ButtonProcess")
				gosub, ButtonProcess
		return
			
		default:
		return
	}
return

OutputMenuHandler:	
	switch A_ThisMenuItem
	{
		case "Copy":
			if IsLabel("ButtonCopy")
				gosub, ButtonCopy
			return
			
		case "Move Up":
			if IsLabel("ButtonMoveUp")
				gosub, ButtonMoveUp
			return
			
		default:
			return
	}
return

HelpMenuHandler:
	switch A_ThisMenuItem
	{
		case "About Text Manipulation":
			MsgBox % "Text Manipulation Tool written by James Burns`r`nVersion: " appVersion
			return
			
		default:
			return
	}
return


; When GUI size changes:
GuiSize:
	;MsgBox % "GuiSize fired."
	; Width: 480
	; Height: 425
	
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
		;MsgBox % "hMaxEditInput: " hMaxEditInput
		
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
	;
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

; Workaround to execute code when pressing enter in ListView, when there are multiple ListView and multiple Tabs
; https://autohotkey.com/board/topic/55167-using-enter-with-multiple-listviews-in-one-script/
#IfWinActive, Text Manipulation

/*
^!+d::
	DebuggerMode := !DebuggerMode
	if (DebuggerMode)
		MsgBox % "DebuggerMode On!"
	else
		MsgBox % "DebuggerMode Off!"
return
*/

; CTRL+ALT+E to Edit this script
^!e::
	MsgBox, 4,, Program: Text Manipulation`nWould you like to edit the (main file) script?
	IfMsgBox, Yes, Edit
return

; CTRL+ALT+F to open folder in File Explorer
^!f::
	MsgBox, 4,, Program: Text Manipulation`nOpen source code folder in File Explorer?
	IfMsgBox, Yes
		Run, %A_ScriptDir%
return

; Removed - Created "My AHK Script Reloader" to handle the reloading
; CTRL+ALT+R to Reload this script
;^!r::
;	Reload
;	Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
;	MsgBox, 4,, Program: Text Manipulation`nThe script could not be reloaded. Would you like to open it for editing?
;	IfMsgBox, Yes, Edit
;return

; When user hits escape while AutoHotkey GUI is active
ESC::
	Gui, Hide
return

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
	
	;if (DebuggerMode)
	;	MsgBox % "DEBUG WheelDown FocusV: " OutputVarFocusedVariable 
	
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
	
	;if (DebuggerMode)
	;	MsgBox % "DEBUG WheelUp FocusV: " OutputVarFocusedVariable 
	
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
	
	;if (DebuggerMode)
	;	MsgBox % "DEBUG +WheelDown FocusV: " OutputVarFocusedVariable 
	
	if ( OutputVarFocusedVariable == "MyListView" )
	{
		Gui, ListView, MyListView
		GuiControl,, MyListView
		sendMessageMouseWheel(MyListView, "Down", 7) ; (vID, vDirection, vTimes := 0)
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg" )
	{
		Gui, ListView, MyListViewProg
		GuiControl,, MyListViewProg
		sendMessageMouseWheel(MyListViewProg, "Down", 7) ; (vID, vDirection, vTimes := 0)
	}
	else
	{
		;SendInput, {WheelDown}	; Allow default behavior
	}
return

+WheelUp::
	; Horizontal scroll, WheelLeft
	; 
	GuiControlGet, OutputVarFocusedVariable, FocusV
	;GuiControlGet, Current_Tab,,TabVar ; Does this need to be here?
	
	;if (DebuggerMode)
	;	MsgBox % "DEBUG +WheelUp FocusV: " OutputVarFocusedVariable 
	
	if ( OutputVarFocusedVariable == "MyListView"  or Current_Tab == "General" )
	{
		Gui, ListView, MyListView
		GuiControl,, MyListView
		sendMessageMouseWheel(MyListView, "Up", 7) ; (vID, vDirection, vTimes := 0)
	}
	else if ( OutputVarFocusedVariable == "MyListViewProg" or Current_Tab == "Programmer" )
	{
		Gui, ListView, MyListViewProg
		GuiControl,, MyListViewProg
		sendMessageMouseWheel(MyListViewProg, "Up", 7) ; (vID, vDirection, vTimes := 0)
	}
	else
	{
		;SendInput, {WheelDown}	; Allow default behavior
	}
return

sendMessageMouseWheel(vID, vDirection, vTimes := 0){
	
		Loop % vTimes
		{
			if ( vDirection == "Up" )
				sendmessage, 0x114, 0, 0,, ahk_id %vID% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
			else if ( vDirection == "Down" )
				sendmessage, 0x114, 1, 0,, ahk_id %vID% ; Works with hwndhlv on Gui, Add, ListView ; 0x114 is WM_HSCROLL
		}
}

#IfWinActive ; Restore hotkeys to work from any environment

; DEBUG CODE TO SEE WHY LISTVIEW LOSES ABILITY TO BE SCROLLED
WheelDown::
	GuiControlGet, OutputVarFocusedVariable, FocusV
	GuiControlGet, Current_Tab,,TabVar
	
	;if (DebuggerMode)
	;	MsgBox % "DEBUG WheelDown FocusV: " OutputVarFocusedVariable 
	
	SendInput, {WheelDown}
return

; DEBUG CODE TO SEE WHY LISTVIEW LOSES ABILITY TO BE SCROLLED
WheelUp::
	GuiControlGet, OutputVarFocusedVariable, FocusV
	;GuiControlGet, Current_Tab,,TabVar
	
	;if (DebuggerMode)
	;	MsgBox % "DEBUG WheelUp FocusV: " OutputVarFocusedVariable 
	
	SendInput, {WheelUp} ; Do not prevent default behavior
return


#Include %A_ScriptDir%\MVC\Tab General CONTROLLER.ahk		; General Tab's Controller - button logic
#Include %A_ScriptDir%\MVC\Tab General MODEL.ahk 			; Includes all the functions called upon by the General Tab's Functions

#Include %A_ScriptDir%\MVC\Tab Programmer CONTROLLER.ahk	; Programmer Tab's Controller - button logic, listview initiator
#Include %A_ScriptDir%\MVC\Tab Programmer MODEL.ahk			; Includes all the functions called upon by the Programmer Tab's Functions (unless shared with an existing function)