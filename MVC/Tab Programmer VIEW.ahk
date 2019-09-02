Gui, Add, ListView, gMyListViewProg vMyListViewProg x17 y30 w440 h130 hwndMyListViewProg -Multi Sort, Function|Input|Description|Function Name
; Changing function to the first column so that the user can use keyboard keys to search the listing while the ListView has focus.
;	"In addition to navigating from row to row with the keyboard, the user may also perform incremental search by typing the first few characters of an item in the first column.
;	This causes the selection to jump to the nearest matching row."
; Global variable "g" to work with click events
; Variable "v" to work with FocusV
; Making it the first element may allow it to start with focus

; Moving elements to be within the tabs - y + 20, x + 10
;~ Gui, Add, Edit, vEditInputProg x17 y205 w330 h70
;~ Gui, Add, Edit, vEditOutputProg x17 y315 w330 h70 ReadOnly, Readonly processed results appear here. (Can be selected and copied)
;~ Gui, Add, Button, gButtonClearProg x367 y165 w90 h30 , Clear
;~ Gui, Add, Button, gButtonPasteProg x367 y205 w90 h30 , Paste
;~ Gui, Add, Button, gButtonProcessProg x367 y245 w90 h30 , Process
;~ Gui, Add, Button, gButtonCopyProg x367 y315 w90 h30 , Copy
;~ Gui, Add, Button, gButtonMoveUpProg x367 y355 w90 h30 , Move Up
;~ Gui, Add, Text, x17 y185 w80 h20, User Input  ; Additional 10 for y
;~ Gui, Add, Text, x17 y295 w80 h20, Ouput		  ; Additional 10 for y
;Gui, Add, Text, x17 y395 w32 h20, Status:
;Gui, Add, Text, vStatusMessage x53 y395 w407 h20

;Gui, Add, Button, gButtonOKProg Hidden Default, OK
;	"To detect when the user has pressed Enter while a ListView has focus, use a default button (which can be hidden if desired)."

; Generated using SmartGUI Creator for SciTE

gosub, ListViewInitProg		; Initialize ListView values
LV_ModifyCol()  ; There are no parameters in this mode. ; Auto-size all columns to fit their contents
;~ LV_ModifyCol(1, 130)		; Modify width of column 1
;~ LV_ModifyCol(2, 55)			; Modify width of column 2
;~ LV_ModifyCol(3, 180)		; Modify width of column 3
;LV_Modify(0, "Select")		; Make first option selected, to ensure an option is selected (Will need to check for a selection when process is clicked.)
;LV_ModifyCol(1, "Sort")

GuiControl, Focus, MyListViewProg
LV_Modify(1, "Select")