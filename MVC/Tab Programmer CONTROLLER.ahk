;~ ProcessFunctionProg(OutputVarFunctionName, OutputVarInputType){ ; Main processing function for MEDITECH tab
	;~ GuiControlGet, OutputVarInput,,EditInputProg,Edit
	
	;~ ;MsgBox % "Input Type: " OutputVarInputType
	
	;~ if ( StrLen(OutputVarInput) = 0  and (OutputVarInputType != "N/A") )
	;~ {
		;~ SB_SetText("User input field is blank. (ProcessFunctionProg)")
		;~ return
	;~ }

	;~ ReturnFunctionValue := %OutputVarFunctionName%(OutputVarInput)
	;~ GuiControl,, EditOutputProg, %ReturnFunctionValue%
	;~ ;SetStatusMessageAndColor("Done.", "Green") ; Allow functions to send back their own messages, whether positive or negative.
;~ }

;~ ButtonOKProg:
	;~ Gui, ListView, MyListViewProg
	;~ ;	"To detect when the user has pressed Enter while a ListView has focus, use a default button (which can be hidden if desired)."
	;~ GuiControlGet, OutputVarFocusedVariable, FocusV	; make sure the "v" flag is set on control element. "g" caused this not to work.
	;~ ;MsgBox % "OutputVarFocusedVariable: " OutputVarFocusedVariable
	;~ if (OutputVarFocusedVariable != "MyListViewProg")
		;~ return

	;~ ;;MsgBox % "Enter was pressed. The focused row number is " . LV_GetNext(0, "Focused")
	
	;~ ; Function|Input|Description|Function Name
	;~ selRow := LV_GetNext(0, "Focused")
	;~ LV_GetText(OutputVarFunctionName, selRow, 4)
	;~ LV_GetText(OutputVarInput, selRow, 2)
	;~ ProcessFunctionProg(OutputVarFunctionName, OutputVarInput)
;~ return

;~ ButtonClearProg:
	;~ GuiControl,, EditInputProg,
	;~ GuiControl,, EditOutputProg,
	;~ ;SetStatusMessageAndColor("Cleared all inputs.", "Green")
	;~ SB_SetText("Cleared all inputs.")
;~ return

;~ ButtonCopyProg:
	;~ GuiControlGet, OutputVarOutput,,EditOutputProg,Edit
	;~ OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")	; The "Edit" controls only have `n so need to add `r`n on "Copy"
	;~ Clipboard := OutputVarStrReplace
	;~ ;SetStatusMessageAndColor("Copied output to Windows Clipboard.", "Green")
	;~ SB_SetText("Copied output to Windows Clipboard.")
;~ return

;~ ButtonMoveUpProg:
	;~ GuiControlGet, OutputVarOutput,,EditOutputProg,Edit
	;~ OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")
	;~ GuiControl,, EditInputProg, %OutputVarStrReplace%
	;~ ;SetStatusMessageAndColor("Copied output to input.", "Green")
	;~ SB_SetText("Copied output to input.")
;~ return

;~ ButtonPasteProg:
	;~ GuiControl,, EditInputProg, %Clipboard%
	;~ ;SetStatusMessageAndColor("Pasted Windows Clipboard content to User Input.", "Green")
	;~ SB_SetText("Pasted Windows Clipboard content to User Input.")
;~ return

;~ ButtonProcessProg:
	;~ Gui, ListView, MyListViewProg
	;~ RowNumber := LV_GetNext(0,"F")
	;~ if (!RowNumber)
	;~ {
		;~ ;SetStatusMessageAndColor("No selected row found.", "Red")
		;~ SB_SetText("No selected row found.")
		;~ return
	;~ }
	
	;~ ; Function|Input|Description|Function Name
	;~ LV_GetText(OutputVarInput, RowNumber, 2)
	;~ LV_GetText(OutputVarFunctionName, RowNumber, 4)
	;~ ProcessFunctionProg(OutputVarFunctionName, OutputVarInput)
;~ return

MyListViewProg:
	GuiControlGet, OutputVarFocusedVariable, FocusV	; make sure the "v" flag is set on control element. "g" caused this not to work.
	
	if ( OutputVarFocusedVariable != "MyListViewProg")
	{
		SB_SetText("Wrong element has focus (MyListViewProg)")
		return
	}
	
	Gui, ListView, MyListViewProg
	
	; DEBUG
	;MsgBox % "A_GuiEvent: " A_GuiEvent
	
	if (A_GuiEvent = "DoubleClick")
	{
		LV_GetText(OutputVarFunction, A_EventInfo, 1)
		LV_GetText(OutputVarInput, A_EventInfo, 2)
		LV_GetText(OutputVarDescription, A_EventInfo, 3)
		LV_GetText(OutputVarFunctionName, A_EventInfo, 4)
		
		; DEBUG
		;MsgBox % "OutputVarFunction: " OutputVarFunction "`nOutputVarInput: " OutputVarInput "`nOutputVarDescription: " OutputVarDescription "`nOutputVarFunctionName: " OutputVarFunctionName
		;MsgBox % "MyListViewProg - Function name: " OutputVarFunctionName
		
		ProcessFunction(OutputVarFunctionName, OutputVarInput) ; Use Generic function
	}
return

ListViewInitProg:
	Gui, ListView, MyListViewProg
	; Type, Function, Description
	LV_Add("", "Browser Tab (Simple HTML Notepad)", "N/A", "Opens browser tab to editable textarea for text testing", "BrowserOpenTabToHtmlNotepad")
	LV_Add("", "Browser Tab (Eval HTML Text)", "1+ Lines", "Opens browser tab to flattened HTML string", "BrowserOpenTabToHtml")
	;
	LV_Add("", "Cmd.exe (Get IPConfig /all)", "N/A", "Runs IPConfig /all and copies it back", "CommandPromptGetIpconfigAll")
	LV_Add("", "Cmd.exe (Get NSLookup)", "1 Line", "Runs NSLookup for server/URL and copies it back", "CommandPromptGetNsLookup")
	LV_Add("", "Cmd.exe (Ping)", "1 Line", "Opens Cmd.exe and runs ping command for server name/URL/IP Address", "CommandPromptPing")
	LV_Add("", "Cmd.exe (SystemInfo)", "N/A", "Opens Cmd.exe and runs systeminfo command", "CommandPromptSystemInfo")
	LV_Add("", "Cmd.exe (Tracert)", "1 Line", "Opens Cmd.exe and runs tracert command for input", "CommandPromptTracert")
	;
	LV_Add("", "Convert LF to CRLF", "1+ Lines", "Converts Linefeeds to Carriage Return Linefeeds", "ConvertCarriageReturnLineFeed")	; From General
	;
	LV_Add("", "CSV (To <table>)", "CSV", "Convert CSV to html table code", "CsvToHtmlTable")
	;
    LV_Add("", "File (List Properties)", "N/A", "Returns a list of properties for the file chosen in the file dialog box.", "FileGetProperties")
	;
	;LV_Add("", "Show List (Windows Extended Properties)", "N/A", "Launches a child window to display list of information.", "ShowListOfWindowsExtendedProperties")
	;
	LV_Add("", "Generate Text (Lorem Ipsum)", "N/A", "Returns sample text (Lorem Ipsum) to use for testing.", "TextGenerateLoremIpsum")	; From General
	; Generate Text (GUID)
	;
	LV_Add("", "Char (Ascii)", "Char", "Encodes the leftmost character to ASCII value", "EncodeToAscii")
	LV_Add("", "Date MM/DD/YYYY (Get Full Date)", "1 Line", "Converts MM/DD/YYYY to full date", "TextGetDateFromMMDDYYYY")				; From General
	LV_Add("", "JSON (Format)", "1 Line", "Formats JSON notation to pretty print", "TextFormatJson")
	LV_Add("", "Number (Char Code)", "Char", "Encodes a number to Char Code value", "EncodeToCharCode")
	LV_Add("", "HTML (Encode)", "1+ Lines", "Encodes text to HTML entities (& to &amp;)", "EncodeToHtml")
	LV_Add("", "HTML (<table> to CSV)", "1+ Lines", "Simple conversion from HTML <table> to CSV", "HtmlTableToCSV")
	LV_Add("", "URL (Encode)", "1+ Lines", "Encodes text to URI/URL values", "EncodeToUrl")
	LV_Add("", "URL (Decode)", "1+ Lines", "Decodes URI text", "DecodeUrl")
	LV_Add("", "URL (Decode, Get Params)", "1 Line", "Decodes URI text and lists URI params", "DecodeUrlAndGetParams")
	LV_Add("", "URL (Download to Text File)", "1 Line", "Downloads internet file to a text file and opens it.", "UrlDownloadToTextFileAndOpen")
	;
	LV_Add("", "Open Link (AHK RegEx Quick Reference)", "N/A", "Opens AHK RegEx Quick Reference page in default browser", "OpenLinkAHKRegexQuickReference")	; From General
	LV_Add("", "Open Link (RegExr.com - RegEx Tester)", "N/A", "Opens AHK RegEx Quick Reference page in default browser", "OpenLinkRegExr")					; From General
	;
	;
	LV_Add("", "HTML (Add Tag)", "1+ Lines", "Surround text with user defined HTML tag", "TextSurroundingHTMLTag")
	LV_Add("", "HTML (Add Tag Each Line)", "1+ Lines", "Surround text on each line with defined HTML tag", "TextSurroundingHTMLTagEachLine")
	;
	LV_Add("", "Replace (Windows)", "1+ Lines", "Replaces non-Windows filesystem friendly chars with underscores", "TextMakeWindowsFileFriendly")	; From General
	; Markdown (to HTML), "1+ Lines", "Convert Markdown (MD) to HTML" -- see Markdown_2_HTML.ahk by JasonDavis for code and GUI application (From Others folder)
	;
	LV_Add("", "Windows (Run)", "1 Line", "Use the Windows OS Run command on input.", "WindowsRun")
	;LV_Add("", "Function", "1+ Lines", "Desc", "FunctionName")
return