#NoTrayIcon

#Include Text-to-Speech.ahk
#Include JSON_Beautify.ahk
;#Include jsonFormatter.ahk ; https://autohotkey.com/board/topic/94687-jsonformatter-json-pretty-print-using-javascript/


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
Gui, Add, Edit, vEditInput x7 y185 w330 h70
Gui, Add, Edit, vEditOutput x7 y295 w330 h70 ReadOnly, Readonly processed results appear here
Gui, Add, Button, x357 y145 w90 h30 , Clear
Gui, Add, Button, x357 y185 w90 h30 , Paste
Gui, Add, Button, x357 y225 w90 h30 , Process
Gui, Add, Button, x357 y295 w90 h30 , Copy
Gui, Add, Button, x357 y335 w90 h30 , Move Up
Gui, Add, Text, x7 y155 w80 h20, User Input
Gui, Add, Text, x7 y265 w80 h20, Ouput
Gui, Add, Text, x7 y375 w32 h20, Status:
Gui, Add, Text, vStatusMessage x43 y375 w407 h20

Gui, Add, Button, Hidden Default, OK
;	"To detect when the user has pressed Enter while a ListView has focus, use a default button (which can be hidden if desired)."

Gui, Add, ListView, gMyListView vMyListView x12 y10 w440 h130 -Multi Sort, Function|Input|Description|Function Name
; Changing function to the first column so that the user can use keyboard keys to search the listing while the ListView has focus.
;	"In addition to navigating from row to row with the keyboard, the user may also perform incremental search by typing the first few characters of an item in the first column.
;	This causes the selection to jump to the nearest matching row."
; Global variable "g" to work with click events
; Variable "v" to work with FocusV

; Generated using SmartGUI Creator for SciTE
;Gui, Show, w471 h396, Text Manipulation ; Use Win+3 to show
Gui, Show, Hide w471 h396, Text Manipulation ; Use Win+3 to show

gosub, ListViewInit			; Initialize ListView values
LV_ModifyCol(1, 130)		; Modify width of column 1
LV_ModifyCol(2, 55)			; Modify width of column 2
LV_ModifyCol(3, 180)		; Modify width of column 3
;LV_Modify(0, "Select")		; Make first option selected, to ensure an option is selected (Will need to check for a selection when process is clicked.)
;LV_ModifyCol(1, "Sort")

GuiControl, Focus, MyListView
LV_Modify(1, "Select")
return

GuiClose:
;ExitApp
Gui, Hide
return

#IfWinActive ahk_class AutoHotkeyGUI
ESC::
;ExitApp
Gui, Hide
return

#IfWinActive
#4::
Gui, Show
return

SetStatusMessageAndColor(messageText, messageColor)
{
	Gui, Font, c%messageColor%
	GuiControl, Font, StatusMessage
	GuiControl,, StatusMessage, %messageText%
}

ProcessFunction(OutputVarFunctionName, OutputVarInputType){
	GuiControlGet, OutputVarInput,,EditInput,Edit
	
	;MsgBox % "Input Type: " OutputVarInputType
	
	if ( StrLen(OutputVarInput) = 0  and (OutputVarInputType != "N/A") )
	{
		;MsgBox % "User input is blank"
		;return
		;GuiControl,, StatusMessage, "User input is blank."
		SetStatusMessageAndColor("User input field is blank. (ProcessFunction)", "Red")
		return
	}

	ReturnFunctionValue := %OutputVarFunctionName%(OutputVarInput)
	GuiControl,, EditOutput, %ReturnFunctionValue%
	;SetStatusMessageAndColor("Done.", "Green") ; Allow functions to send back their own messages, whether positive or negative.
}



ButtonOK:
	;	"To detect when the user has pressed Enter while a ListView has focus, use a default button (which can be hidden if desired)."
	GuiControlGet, OutputVarFocusedVariable, FocusV	; make sure the "v" flag is set on control element. "g" caused this not to work.
	;MsgBox % "OutputVarFocusedVariable: " OutputVarFocusedVariable
	if (OutputVarFocusedVariable != "MyListView")
		return

	;;MsgBox % "Enter was pressed. The focused row number is " . LV_GetNext(0, "Focused")
	
	; Function|Input|Description|Function Name
	selRow := LV_GetNext(0, "Focused")
	LV_GetText(OutputVarFunctionName, selRow, 4)
	LV_GetText(OutputVarInput, selRow, 2)
	ProcessFunction(OutputVarFunctionName, OutputVarInput)
return

ButtonClear:
	GuiControl,, EditInput,
	GuiControl,, EditOutput,
	SetStatusMessageAndColor("Cleared all inputs.", "Green")
return

ButtonCopy:
	GuiControlGet, OutputVarOutput,,EditOutput,Edit
	OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")	; The "Edit" controls only have `n so need to add `r`n on "Copy"
	Clipboard := OutputVarStrReplace
	SetStatusMessageAndColor("Copied output to Windows Clipboard.", "Green")
return

ButtonMoveUp:
	GuiControlGet, OutputVarOutput,,EditOutput,Edit
	OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")
	GuiControl,, EditInput, %OutputVarStrReplace%
	SetStatusMessageAndColor("Copied output to input.", "Green")
return

ButtonPaste:
	GuiControl,, EditInput, %Clipboard%
	SetStatusMessageAndColor("Pasted Windows Clipboard content to User Input.", "Green")
return

ButtonProcess:
	RowNumber := LV_GetNext(0,"F")
	if (!RowNumber)
	{
		SetStatusMessageAndColor("No selected row found.", "Red")
		return
	}
	
	; Function|Input|Description|Function Name
	LV_GetText(OutputVarInput, RowNumber, 2)
	LV_GetText(OutputVarFunctionName, RowNumber, 4)
	ProcessFunction(OutputVarFunctionName, OutputVarInput)
return

MyListView:
	;MsgBox % "A_GuiEvent: " A_GuiEvent
	if (A_GuiEvent = "DoubleClick")
	{
		;ToolTip You double-clicked row number %A_EventInfo%. Text: %RowText%
		LV_GetText(OutputVarInput, A_EventInfo, 1)
		LV_GetText(OutputVarFunction, A_EventInfo, 2)
		LV_GetText(OutputVarDescription, A_EventInfo, 3)
		LV_GetText(OutputVarFunctionName, A_EventInfo, 4)
		;MsgBox % "OutputVarType: " OutputVarType "`nOutputVarFunction: " OutputVarFunction "`nOutputVarDescription: " OutputVarDescription "`nOutputVarName: " OutputVarName
		
		/*
		GuiControlGet, OutputVarInput,,EditInput,Edit
		if ( StrLen(OutputVarInput) = 0 )
		{
			;MsgBox % "User input is blank"
			SetStatusMessageAndColor("User input field is blank. (MyListView)", "Red")
			return
		}
		*/
		
		ProcessFunction(OutputVarFunctionName, OutputVarInput)
	}
return

ListViewInit:
	; Type, Function, Description
	; 1+ Lines, Convert (Uppercase), Converts text to uppercase, for one or more lines of text
	LV_Add("", "Cmd.exe (Get IPConfig /all)", "N/A", "Runs IPConfig /all and copies it back", "CommandPromptGetIpconfigAll")
	LV_Add("", "Cmd.exe (Get NSLookup)", "1 Line", "Runs NSLookup for server/URL and copies it back", "CommandPromptGetNsLookup")
	LV_Add("", "Cmd.exe (Ping)", "1 Line", "Opens Cmd.exe and runs ping command for server name/URL/IP Address", "CommandPromptPing")
	LV_Add("", "Cmd.exe (SystemInfo)", "N/A", "Opens Cmd.exe and runs systeminfo command", "CommandPromptSystemInfo")
	LV_Add("", "Cmd.exe (Tracert)", "1 Line", "Opens Cmd.exe and runs tracert command for input", "CommandPromptTracert")
	;
	LV_Add("", "Convert (Uppercase)", "1+ Lines", "CONVERTS TEXT TO UPPERCASE, for one or more lines of text", "ConvertUppercase")
	LV_Add("", "Convert (Lowercase)", "1+ Lines", "converts text to lowercase, for one or more lines of text", "ConvertLowercase")
	LV_Add("", "Convert (Title case)", "1+ Lines", "Converts Text to Title Case, for one or more lines of text", "ConvertTitleCase")
	LV_Add("", "Convert (Sentence case)", "1+ Lines", "Converts text to sentence case, for one or more lines of text", "ConvertSentenceCase")
	LV_Add("", "Convert (Invert case)", "1+ Lines", "Inverts the case in a line of text", "ConvertInvertCase")
	LV_Add("", "Convert LF to CRLF", "1+ Lines", "Converts Linefeeds to Carriage Return Linefeeds", "ConvertCarriageReturnLineFeed")
	LV_Add("", "Convert (Strikeout)", "1+ Lines", "Converts Text to Text with Strikeout", "ConvertStrikethrough")
	LV_Add("", "Convert (Underscore)", "1+ Lines", "Converts Text to Text with Underscore", "ConvertUnderscore")
	;
	LV_Add("", "Clipboard (Append)", "1+ Lines", "Appends text to last character in clipboard, without including a space", "ClipboardAppend")
	LV_Add("", "Clipboard (Prepend)", "1+ Lines", "Prepends text to the first character in clipboard, without including a space", "ClipboardPrepend")
	LV_Add("", "Clipboard (Append line)", "1+ Lines", "Appends a new line and then text to the last character in clipboard", "ClipboardAppendAsNewLine")
	LV_Add("", "Clipboard (Prepend line)", "1+ Lines", "Prepends text and then a new line to the first character in clipboard", "ClipboardPrependAsNewLine")
	;
	LV_Add("", "Generate Text (Lorem Ipsum)", "N/A", "Returns sample text (Lorem Ipsum) to use for testing.", "TextGenerateLoremIpsum")
	; Generate Text (GUID)
	;
	LV_Add("", "Char (Ascii)", "Char", "Encodes the leftmost character to ASCII value", "EncodeToAscii")
	LV_Add("", "Date MM/DD/YYYY (Get Full Date)", "1 Line", "Converts MM/DD/YYYY to full date", "TextGetDateFromMMDDYYYY")
	LV_Add("", "JSON (Format)", "1 Line", "Formats JSON notation to pretty print", "TextFormatJson")
	LV_Add("", "Number (Char Code)", "Char", "Encodes a number to Char Code value", "EncodeToCharCode")
	LV_Add("", "Number (Telephone)", "1+ Lines", "Formats matched 10 numbers to a telephone format", "TextFormatTelephoneNumber")
	LV_Add("", "HTML (Encode)", "1+ Lines", "Encodes text to HTML entities (& to &amp;)", "EncodeToHtml")
	LV_Add("", "URL (Encode)", "1+ Lines", "Encodes text to URI/URL values", "EncodeToUrl")
	LV_Add("", "URL (Decode)", "1+ Lines", "Decodes URI text", "DecodeUrl")
	LV_Add("", "URL (Decode, Get Params)", "1 Line", "Decodes URI text and lists URI params", "DecodeUrlAndGetParams")
	LV_Add("", "URL (Download to Text File)", "1 Line", "Downloads internet file to a text file and opens it.", "UrlDownloadToTextFileAndOpen")
	;
	LV_Add("", "Duplicate X Times", "1 Line", "Duplicates text X times with new lines in between", "TextDuplicateXTimes")
	LV_Add("", "Duplicate Increment Number X Times", "1 Line", "Duplicates text X times with new lines, incrementing the last number found each time by 1", "TextDuplicateIncrementLastNumberXTimes")
	;
	LV_Add("", "CSV (Extract 1 Column)", "CSV", "Pick a column to extract CSV row data from", "ExtractColumnDataForCSV")
	LV_Add("", "CSV (Filter 1 Column)", "CSV", "Pick a column to remove CSV row data from", "FilterColumnDataForCSV")
	LV_Add("", "CSV (To <table>)", "CSV", "Convert CSV to html table code", "CsvToHtmlTable")
	LV_Add("", "CSV (Swap 1 Column)", "CSV", "Swap the position of one column's data with another column's position", "SwapColumnForCSV")
	LV_Add("", "TSV (Extract 1 Column)", "TSV", "Pick a column to extract TSV row data from", "ExtractColumnDataForTSV")
	LV_Add("", "TSV (Filter 1 Column)", "TSV", "Pick a column to remove TSV row data from", "FilterColumnDataForTSV")
	LV_Add("", "TSV (Swap 1 Column)", "TSV", "Swap the position of one column's data with another column's position", "SwapColumnForTSV")
	LV_Add("", "*SV (Extract 1 Column)", "*SV", "Pick a column to extract User-specific delimiter-SV row data from", "ExtractColumnDataForAnySV")
	LV_Add("", "*SV (Filter 1 Column)", "*SV", "Pick a column to remove User-specific delimiter-SV row data from", "FilterColumnDataForAnySV")
	LV_Add("", "*SV (Swap 1 Column)", "*SV", "Swap the position of one column's data with another column's position", "SwapColumnForAnySV")
	LV_Add("", "*SV (Spaced Columns)", "*SV", "Print delimiter-SV data in spaced columns based on column data length", "SpacedColumnForAnySV")
	;
	LV_Add("", "Flatten to 1 Line", "1+ Lines", "Condenses to one line - removing duplicate white space and new lines", "TextFlattenSingleLine")
	LV_Add("", "Flatten each line", "1+ Lines", "Trims and removes duplicate whitespace for each line", "TextFlattenMultiLine")
	;
	LV_Add("", "Indent More", "1+ Lines", "Tabs one or more lines of text", "TextFormatIndentMore")
	LV_Add("", "Indent Less", "1+ Lines", "Decreases tabs by one for one or more lines of text", "TextFormatIndentLess")
	;
	LV_Add("", "Join Lines w/ Char", "1+ Lines", "Joins lines of text by user entered separator", "TextJoinLinesWithChar")
	LV_Add("", "Join SV Line w/ Char", "1+ Lines", "Split and join a line of text by user entered separator", "TextJoinSplittedLineWithChar")
	LV_Add("", "Join SV Lines w/ Char", "1+ Lines", "Split and join lines of text by user entered separator", "TextJoinSplittedLinesWithChar")
	;
	LV_Add("", "Replace (AHK RegEx)", "1+ Lines", "Use AHK's RegEx to find and replace all matching text patterns", "TextReplaceRegex")
	LV_Add("", "Replace (AHK RegEx Each Line)", "1+ Lines", "Use AHK's RegEx to find and replace text patterns line by line", "TextReplaceRegexEachLine")
	LV_Add("", "Replace (Chars X Times)", "1+ Lines", "Replace a string of characters up to X occurrences, from left to right", "TextReplaceStringXTimes")
	LV_Add("", "Replace (Chars w/ New Line)", "1+ Lines", "Replace character(s) with Carriage Return New Line feed (CRLF)", "TextReplaceCharsNewLine")
	LV_Add("", "Replace (First Occurrence)", "1+ Lines", "Replace the first occurrence of a string in text block", "TextReplaceFirstStringOccurrence")
	LV_Add("", "Replace (First Occurrence Each Line)", "1+ Lines", "Replace the first occurrence of a string in text block", "TextReplaceFirstStringOccurrenceEachLine")
	LV_Add("", "Replace (Last Occurrence)", "1+ Lines", "Replace the last occurrence of a string in text block", "TextReplaceLastStringOccurrence")
	LV_Add("", "Replace (Last Occurrence Each Line)", "1+ Lines", "Replace the last occurrence of a string in each line", "TextReplaceLastStringOccurrenceEachLine")
	LV_Add("", "Replace (Spaces)", "1+ Lines", "Replaces whitespaces with nothing, shortening the text", "TextRemoveSpaces")
	LV_Add("", "Replace (Spaces w/ Underscores)", "1+ Lines", "Replaces whitespaces in text with underscores", "TextReplaceSpacesUnderscore")
	LV_Add("", "Replace (Spaces w/ Dashes", "1+ Lines", "Replaces whitespaces in text with dashes", "TextReplaceSpacesDash")
	LV_Add("", "Replace (Windows)", "1+ Lines", "Replaces non-Windows filesystem friendly chars with underscores", "TextMakeWindowsFileFriendly")
	;
	LV_Add("", "Reduce (Blank Lines)", "1+ Lines", "Reduces the amount of consecutive blank lines", "RemoveBlankLines")
	LV_Add("", "Reduce Duplicate Lines", "1+ Lines", "Reduces the amount of lines until each line is unique.", "RemoveDuplicateLines")
	LV_Add("", "Reduce (Consecutive Repeating Lines)", "1+ Lines", "Reduces the amount of repeating duplicate lines", "RemoveRepeatingDuplicateLines")
	LV_Add("", "Reduce (Consecutive Repeating Characters)", "1+ Lines", "Reduces the amount of repeating duplicate characters", "RemoveRepeatingDuplicateChars")
	;
	LV_Add("", "List (Ordinal Number)", "1+ Lines", "Prefixes lines of text with ordinal numbers", "TextPrefixOrdinalNumber")
	LV_Add("", "List (Ordinal Lowercase)", "1+ Lines", "Prefixes lines of text with ordinal lowercase letters", "TextPrefixOrdinalLowercase")
	LV_Add("", "List (Ordinal Uppercase)", "1+ Lines", "Prefixes lines of text with ordinal uppercase letters", "TextPrefixOrdinalUppercase")
	;
	LV_Add("", "Trim (Both)", "1+ Lines", "Removes whitespace before and after text", "TextTrim")
	LV_Add("", "Trim (Left)", "1+ Lines", "Removes whitespace before text", "TextTrimLeft")
	LV_Add("", "Trim (Right)", "1+ Lines", "Removes whitespace after text", "TextTrimRight")
	;
	LV_Add("", "Pad (Spaces to Align)", "1+ Lines", "Add spaces to each line to align text to user entered character", "TextCharLineUp")
	LV_Add("", "Pad (Spaces to Left Align)", "1+ Lines", "Add spaces to each line to align text to user entered character", "TextPadLeftCharLineUp")
	LV_Add("", "Pad (Spaces to Left)", "1+ Lines", "Add spaces to the left of each line until length is reached.", "TextPadSpaceLeft")
	LV_Add("", "Pad (Spaces to Right)", "1+ Lines", "Add spaces to the left of each line until length is reached.", "TextPadSpaceRight")
	;
	LV_Add("", "Speak (Start)", "1+ Lines", "Use Windows Speak to say text out loud", "TextSay")
	LV_Add("", "Speak (Stop)", "1+ Lines", "Stop using Windows Speak", "TextSayStop")
	;
	LV_Add("", "Prefix/Suffix (Matching)", "1+ Lines", "Surround text block with matching symbols", "TextSurroundingChar")
	LV_Add("", "Prefix/Suffix Remove", "1+ Lines", "Remove Prefix/Suffix of text from each line", "TextPrefixSuffixRemoveEachLine")
	LV_Add("", "Prefix/Suffix Lines", "1+ Lines", "Add Prefix/Suffix of text to each line", "TextPrefixSuffixEachLine")
	;
	LV_Add("", "HTML (Add Tag)", "1+ Lines", "Surround text with user defined HTML tag", "TextSurroundingHTMLTag")
	LV_Add("", "HTML (Add Tag Each Line)", "1+ Lines", "Surround text on each line with defined HTML tag", "TextSurroundingHTMLTagEachLine")
	;
	; Markdown (to HTML), "1+ Lines", "Convert Markdown (MD) to HTML" -- see Markdown_2_HTML.ahk by JasonDavis for code and GUI application (From Others folder)
	;
	LV_Add("", "Sort (Alpha Asc)", "1+ Lines", "Sort lines of text alphabetically ascendingly", "TextSortAlphabetically")
	LV_Add("", "Sort (Alpha Unique Asc)", "1+ Lines", "Sort unique lines of text alphabetically ascendingly", "TextSortAlphabeticallyUnique")
	LV_Add("", "Sort (Alpha Unique Case Insensitive Asc)", "1+ Lines", "Sort unique lines of text alphabetically, ascendingly, and case sensitive", "TextSortAlphabeticallyUniqueCaseInsensitive")
	LV_Add("", "Sort Line w/ Delimiter (Alpha Asc)", "1+ Lines", "Split line by delimiter and sort alphabetically, ascendingly", "TextSortAlphabeticallyWithDelimiter")
	LV_Add("", "Sort (Number Asc)", "1+ Lines", "Sort lines of numeric text, ascendingly", "TextSortNumericAscending")
	LV_Add("", "Sort (Number Desc)", "1+ Lines", "Sort lines of numeric text, descendingly", "TextSortNumericDescending")
	LV_Add("", "Sort Line w/ Delimiter (Number Asc)", "1+ Lines", "Split line by delimiter and sort numerically, ascendingly", "TextSortNumericWithDelimiter")
	LV_Add("", "Sort (Alpha Desc)", "1+ Lines", "Sort lines of text in reverse alphabetically order, or descendingly", "TextSortReverse")
	LV_Add("", "Sort (Reverse As is)", "1+ Lines", "Sort lines of text in reverse order as it is listed", "TextSortReverseOrder")
	LV_Add("", "Sort Line w/ Delimier (Reverse Alpha Asc)", "1+ Lines", "Split line by delimiter and sort alphabetically, descending", "TextSortReverseWithDelimiter")
	LV_Add("", "Sort (Random)", "1+ Lines", "Sort lines randomly", "TextSortRandom")
	;
	LV_Add("", "Analysis (Count)", "1+ Lines", "Counts length, spaces, tabs, characters, and lines", "TextAnalysisCount")
	;
	LV_Add("", "Windows (Run)", "1 Line", "Use the Windows OS Run command on input.", "WindowsRun")
	;LV_Add("", "Function", "1+ Lines", "Desc", "FunctionName")
	
return

LBTextTypes:
ListBoxTextTypeList =
(LTrim Join|
Line (Single)|
Line (Multiple)
CSV
TSV
)
return

LBTextFunctions:
ListBoxTextFunctionList =
(LTrim Join|
Convert (Uppercase)
)
return


TextAnalysisCount(UserInput) {
	; Mango.ahk
	
	Text := UserInput
	;MsgBox % "UserInput: " Text
	
	len := StrLen(Text)
	;MsgBox % "len: " len
	
	Spaces:=""
	newline:=""
	chars:=""
	Tabs:=""
	
	Loop, Parse, Text, `n, `r
		newline++
	
	Loop, Parse, Text
	{
		If A_loopField=%A_Space%
			Spaces++
		
		If A_loopField=%A_Tab%
			tabs++
	}

	If (tabs=="")
		tabs:=0
	
	If (Spaces=="")
		Spaces:=0

	chars:= len-Spaces-tabs
	
	SetStatusMessageAndColor("Analysis complete.", "Green")
	return "Length: " . len . "`nSpaces: " . spaces . "`nTabs: " . Tabs . "`nCharacters: " . chars . "`nLines: " . newline
	
}

ClipboardAppend(UserInput)
{
	; TESTED?
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly appended data
	SetStatusMessageAndColor("Clipboard + Appeneded data in Output.", "Green")
	return Clipboard . UserInput
}

ClipboardPrepend(UserInput)
{
	; TESTED?
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly prepended data
	SetStatusMessageAndColor("Prepended data + Clipboard to Output.", "Green")
	return UserInput . Clipboard
}
ClipboardAppendAsNewLine(UserInput)
{
	; TESTED?
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly appended data
	SetStatusMessageAndColor("Clipboard + Appeneded data + New Line in Output.", "Green")
	return Clipboard . "`n" . UserInput
}

ClipboardPrependAsNewLine(UserInput)
{
	; TESTED?
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly prepended data
	SetStatusMessageAndColor("Prepended data + New Line + Clipboard to Output.", "Green")
	return UserInput . "`n" . Clipboard
}

HelperCommandPromptClip(Command, Method)
{
	Run, cmd
	WinWait, ahk_class ConsoleWindowClass
	
	if ( Method == 1 )
	{ ; Method == Clip - Copy command result to clipboard and exit command prompt
		SendInput, %Command% | clip
		Sleep 1000
		Send, {Enter}
		ClipWait,,1
		SendInput, exit
		Sleep, 100
		Send, {Enter}
		Sleep, 100
	
	} else {
		; Run command in commmand prompt and do not return
		; - Helpful when it is not clear how long the command will take to process
		; - Examples: SystemInfo, Ping, Tracert, ...
		SendInput, %Command%
		Sleep 1000
		Send, {Enter}
	}
	
	return
}

CommandPromptSystemInfo()
{
	Command := "systeminfo"
	HelperCommandPromptClip(Command, 0)
	SetStatusMessageAndColor("Ran systeminfo in Command Prompt.", "Green")
	return
}

CommandPromptTracert(UserInput)
{
	Command := "tracert " . UserInput
	HelperCommandPromptClip(Command, 0)
	SetStatusMessageAndColor("Ran tracert in Command Prompt.", "Green")
	return
}

CommandPromptPing(UserInput)
{
	; Command Prompt: Ping
	Command := "ping " . UserInput
	HelperCommandPromptClip(Command, 0)
	SetStatusMessageAndColor("Ran ping in Command Prompt.", "Green")
	return
}

CommandPromptGetIpconfigAll()
{
	; Command Prompt: IPCONFIG info
	Command := "ipconfig /all"
	HelperCommandPromptClip(Command, 1)
	;MsgBox % clipboard
	SetStatusMessageAndColor("Ran Ipconfig in Command Prompt.", "Green")
	return clipboard
}

CommandPromptGetNsLookup(UserInput)
{
	Command := "nslookup " . UserInput
	HelperCommandPromptClip(Command, 1)
	SetStatusMessageAndColor("Ran NSLookup in Command Prompt.", "Green")
	return clipboard
}

ConvertUppercase(UserInput)
{
	; Converts all text to uppercase text
	StringUpper, OutputVarStringUpper, UserInput
	SetStatusMessageAndColor("Converted input to Uppercase characters.", "Green")
	return OutputVarStringUpper
}

ConvertLowercase(UserInput)
{
	; Converts all text to lowercase text
	StringLower, OutputVarStringLower, UserInput
	SetStatusMessageAndColor("Converted input to Lowercase characters.", "Green")
	return OutputVarStringLower
}

ConvertTitleCase(UserInput)
{
	; Converts all text/sentences to title case
	StringUpper, OutputVarStringUpper, UserInput, T	; "T" for Title Case
	SetStatusMessageAndColor("Converted input to Title case.", "Green")
	return OutputVarStringUpper
}

ConvertSentenceCase(UserInput)
{
	; SENTENCE CASE - quick brown fox. went home. = Quick brown fox. Went home.
	Text := UserInput
	Text := RegExReplace(Text, "([.?\s!]\s\w)|^(\.\s\b\w)|^(.)", "$U0")
	SetStatusMessageAndColor("Converted input to Sentence case.", "Green")
	return Text
}

ConvertInvertCase(UserInput) {
	; Thanks JDN ; Mango.ahk
	
	Text := UserInput
	Lab_Invert_Char_Out:= ""
	;Loop % Strlen(Clipboard) { ;%
	Loop % Strlen(Text) {
		;Lab_Invert_Char:= Substr(Clipboard, A_Index, 1)
		Lab_Invert_Char:= Substr(Text, A_Index, 1)
		If Lab_Invert_Char is Upper
			Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) + 32)
		Else If Lab_Invert_Char is Lower
			Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) - 32)
		Else
			Lab_Invert_Char_Out:= Lab_Invert_Char_Out Lab_Invert_Char
	}
	;Clipboard:=Lab_Invert_Char_Out
	Text:=Lab_Invert_Char_Out
	SetStatusMessageAndColor("Converted input to inverted case.", "Green")
	return Text
}

ConvertCarriageReturnLineFeed(UserInput)
{
	; Turn `r to `r`n, in order to allow to work with Notepad and other Windows programs
	Text := UserInput
	Text := RegExReplace(Text, "`r", "`r`n")
	return Text
}

ConvertStrikethrough(UserInput)
{
	; Strikeout with Unicode('\u0336') - borrowed from Text Tools (https://chrome.google.com/webstore/detail/text-tools/mpcpnbklkemjinipimjcbgjijefholkd)
	vText := UserInput
	vSplit := StrSplit(vText, "")
	for k,v in vSplit
		vSplit[k] := v . "̶"
	vText := Join(vSplit, "")
	;MsgBox % vText
	return vText
}

ConvertUnderline(UserInput)
{
	; Underscore with Unicode ('\u0332') - borrowed from TextTools (https://chrome.google.com/webstore/detail/text-tools/mpcpnbklkemjinipimjcbgjijefholkd)
	vText := UserInput
	vSplit := StrSplit(vText, "")
	for k,v in vSplit
		vSplit[k] := v . "̲"
	vText := Join(vSplit, "")
	;MsgBox % vText
	return vText
}

CsvToHtmlTable(UserInput)
{
	; CSV -> <table>...</table>
	isFirstRowHeader := false
	isHeaderAndFooterSame := false

	MsgBox, 4,, Is the first row also the header? (Press Yes or No)
	IfMsgBox Yes
		isFirstRowHeader := true

	if (isFirstRowHeader)
	{
		MsgBox, 4,, Is the header data also the footer data? (Press Yes or No)
		IfMsgBox Yes
			isHeaderAndFooterSame := true
	}

	;data := "a,b,c`r`nd,e,f`r`ng,h,i"
	data := UserInput
	delim := ","
	arrCsvData := HelperGetSVDataIntoRowsArray(data, delim)

	tagTableOpen := "<table>"
	tagTableClose := "</table>"
	tagTheadOpen := "<thead>"
	tagTheadClose := "</thead>"
	tagTrOpen := "<tr>"
	tagTrClose := "</tr>"
	tagThOpen := "<th>"
	tagThClose := "</th>"
	tagTbodyOpen := "<tbody>"
	tagTbodyClose := "</tbody>"
	tagTdOpen := "<td>"
	tagTdClose := "</td>"
	tagTfootOpen := "<tfoot>"
	tagTfootClose := "</tfoot>"

	;newLine := "`r`n"
	dataThead := ""
	dataTbody := ""

	if (isFirstRowHeader)
	{
		dataThead .= tagTrOpen
		
		for kCol, vCol in arrCsvData[1]
		{
			dataThead .= tagThOpen
			dataThead .= vCol
			dataThead .= tagThClose
		}
		
		dataThead .= tagTrClose
	}

	for kRow, vRow in arrCsvData
	{
		dataTbody .= tagTrOpen
		
		for kCol, vCol in arrCsvData[A_Index]
		{
			if (isFirstRowHeader and kRow == 1)
				continue ; Skip the first row's data if it's intended to be displayed in the header row
			
			dataTBody .= tagTdOpen
			dataTBody .= vCol
			dataTBody .= tagTdClose
		}
		
		dataTbody .= tagTrClose
		;dataTbody .= newLine ; comment after testing
	}

	htmlTable := ""
	htmlTable .= tagTableOpen
	htmlTable .= tagTheadOpen

	if (isFirstRowHeader)
		htmlTable .= dataThead

	htmlTable .= tagTheadClose
	htmlTable .= tagTbodyOpen
	htmlTable .= dataTbody ; <tr><td>..</td></tr>`r`n<tr><td>..</td>`r`n ...
	htmlTable .= tagTbodyClose

	if (isHeaderAndFooterSame)
	{
		htmlTable .= tagTfootOpen
		htmlTable .= dataThead
		htmlTable .= tagTfootClose
	}

	htmlTable .= tagTableClose

	;MsgBox % htmlTable
	SetStatusMessageAndColor("Converted CSV to <table>.", "Green")
	return htmlTable
}


EncodeToAscii(UserInput)
{
	Transform, Text, ASc, %UserInput% ; Encode ASCII
	SetStatusMessageAndColor("Encoded input (ASCII).", "Green")
	return Text
}

EncodeToHtml(UserInput)
{
	Transform, Text, HTML, %UserInput% ; Encode HTML
	SetStatusMessageAndColor("Encoded input (HTML).", "Green")
	return Text
}

EncodeToCharCode(UserInput) {
	
	Text := UserInput
	
	if Text is not Integer
	{
		;MsgBox % "User input must be in the form of an integer."
		SetStatusMessageAndColor("User input must be in the form of an integer. (EncodeToCharCode)", "Red")
		return
	}
		
	Chr := Chr(Text)
	SetStatusMessageAndColor("Encoded input (CharCode).", "Green")
	return Chr
}

uriDecode(str) {
;https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/
	Loop
		If RegExMatch(str, "i)(?<=%)[\da-f]{1,2}", hex)
			StringReplace, str, str, `%%hex%, % Chr("0x" . hex), All
		Else Break
	
	SetStatusMessageAndColor("Decoded input (URI).", "Green")
	Return, str
}

UriEncode(Uri, RE="[0-9A-Za-z]"){
;https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/
	VarSetCapacity(Var,StrPut(Uri,"UTF-8"),0),StrPut(Uri,&Var,"UTF-8")
	While Code:=NumGet(Var,A_Index-1,"UChar")
		Res.=(Chr:=Chr(Code))~=RE?Chr:Format("%{:02X}",Code)
	
	SetStatusMessageAndColor("Encoded input (URI).", "Green")
	Return,Res  
}

EncodeToUrl(UserInput)
{
	Text := UserInput
	Text := uriEncode(Text) ; Encode URI
	
	SetStatusMessageAndColor("Encoded input (URL).", "Green")
	return Text
}

ExtractColumnDataForCSV(UserInput)
{
	Text := UserInput
	delim := ","
	Text := HelperExtractColumnDataForAnySV(Text, delim)
	
	SetStatusMessageAndColor("Extracted 1 Column's data (CSV).", "Green")
	return Text
}

ExtractColumnDataForTSV(UserInput)
{
	Text := UserInput
	delim := "`t"
	Text := HelperExtractColumnDataForAnySV(Text, delim)
	
	SetStatusMessageAndColor("Extracted 1 Column's data (TSV).", "Green")
	return Text
}

ExtractColumnDataForAnySV(UserInput)
{
	Prompt := "What character separates the data?"
	ErrorLevel =
	
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (strlen(OutputVar) < 1 and not ErrorLevel)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	if ( ErrorLevel )
		return ; User canceled
	
	Text := UserInput
	delim := OutputVar
	Text := HelperExtractColumnDataForAnySV(Text, delim)
	
	SetStatusMessageAndColor("Extracted 1 Column's data (*SV).", "Green")
	return Text
}

HelperGetSVDataIntoRowsArray(data, delim)
{
	arrRows := Object()
	
	; Below loop/parse works for all tested scenarios: `r`n, `n, and a single line of text
	Loop, Parse, data, `n, `r
	{
		arrSplitted := StrSplit(A_LoopField, delim)
		
		if ( arrSplitted.Length() == 0 or arrSplitted == )
			continue ; Even a non-delimiter value returns something so return if somehow we have nothing.
		
		arrRows[A_Index] := arrSplitted
	}
	
	return arrRows
}

HelperExtractColumnDataForAnySV(data, delim)
{
	arrRows := HelperGetSVDataIntoRowsArray(data, delim)

	; Ask user to pick a column to extract
	intColumns := arrRows[1].Length()

	if ( intColumns < 1 )
		return

	SampleRow := Join(arrRows[1], ",")
	Prompt := "Extract Column: Pick a number between 1 and " . intColumns . "`r`n`r`nHere is a Sample row: `r`n" . SampleRow

	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar = " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}

	NewLine := "`r`n"
	retValue := 
	
	for indexRow, elementRow in arrRows
		retValue .= NewLine elementRow[OutputVar]

	retValue := SubStr(retValue, StrLen(NewLine) + 1)
	return retValue
}

SwapColumnForCSV(UserInput)
{
	; Allow the user to swap the position of a columns data with another one.
	; 1.) Ask the user which column numbers to swap
	; 2.) Rebuild the CSV data based on the user's input
	
	vText := UserInput
	vDelim := ","
	vText := HelperSwapColumnDataForAnySV(vText, vDelim)
	
	SetStatusMessageAndColor("Swap 1 Column's data (CSV).", "Green")
	return vText
}

SwapColumnForTSV(UserInput)
{
	vText := UserInput
	vDelim := "`t"
	vText := HelperSwapColumnDataForAnySV(vText, vDelim)
	
	SetStatusMessageAndColor("Swap 1 Column's data (TSV).", "Green")
	return vText
}

SwapColumnForAnySV(UserInput)
{
	Prompt := "What character separates the data?"
	ErrorLevel =
	
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (strlen(OutputVar) < 1 and not ErrorLevel)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	if ( ErrorLevel )
		return ; User canceled
	
	vText := UserInput
	vDelim := OutputVar
	vText := HelperSwapColumnDataForAnySV(vText, vDelim)
	
	SetStatusMessageAndColor("Swap 1 Column's data (*SV).", "Green")
	return vText
}

HelperSwapColumnDataForAnySV(data, delim)
{
	arrRows := HelperGetSVDataIntoRowsArray(data, delim)

	; Ask user to pick a column to extract
	intColumns := arrRows[1].Length()

	if ( intColumns < 2 )
		return

	SampleRow := Join(arrRows[1], ",")
	
	Prompt := "Select Column to move: Pick a number between 1 and " . intColumns . "`r`n`r`nHere is a Sample row: `r`n" . SampleRow
	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar == " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	columnMove := OutputVar
	
	Prompt := "Select Column to be swapped: Pick a number between 1 and " . intColumns . "`r`n`r`nHere is a Sample row: `r`n" . SampleRow
	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar == " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}

	columnSwap := OutputVar
	
	if ( columnMove == columnSwap )
		return ; Cannot move column to its original position
	
	NewLine := "`r`n"
	retValue := 
	
	for indexRow, elementRow in arrRows
	{
		retValue .= NewLine 
		
		for indexColumn, elementColumn in arrRows[indexRow]
		{
			if (indexColumn == columnMove )
			{
				retValue .= elementRow[columnSwap] . delim
			} else if (indexColumn == columnSwap)
			{
				retValue .= elementRow[columnMove] . delim
			} else 
				retValue .= elementRow[indexColumn] . delim
		}
		retValue := Substr(retValue, 1, StrLen(retValue) - StrLen(delim)) ; Remove last delimiter added
	}
	
	retValue := SubStr(retValue, StrLen(NewLine) + 1) ; Remove initial NewLine
	return retValue
}


FilterColumnDataForCSV(UserInput)
{
	Text := UserInput
	delim := ","
	Text := HelperFilterColumnDataForAnySV(Text, delim)
	
	SetStatusMessageAndColor("Filtered 1 Column's data (CSV).", "Green")
	return Text
}

FilterColumnDataForTSV(UserInput)
{
	Text := UserInput
	delim := "`t"
	Text := HelperFilterColumnDataForAnySV(Text, delim)
	
	SetStatusMessageAndColor("Filtered 1 Column's data (TSV).", "Green")
	return Text
}

FilterColumnDataForAnySV(UserInput)
{
	Prompt := "What character separates the data?"
	ErrorLevel =
	
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (strlen(OutputVar) < 1 and not ErrorLevel)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	if ( ErrorLevel )
		return ; User canceled
	
	Text := UserInput
	delim := OutputVar
	Text := HelperFilterColumnDataForAnySV(Text, delim)
	
	SetStatusMessageAndColor("Filtered 1 Column's data (*SV).", "Green")
	return Text
}

HelperFilterColumnDataForAnySV(data, delim)
{
	arrRows := HelperGetSVDataIntoRowsArray(data, delim)

	; Ask user to pick a column to filter out
	intColumns := arrRows[1].Length()

	if ( intColumns < 1 )
		return

	SampleRow := Join(arrRows[1], delim)
	Prompt := "Filter Column: Pick a number between 1 and " . intColumns . "`r`n`r`nHere is a Sample row: `r`n" . SampleRow

	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar = " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}

	NewLine := "`r`n"
	retValue := 
	
	for indexRow, elementRow in arrRows
	{
		retValue .= NewLine 
		for indexColumn, elementColumn in arrRows[indexRow]
		{
			if (indexColumn == OutputVar)
				continue
			else
				retValue .= elementRow[indexColumn] . delim
		}
		retValue := Substr(retValue, 1, StrLen(retValue) - StrLen(delim)) ; Remove last delimiter added
	}
	
	retValue := SubStr(retValue, StrLen(NewLine) + 1) ; Remove initial NewLine
	return retValue
}

HelperFindColumnMaxWidth(arrData)
{
	; Find the max width per column
	colMaxWidth := Object()

	Loop, % arrData.MaxIndex()
	{ ; Rows
		rIndex := A_Index ; row Index
		
		Loop, % arrData[A_Index].MaxIndex()
		{ ; Columns
			cIndex := A_Index ; column index
			
			if ( rIndex == 1 )
				colMaxWidth[cIndex] := StrLen(arrData[rIndex][cIndex])
			else if (colMaxWidth[cIndex] < StrLen(arrData[rIndex][cIndex]) )
				colMaxWidth[cIndex] := StrLen(arrData[rIndex][cIndex])
		}
	}
	
	return colMaxWidth
}

HelperPrintSpacedColumns(arrData, colMaxWidth, hasHeader)
{
	colBetweenSpace := 2
	@ :=
	NewLine := "`r`n"

	Loop, % arrData.MaxIndex()
	{ ; Rows
		rIndex := A_Index ; row Index
		@ .= NewLine
		
		if ( hasHeader and A_Index == 2 )
		{
			; Print ----- for columns
			Loop, % arrData[A_Index].MaxIndex()
			{
				separator := 
				loop % colMaxWidth[A_Index]
					separator .= "-"
				
				if ( A_Index == arrData[rIndex].MaxIndex() )
					@ .= separator
				else {
					vLength := colMaxWidth[A_Index] + colBetweenSpace
					vText := separator
					vDirection := "Right"
					@ .= HelperPadSpaces(vLength, vText, vDirection)
				}
			}
			
			@.= NewLine
		}
		
		Loop, % arrData[A_Index].MaxIndex()
		{ ; Columns
			cIndex := A_Index ; column index
			
			if ( A_Index == arrData[rIndex].MaxIndex() ){ ; Do not add padding if last column
				@ .= arrData[rIndex][cIndex]
			
			} else {			
				vLength := colMaxWidth[cIndex] + colBetweenSpace
				vText := arrData[rIndex][cIndex]
				vDirection := "Right"
				@ .= HelperPadSpaces(vLength, vText, vDirection)
			}
		}
	}

	return SubStr(@, StrLen(NewLine)+1)
}

SpacedColumnForAnySV(UserInput)
{
	Prompt := "What character separates the data?"
	ErrorLevel =
	
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (strlen(OutputVar) < 1 and not ErrorLevel)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	if ( ErrorLevel )
		return ; User canceled
	
	MsgBox, 4, Data Table, First row header row? (Yes or No)
	IfMsgBox Yes
		hasHeader := true
	
	vText := UserInput
	vDelim := OutputVar
	arrData := HelperGetSVDataIntoRowsArray(vText, vDelim)
	colMaxWidth := HelperFindColumnMaxWidth(arrData)
	
	SetStatusMessageAndColor("Done.", "Green")
	
	if ( hasHeader )
		return HelperPrintSpacedColumns(arrData, colMaxWidth, true)
	else
		return HelperPrintSpacedColumns(arrData, colMaxWidth, false)
}

DecodeUrl(UserInput)
{
	Text := UserInput
	Text := uriDecode(Text) ; Decode URI
	
	SetStatusMessageAndColor("Decoded input (URL).", "Green")
	return Text
}

DecodeUrlAndGetParams(UserInput) {
	;Decode URL and Wrap on parameters
	;http://the-automator.com/parse-url-parameters/
	
	Text := UserInput
	Text := uriDecode(Text) 
	StringReplace,Text,Text,?,`r`n`t?,All ;Line break and tab indent <strong>parse URL parameters</strong>.
	StringReplace,Text,Text,&,`r`n`t`t&,All ;Line break and double tab indent
	
	SetStatusMessageAndColor("Decoded input and got params (URL).", "Green")
	return Text
}

TextDuplicateXTimes(UserInput){
	; Input: (1) Single line of text, (2) Number of times to duplicate
	; Output: Duplicated text (on multiple lines)
	
	;If ("" <> Text := Clip()) {
	Text := UserInput
		
		Lines := StrSplit(Text, "`r`n")
		
		if (Lines.Length() > 1) {
			;MsgBox,,SelectedTextDuplicateXTimes,Please only select/copy one line of text for processing. Thank you.
			SetStatusMessageAndColor("Select/copy one line of text for processing.","Red")
			return
		}

		Prompt = "Increment how many times?"
		InputBox, varTimes, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

		if varTimes is integer
		{
			if ( varTimes > 0 )
			{
				NewLine := "`r`n"
				@ := Text
				
				Loop % varTimes
				{
					@ .= NewLine Text
				}
				
				;Clip(@)
				SetStatusMessageAndColor("Input duplicated X times.", "Green")
				return @
			}
		} else 
		{
			;MsgBox % "Not a valid number."
			SetStatusMessageAndColor("Not a valid number.","Red")
			return
		}
	;}
}

TextDuplicateIncrementLastNumberXTimes(UserInput){
	; Input: Single line of text
	; Output: Duplicated text, where the last number is incremented x amount of times.
	
	;If ("" <> Text := Clip()) {
	Text := UserInput
		
		Lines := StrSplit(Text, "`r`n")
		
		if (Lines.Length() > 1) {
			;MsgBox,,SelectedTextDuplicateIncrementLastNumberXTimes,Please only select/copy one line of text for processing. Thank you.
			SetStatusMessageAndColor("Select/copy one line of text for processing","Red")
			return
		}
		
		; DEBUG
		;Text := "C:\Path\To\File\12345\12345_0001.tif"

		; Regular expression used: [1-9]+[0-9]*
		;  - To avoid counting padded zeroes, we'll check for numbers above zero and
		;    then count the zeroes thereafter as part of the pattern to match
		;  - When we replace the number this way, we keep the padded zeroes previously
		;    found in the string
		;  - In this future, we may have to account for padded number/max char length
		Matches := []
		MatchPositions := []
		pos = 1
		While pos := RegExMatch(Text,"[1-9]+[0-9]*", match, pos+StrLen(match))
		{
			Matches.Push(match)
			MatchPositions.Push(pos)
		}

		Prompt = "Increment how many times?"
		InputBox, varTimes, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

		if varTimes is integer
		{
			if ( varTimes > 0 )
			{
				Replacement := Matches[Matches.MaxIndex()] + 0
				LastPos := MatchPositions[MatchPositions.MaxIndex()]
				NewLine := "`r`n"
				@ := Text
				
				Loop % varTimes
				{
					Replacement += 1
					strNew := RegExReplace(Text, "[\d]+", Replacement, , , LastPos)
					@ .= NewLine strNew
				}
				
				;@ := SubStr(@, StrLen(NewLine) + 1)
				;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
				;Clip(SubStr(@, StrLen(NewLine) + 1))
				;Clip(@)
				SetStatusMessageAndColor("Duplicated and incremented input X times.", "Green")
				return @
			}
		} else 
		{
			;MsgBox % "Not a valid number."
			SetStatusMessageAndColor("Not a valid number.","Red")
			return
		}
	;}
}

TextFlattenSingleLine(UserInput)
{
	; Flatten to a single line (Trim, More than one space, tabs, New lines, Carriage Returns)
	; 1.) Trim (Preceeding and exceeding white spaces)
	; 2.) RegExReplace (More than one space, tabs, New lines, Carriage Returns)
	
	Text := Trim(UserInput)
	Text := RegExReplace(Text, "\s", " ") ; Convert any whitespace char (space, tab, newline)
	Text := RegExReplace(Text, "\s+", " ") ; Converts grouped spaces (1 or more)
	
	SetStatusMessageAndColor("Flattened input to one line.", "Green")
	return Text
}

TextFlattenMultiLine(UserInput) {
	; Flatten each line (Trim, More than one space, tabs, New lines, Carriage Returns)
	; 1.) Trim (Preceeding and exceeding white spaces)
	; 2.) RegExReplace (More than one space, tabs, New lines, Carriage Returns)
	
	Text := UserInput
	NewLine := "`r`n"
	@ := ""
	Loop, Parse, Text, `n, `r
	{
		Text := Trim(A_LoopField)
		Text := RegExReplace(Text, "\s", " ") ; Convert any whitespace char (space, tab, newline)
		Text := RegExReplace(Text, "\s+", " ") ; Converts grouped spaces (1 or more)
		@ .= NewLine Text
	}
		
	;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
	SetStatusMessageAndColor("Flattened input each line.", "Green")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextFormatJson(UserInput)
{	
	;vFormatted := JSON_Beautify(UserInput)
	vFormatted := JSON_Beautify(UserInput, "  ")
	;MsgBox % vFormatted
	
	SetStatusMessageAndColor("Formatted (JSON).", "Green")
	return vFormatted
}

TextFormatIndentMore(UserInput)
{
	; https://autohotkey.com/board/topic/70404-clip-send-and-retrieve-text-using-the-clipboard/
	; Indent a block of text. (This introduces a small delay for typing a normal tab; to avoid this you can use ^Tab / ^+Tab as a hotkey instead.)
	
	Text := UserInput
	TabChar := A_Tab ; this could be something else, say, 4 spaces
	NewLine := "`r`n"
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r
			@ .= NewLine (InStr(A_ThisHotkey, "+") ? SubStr(A_LoopField, (InStr(A_LoopField, TabChar) = 1) * StrLen(TabChar) + 1) : TabChar A_LoopField)
		
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		SetStatusMessageAndColor("Indented more.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;} Else
	;	Send % (InStr(A_ThisHotkey, "+") ? "+" : "") "{Tab}"
}

TextFormatIndentLess(UserInput){
	
	Text := UserInput
	TabChar := A_Tab
	NewLine := "`r`n"
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r
		{
			; If we find a tab at the beginning of the string, substring remove it to "indent less"
			FoundPos := RegExMatch(A_LoopField, "O)(^\t)", RegExMatchOutputVar)
			if (RegExMatchOutputVar.Count() > 0)
			{
				@ .= NewLine
				@ .= SubStr(A_LoopField, StrLen(A_Tab)+1)
			}
			else
			{
				@ .= NewLine A_LoopField
			}
		}
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		SetStatusMessageAndColor("Indented less.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

TextFormatTelephoneNumber(UserInput)
{
	; Formats the Clipboard value to (###) ###-#### format
	; Ex: 1234567890 --> (123) 456-7890
	
	/*
	Text := UserInput
	
	;FoundPos := RegExMatch(Clipboard, "O)(\d{10})", RegExMatchOutputVar)
	FoundPos := RegExMatch(Text, "O)(\d{10})", RegExMatchOutputVar)
	if ( RegExMatchOutputVar.Count() > 0 )
	{
		
		strTelephoneNumber := RegExMatchOutputVar.Value(0)
		strTelephoneNumber := "(" . SubStr(strTelephoneNumber, 1, 3) . ") " . SubStr(strTelephoneNumber, 4, 3) . "-" . Substr(strTelephoneNumber, 7, 4)
		;MsgBox % strTelephoneNumber " has been added to the Windows Clipboard."
		;Clipboard := strTelephoneNumber
		return strTelephoneNumber
	}
	*/
	
	;vText := "0123456789`n9876543210`n876374827"
	vText := UserInput
	vFormatted := RegExReplace(vText, "(\d{1})(\d{3})(\d{3})(\d{3})", "$1-($2) $3-$4")
	vFormatted := RegExReplace(vFormatted, "(\d{3})(\d{3})(\d{3})", "($1) $2-$3")
	;MsgBox % vFormatted
	
	SetStatusMessageAndColor("Formatted numbers (Telephone).", "Green")
	return vFormatted
	
}

TextGetDateFromMMDDYYYY(UserInput){
	; Get the full date from mm/dd/yyyy
	; Ex: 1/1/2010 = Friday, January 1, 2010
	;datestring := "1/1/2010"
	
	;strDate := Clipboard
	Text := UserInput
	;FoundPos := RegExMatch(strDate, "O)(\d+\/\d+\/\d+)", RegExMatchOutputVar)
	FoundPos := RegExMatch(Text, "O)(\d+\/\d+\/\d+)", RegExMatchOutputVar)
	if ( RegExMatchOutputVar.Count() > 0 ) {
		;MsgBox % RegExMatchOutputVar.Value(0)
		SetFormat, float, 02.0
		StringSplit, d, strDate, / 
		FormatTime, FullDate, % d3 . d1+0. . d2+0., dddd, MMMM d, yyyy
		;MsgBox % FullDate " has been added to the Windows Clipboard."
		;Clipboard := FullDate
		SetStatusMessageAndColor("Got Full Date.", "Green")
		return FullDate
	}
	else
	{
		;Msgbox % "No date found in Windows Clipboard to process. Aborted."
		SetStatusMessageAndColor("No date found in Windows Clipboard to process. Aborted.","Red")
	}
		
}

TextGenerateLoremIpsum()
{
	SetStatusMessageAndColor("Generated text (Lorem Ipsum).", "Green")
	return "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
}

TextJoinLinesWithChar(UserInput) {
	; 1.) Get content (Text) and put into clipboard
	; 2.) Prompt user for split delimiter
	; 3.) Prompt user for join delimiter
	; 4.) Loop by new line
	;      - For each line, split and join
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
					
		Prompt = "Join lines by..."
		InputBox, varJoiner, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		if ( StrLen(varJoiner) > 0 ) {
			
			@ :=
			
			Loop, Parse, Text, `n, `r 
			{
				@ .= varJoiner A_LoopField
			}
			
			;Clipboard := Substr(@, StrLen(varJoiner) + 1)
			retValue := Substr(@, StrLen(varJoiner) + 1) ; Remove last character join
			
			SetStatusMessageAndColor("Joined lines with character.", "Green")
			return retValue
		}
	;}
}

Join(Array, Sep)
{
	for k, v in Array
		out .= Sep . v
	return SubStr(out, 1+StrLen(Sep))
}

TextJoinSplittedLineWithChar(UserInput) {
	; 1.) Get content (Text) and put into clipboard
	; 2.) Prompt user for split delimiter
	; 3.) Prompt user for join delimiter
	; 4.) Loop by new line
	;      - For each line, split and join
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		
		Prompt = "Split text by..."
		InputBox, varSplitter, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		if ( StrLen(varSplitter) > 0 ) {
			
			Prompt = "Join lines by..."
			InputBox, varJoiner, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
			
			if ( StrLen(varJoiner) > 0 ) {
				
				TextArray := StrSplit(Text, varSplitter)
				@ := Join(TextArray, varJoiner)
				
				;Clipboard := @
				SetStatusMessageAndColor("Joined splitted line with character.", "Green")
				return @
			}
		}
	;}
}

TextJoinSplittedLinesWithChar(UserInput) {
	; 1.) Get content (Text) and put into clipboard
	; 2.) Prompt user for split delimiter
	; 3.) Prompt user for join delimiter
	; 4.) Loop by new line
	;      - For each line, split and join
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		
		Prompt = "Split text by..."
		InputBox, varSplitter, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		if ( StrLen(varSplitter) > 0 ) {
			
			Prompt = "Join lines by..."
			InputBox, varJoiner, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
			
			if ( StrLen(varJoiner) > 0 ) {
				
				NewLine := "`r`n"
				@ :=
				
				Loop, Parse, Text, `n, `r 
				{
					TextArray := StrSplit(A_LoopField, varSplitter)
					@ .= NewLine Join(TextArray, varJoiner)
				}
				
				;Clip(SubStr(@, StrLen(NewLine) + 1))
				;Clipboard := SubStr(@, StrLen(NewLine) + 1)
				SetStatusMessageAndColor("Joined splitted lines with character.", "Green")
				return SubStr(@, StrLen(NewLine) + 1)
			}
		}
	;}
}

TextMakeWindowsFileFriendly(UserInput) {
	; "A file name can't contain any of the following characters: \ / : * ? " < > |
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r 
		{
			tempText := A_LoopField
			; To escape a double quote ("), use two back-to-back so ""
			tempText := RegExReplace(tempText, "/", "_")
			tempText := RegExReplace(tempText, "\\", "_")
			tempText := RegExReplace(tempText, ":", "_")
			tempText := RegExReplace(tempText, "\*", "_")
			tempText := RegExReplace(tempText, "\?", "_")
			tempText := RegExReplace(tempText, """", "_")
			tempText := RegExReplace(tempText, "<", "_")
			tempText := RegExReplace(tempText, ">", "_")
			tempText := RegExReplace(tempText, "\|", "_")
			@ .= tempText
	    }
	
		;Clip(SubStr(@, StrLen(tempText) + 1), 2)
		SetStatusMessageAndColor("Characters are now Windows Filesystem friendly.", "Green")
		return SubStr(@, StrLen(tempText) + 1)
	;}
}

TextPrefixOrdinalNumber(UserInput) {
	; Prefix each line of text with an ordinal number
	;  Purpose:
	;   Make a simple text listing wherever needed.
	;  
	;  * Also removes blank lines as the return value is built during loop process.
	;
	;  Conditions:
	;   A.) If there is text on the line, put prefix right in front of text (indented text for example)
	;   B.) If there is no text on the line, skip it. (Create another function to number blank lines if needed, add "Blank" to the end of the name)
	;
	;  Ex:
	;   Text Line Item\r\n  --> 1. Text Line Item
	;   Text Line Item\r\n  --> 2. Text Line Item
	;   Text Line Item\r\n  --> 3. Text Line Item
	
	NewLine := "`r`n"
	Numeral := 1 ; Ordinal number listing starts with 1
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r 
		{
			; First see if there is text. Then, prefix the text part only (for example, indented text)
			MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar)
			if ( RegExMatchOutputVar.Count() > 0 )
			{	
				OrdinalText := Numeral . ". "
				;@ .= NewLine Numeral ". " A_LoopField
				@ .= NewLine
				@ .= RegExReplace(A_LoopField, "([\S].+)", OrdinalText . "$1" )
				Numeral += 1
			}
	    }
	
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		;Clip(@,2)
		SetStatusMessageAndColor("Done.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

TextPrefixOrdinalLowercase(UserInput) {
	; Prefix each line of text with an ordinal lowercase letter (a-z or 97-122)
	;  Purpose:
	;   Make a simple text listing wherever needed.
	;  
	;  * Also removes blank lines as the return value is built during loop process.
	;
	;  Conditions:
	;   A.) If there is text on the line, put prefix right in front of text (indented text for example)
	;   B.) If there is no text on the line, skip it. (Create another function to number blank lines if needed, add "Blank" to the end of the name)
	;
	;  Ex:
	;   Text Line Item\r\n  --> a. Text Line Item
	;   Text Line Item\r\n  --> b. Text Line Item
	;   Text Line Item\r\n  --> c. Text Line Item
	
	NewLine := "`r`n"
	Numeral := 97 ; Ordinal number listing starts with a
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r 
		{
			; First see if there is text. Then, prefix the text part only (for example, indented text)
			MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar)
			if ( RegExMatchOutputVar.Count() > 0 )
			{	
				OrdinalText := Chr(Numeral) . ". "
				;@ .= NewLine Numeral ". " A_LoopField
				@ .= NewLine
				@ .= RegExReplace(A_LoopField, "([\S].+)", OrdinalText . "$1" )
				Numeral += 1
			}
	    }
	
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		;Clip(@,2)
		SetStatusMessageAndColor("Done.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

TextPrefixOrdinalUppercase(UserInput) {
	; Prefix each line of text with an ordinal uppercase letter (A-Z or 65-90)
	;  Purpose:
	;   Make a simple text listing wherever needed.
	;  
	;  * Also removes blank lines as the return value is built during loop process.
	;
	;  Conditions:
	;   A.) If there is text on the line, put prefix right in front of text (indented text for example)
	;   B.) If there is no text on the line, skip it. (Create another function to number blank lines if needed, add "Blank" to the end of the name)
	;
	;  Ex:
	;   Text Line Item\r\n  --> a. Text Line Item
	;   Text Line Item\r\n  --> b. Text Line Item
	;   Text Line Item\r\n  --> c. Text Line Item
	
	NewLine := "`r`n"
	Numeral := 65 ; Ordinal number listing starts with a
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r 
		{
			; First see if there is text. Then, prefix the text part only (for example, indented text)
			MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar)
			if ( RegExMatchOutputVar.Count() > 0 )
			{	
				OrdinalText := Chr(Numeral) . ". "
				;@ .= NewLine Numeral ". " A_LoopField
				@ .= NewLine
				@ .= RegExReplace(A_LoopField, "([\S].+)", OrdinalText . "$1" )
				Numeral += 1
			}
	    }
	
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		;Clip(@,2)
		SetStatusMessageAndColor("Done.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

TextReplaceCharsNewLine(UserInput){

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
	;vText := "The quick brown fox jumped over the lazy dog.`r`nThen, the dog jumped over the lazy fox."
	vText := UserInput
	
	Prompt = "Enter character(s) to find/replace with new line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

	vUserFind := OutputVar
	;vAltText := RegExReplace(vText, vUserFind, "`r`n")
	vAltText := StrReplace(vText, vUserFind, "`r`n") ; Symbols are free to use in StringReplace
	
	If ErrorLevel
	{
		MsgBox, "(TextReplaceRegex) There was an issue."
		return
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return vAltText
}

TextReplaceRegex(UserInput){

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
	;vText := "The quick brown fox jumped over the lazy dog.`r`nThen, the dog jumped over the lazy fox."
	vText := UserInput
	Prompt = "Enter regular expression for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserRegex := OutputVar

	; replace double quotes with "\x22"

	Prompt = "Enter replacement string for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserReplacement := OutputVar

	if ( !vUserReplacement )
		vUserReplacement := ""

	vAltText := RegExReplace(vText, vUserRegex, vUserReplacement)
	;MsgBox % vText
	;MsgBox % vAltText
	If ErrorLevel
	{
		MsgBox, "(TextReplaceRegex) There was an issue."
		return
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return vAltText
}

TextReplaceRegexEachLine(UserInput){

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
	;vText := "The quick brown fox jumped over the lazy dog.`r`nThen, the dog jumped over the lazy fox."
	vText := UserInput
	Prompt = "Enter regular expression for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserRegex := OutputVar

	; replace double quotes with "\x22"

	Prompt = "Enter replacement string for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserReplacement := OutputVar

	if ( !vUserReplacement )
		vUserReplacement := ""

	NewLine := "`r`n"
	Text := UserInput
	
	@ := ""
	Loop, Parse, Text, `n, `r
	{
		@ .= NewLine RegExReplace(A_LoopField, vUserRegex, vUserReplacement)
		;vAltText := RegExReplace(vText, vUserRegex, vUserReplacement)
		;MsgBox % vText
		;MsgBox % vAltText
		If ErrorLevel
		{
			MsgBox, "(TextReplaceRegex) There was an issue."
			return
		}
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(@, StrLen(NewLine) + 1)
	;return vAltText
}


TextReplaceStringXTimes(UserInput){
	
	; Find/Replace string up to X occurrences found
	vText := UserInput
	
	Prompt = "What do you want to replace?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserFind := OutputVar
	
	Prompt = "Enter replacement string for find/replace"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserReplacement := OutputVar
	
	StrReplace(vText, vUserFind, vUserReplacement, vCount)
	
	Prompt = "Enter a number up to %vCount% (-1 or blank for all occurrences)"
	Dft = -1
	InputBox, OutputVar, How many occurrences to update..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserNumber := OutputVar

	if ( !vUserReplacement )
		vUserReplacement := ""

	vAltText := StrReplace(vText, vUserFind, vUserReplacement,, vUserNumber)
	;MsgBox % vText
	;MsgBox % vAltText
	If ErrorLevel
	{
		MsgBox, "(TextReplaceStringXTimes) There was an issue."
		return
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return vAltText
}

HelperLoopEachLineCallFunction(vText, vCallBack, vCallBackArgs*)
{
	@ :=
	NewLine := "`r`n"
	
	if ( InStr(vText, "`n") )
	{	
		Loop, Parse, vText, `n, `r
		{
			@ .= NewLine %vCallBack%(A_LoopField, vCallBackArgs*)
		}
		return SubStr(@, StrLen(NewLine)+1)
	}
	else
		return %vCallBack%(vText, vCallBackArgs*)
}

CallbackReplaceLastOccurrence(vText, vParams*)
{
	vFind := vParams[1]
	vReplace := vParams[2]
	vFoundPos := InStr(vText, vFind,0,0) ; Find last occurrence ; Return 0 if not found
	
	if ( vFoundPos )
		return RegExReplace(vText, vFind, vReplace, 0, 1, vFoundPos) ; 
	else
		return vText
}

CallbackReplaceFirstOccurrence(vText, vParams*)
{
	vFind := vParams[1]
	vReplace := vParams[2]
	
	/*
	vFoundPos := InStr(vText, vFind,0,0) ; Find last occurrence ; Return 0 if not found
	
	if ( vFoundPos )
		return RegExReplace(vText, vFind, vReplace, 0, 1, vFoundPos) ; 
	else
		return vText
	*/
	
	return RegExReplace(vText, vFind, vReplace,,1)
}

TextReplaceLastStringOccurrenceEachLine(UserInput)
{
	; Replace last occurrence in a string - Using variadic variables and callback functions
	;vText := "RUN RUN RUN GOT YOU!`nRUN RUN RUN GOT YOU!`nRUN RUN RUN GOT YOU!"
	;vFind := "RUN"
	;vReplace := "WALK"
	
	vText := UserInput
	
	Prompt = "What do you want to replace?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vFind := OutputVar
	
	if ( StrLen(vFind) == 0 )
		return
	
	Prompt = "What do you want to replace it with?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vReplace := OutputVar
	
	vResult := HelperLoopEachLineCallFunction(vText, "CallbackReplaceLastOccurrence", vFind, vReplace)
	;MsgBox % vResult
	SetStatusMessageAndColor("Done.", "Green")
	return vResult
}

TextReplaceFirstStringOccurrenceEachLine(UserInput)
{
	; Replace last occurrence in a string - Using variadic variables and callback functions
	
	vText := UserInput
	
	Prompt = "What do you want to replace?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vFind := OutputVar
	
	if ( StrLen(vFind) == 0 )
		return
	
	Prompt = "What do you want to replace it with?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vReplace := OutputVar
	
	vResult := HelperLoopEachLineCallFunction(vText, "CallbackReplaceFirstOccurrence", vFind, vReplace)
	SetStatusMessageAndColor("Done.", "Green")
	return vResult
}

TextReplaceLastStringOccurrence(UserInput)
{
	; Replace last occurrence in a string - Using variadic variables and callback functions
	vText := UserInput
	
	Prompt = "What do you want to replace?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vFind := OutputVar
	
	if ( StrLen(vFind) == 0 )
		return
	
	Prompt = "What do you want to replace it with?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vReplace := OutputVar
	
	vFoundPos := InStr(vText, vFind,0,0) ; Find last occurrence ; Return 0 if not found
	
	if ( vFoundPos )
		return RegExReplace(vText, vFind, vReplace, 0, 1, vFoundPos) ; 
	else
		return vText
}

TextReplaceFirstStringOccurrence(UserInput)
{
	; Replace last occurrence in a string - Using variadic variables and callback functions
	vText := UserInput
	
	Prompt = "What do you want to replace?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vFind := OutputVar
	
	if ( StrLen(vFind) == 0 )
		return
	
	Prompt = "What do you want to replace it with?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vReplace := OutputVar
	
	/*
	vFoundPos := InStr(vText, vFind,0,0) ; Find last occurrence ; Return 0 if not found
	
	if ( vFoundPos )
		return RegExReplace(vText, vFind, vReplace, 0, 1, vFoundPos) ; 
	else
		return vText
	*/
	return RegExReplace(vText, vFind, vReplace,,1)
}


TextTrim(UserInput) {

	NewLine := "`r`n"
	Text := UserInput
	
	@ := ""
	Loop, Parse, Text, `n, `r
	{
		@ .= NewLine Text := Trim(A_LoopField)
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextTrimLeft(UserInput) {

	NewLine := "`r`n"
	Text := UserInput
	
	@ := ""
	Loop, Parse, Text, `n, `r
	{
		@ .= NewLine Text := LTrim(A_LoopField)
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextTrimRight(UserInput) {

	NewLine := "`r`n"
	Text := UserInput
	
	@ := ""
	Loop, Parse, Text, `n, `r
	{
		@ .= NewLine Text := RTrim(A_LoopField)
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextRemoveSpaces(UserInput) {
	NewLine := "`r`n"
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r
		{
			@ .= NewLine Text := RegExReplace(A_LoopField, " ", "")
		}
		
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		SetStatusMessageAndColor("Done.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

TextReplaceSpacesUnderscore(UserInput) {
	NewLine := "`r`n"
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r
		{
			@ .= NewLine Text := RegExReplace(A_LoopField, " ", "_")
		}
		
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		SetStatusMessageAndColor("Done.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

TextReplaceSpacesDash(UserInput) {
	NewLine := "`r`n"
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		@ := ""
		Loop, Parse, Text, `n, `r
		{
			@ .= NewLine Text := RegExReplace(A_LoopField, " ", "-")
		}
		
		;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
		SetStatusMessageAndColor("Done.", "Green")
		return SubStr(@, StrLen(NewLine) + 1)
	;}
}

RemoveBlankLines(UserInput)
{	
	Text := UserInput
	Loop
	{
		;StringReplace, Clipboard, Clipboard, `r`n`r`n, `r`n, UseErrorLevel
		StringReplace, Text, Text, `r`n`r`n, `r`n, UseErrorLevel
		If ErrorLevel = 0  ; No more replacements needed.
			break
	}
	
	SetStatusMessageAndColor("Done.", "Green")
	return Text
}

RemoveDuplicateLines(UserInput)
{
	arrLines := StrSplit(UserInput, ["`r`n", "`n", "`r"])
	trimmedArray := HelperTrimArray(arrLines)
	joinLines := Join(trimmedArray, "`n")
	
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(joinLines, 1, StrLen(joinLines) - StrLen("`n") + 1)
}

; https://stackoverflow.com/questions/46432447/how-do-i-remove-duplicates-from-an-autohotkey-array
HelperTrimArray(arr) { ; Hash O(n)

    hash := {}, newArr := []

    for e, v in arr
        if (!hash.Haskey(v))
            hash[(v)] := 1, newArr.push(v)

	SetStatusMessageAndColor("Done.", "Green")
    return newArr
}

RemoveRepeatingDuplicateLines(UserInput)
{
	Text := UserInput . "`n"
	
	Loop, Parse, Text, `n, `r
	{
		delim := A_LoopField . "`n"
		Text := st_removeDuplicates(Text, delim)
	}
	
	TextMinusDelim := StrLen(Text) - StrLen("`n")
	
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(Text, 1, TextMinusDelim)
}

RemoveRepeatingDuplicateChars(UserInput)
{
	Prompt = "Repeating Character?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if (StrLen(OutputVar) == 0)
		return
	
	Text := UserInput
	
	SetStatusMessageAndColor("Done.", "Green")
	return st_removeDuplicates(Text, OutputVar)
}


/*
String Things - Common String & Array Functions by Tidbit

RemoveDuplicates
   Remove any and all consecutive lines. A "line" can be determined by
   the delimiter parameter. Not necessarily just a `r or `n. But perhaps
   you want a | as your "line".

   string = The text or symbols you want to search for and remove.
   delim  = The string which defines a "line".

example: st_removeDuplicates("aaa|bbb|||ccc||ddd", "|")
output:  aaa|bbb|ccc|ddd
*/
st_removeDuplicates(string, delim="`n")
{
   delim:=RegExReplace(delim, "([\\.*?+\[\{|\()^$])", "\$1")
   
   SetStatusMessageAndColor("Done.", "Green")
   Return RegExReplace(string, "(" delim ")+", "$1")
}


HelperPadSpaces(vLength, vText, vDirection) {
	
	if (vDirection == "Left")
		vFormatStr := "{:" . vLength . "}"
	else if (vDirection == "Right")
		vFormatStr := "{:-" . vLength . "}"
	else
		vFormatStr := "{:" . vLength . "}" ; Default = Left
	
	return Format(vFormatStr, vText)
}

HelperPadSpacesLeftOrRight(vLength, vText, vDirection) {
	
	@ := ""
	NewLine := "`r`n"
	
	Loop, Parse, vText, `n, `r
	{
		@ .= NewLine
		@ .= HelperPadSpaces(vLength, A_LoopField, vDirection)
	}
	
	return SubStr(@, StrLen(NewLine) +1)
}

TextPadSpaceLeft(UserInput)
{
	; Pad to the left of each line with spaces until it reaches a desired length
	;  - If line is greater than length entered, skip
	
	; Input: 1+ Lines
	; Input: User's desired length of padding
	
	Prompt = "Pad up to what length?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
	if OutputVar is not integer
	{ 
		SetStatusMessageAndColor("Invalid value for pad length.", "Red")
		return
	}
	
	vText := HelperPadSpacesLeftOrRight(OutputVar, UserInput, "Left")
	return vText
}

TextPadSpaceRight(UserInput)
{
	; Pad to the right of each line with spaces until it reaches a desired length
	;  - If line is greater than length entered, skip
	
	; Input: 1+ Lines
	; Input: User's desired length of padding
	
	Prompt = "Pad up to what length?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	vText := HelperPadSpacesLeftOrRight(OutputVar, UserInput, "Right")
	return vText
}

HelperFindHighestPositionInTextLines(vFindChar, vText) {
		vMaxChar := 0
		
		Loop, Parse, vText, `n, `r
		{
			; Returns the position of an occurrence of the string Needle in the string Haystack.
			; Position 1 is the first character; this is because 0 is synonymous with "false", making it an intuitive "not found" indicator.
			vFoundPos := InStr(A_LoopField, vFindChar)
			
			; Compare to our placeholder variable for spacing, MaxChar
			; Only overwrite MaxChar if the newly found position is greater than our previously found, and furthest found, position
			If ( vFoundPos <> 0 and vFoundPos > vMaxChar)
				vMaxChar := vFoundPos
		}
		
		return vMaxChar
}

TextPadLeftCharLineUp(UserInput)
{
	
	Prompt = "Character (first occurrence) to line up with?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(OutputVar) == 0 )
	{
		SetStatusMessageAndColor("Character(s) input cannot be nothing.", "Red")
		return
	}
	
	if ( InStr(UserInput, OutputVar) == false )
	{
		SetStatusMessageAndColor("Character(s) should be in string.", "Red")
		return
	}
	
	vText := UserInput
	vFindChar := OutputVar
	vMaxChar := HelperFindHighestPositionInTextLines(vFindChar, vText)
	
	if (vMaxChar == 0 )
	{
		SetStatusMessageAndColor("MaxChar is 0; No need to process or bug.", "Red")
		return
	}
	
	
	NewLine := "`r`n"
	@ := ""
	
	Loop, Parse, vText, `n, `r
	{
		
		vFoundPos := InStr(A_LoopField, vFindChar)
		;MsgBox % "vFoundPos: " vFoundPos ", MaxChar: " vMaxChar
		
		if ( vFoundPos == 0 )
			@ .= NewLine A_LoopField
		else if ( vFoundPos < vMaxChar )
		{
			vPlaceDiff := vMaxChar - vFoundPos
			vNewLength := vPlaceDiff + StrLen(A_Loopfield)
			@ .= NewLine HelperPadSpacesLeftOrRight( vNewLength, A_LoopField, "Left")
		}
		else
			@ .= NewLine A_LoopField
	}
	
	;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
	SetStatusMessageAndColor("Done.", "Green")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextCharLineUp(UserInput) {
	; Maybe allow user to enter a special character or have default (":", "-", ...)
	; Loop to determine the line with the most amount of characters before special character (Item1: Fun = 5 characters before ":")
	; Loop again to precede line with spaces until all lines match (line 1 has 5 characters/spaces before ":" and so done line 2, 3, 4, etc.)
	
	Prompt = "Character to line up with?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
	if (OutputVar <> "")
		FindChar := OutputVar
	else
		FindChar := "a"
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) 
	;{
		
		NewLine := "`r`n"
		MaxChar := 0
		
		; Loop to determine the line with the most amount of characters before special character (Item1: Fun = 5 characters before ":")
		Loop, Parse, Text, `n, `r
		{
			; Returns the position of an occurrence of the string Needle in the string Haystack.
			; Position 1 is the first character; this is because 0 is synonymous with "false", making it an intuitive "not found" indicator.
			FoundPos := InStr(A_LoopField, FindChar)
			
			; Compare to our placeholder variable for spacing, MaxChar
			; Only overwrite MaxChar if the newly found position is greater than our previously found, and furthest found, position
			If ( FoundPos <> 0 and FoundPos > MaxChar)
				MaxChar := FoundPos
		}
		
		If ( MaxChar > 0 )
		{
			; Loop again to precede line with spaces until all lines match (line 1 has 5 characters/spaces before ":" and so done line 2, 3, 4, etc.)
			@ := ""
			Loop, Parse, Text, `n, `r
			{
				; If less than desired length, pad with spaces
				PaddedSpace :=
				if ( InStr(A_LoopField, ":") < MaxChar )
				{
					PlaceDiff := MaxChar - InStr(A_LoopField, FindChar)
					Loop, %PlaceDiff%
					{
						PaddedSpace .= A_Space
					}
					;stdout.WriteLine(PaddedSpace . A_LoopField)
					@ .= NewLine Text := PaddedSpace . A_LoopField
				}
				else
				{
					;stdout.WriteLine(A_LoopField)
					@ .= NewLine Text := A_LoopField
				}
			}
			
			;Clip(SubStr(@, StrLen(NewLine) + 1), 2)
			SetStatusMessageAndColor("Done.", "Green")
			return SubStr(@, StrLen(NewLine) + 1)
		}
	;}
}

TextSay(UserInput)
{
	Text := UserInput
	
	;if ("" <> Text := Clip()) {

		Length := StrLen(Text)
		
		if ( Length > 0 ) {
			
			AnnaVoice := TTS_CreateVoice("Microsoft Anna")
			TTS(AnnaVoice, "SpeakWait", Text)
			SetStatusMessageAndColor("Done.", "Green")
		}
		else {
			;MsgBox,,"Text.ahk",No text available in the clipboard to say
			SetStatusMessageAndColor("No text available in the clipboard to say","Red")
		}
	;}
;return	
}

TextSayStop(UserInput){
	AnnaVoice := TTS_CreateVoice("Microsoft Anna")
	TTS(AnnaVoice,"Stop")
}

TextSurroundingChar(UserInput) {

	Text := UserInput

	;If ("" <> Text := Clip()) {

		Prompt = "Enter left-side surrounding Char"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		Length := StrLen(OutputVar)
		
		if ( Length > 0 ) {
			
			if ( OutputVar == "[" )
				Ending = ] ; Left/Right Square Bracket
			else if ( OutputVar == "<" )
				Ending = > ; Left/Right Angle Bracket
			else if ( OutputVar == "{" )
				Ending = } ; Left/Right Curly Bracket
			else if ( OutputVar == "(" )
				Ending = ) ; Left/Right Parenthesis
			else if ( OutputVar == "`" )
				Ending = ' ; Backtick, Apostrophe
			else if ( OutputVar == "/*" )
				Ending = */ ; Multi-line comment open/close (AHK, CSS, JavaScript)
			else
				Ending = %OutputVar% ; Default - same char on both sides
			
			;Send %OutputVar%%Text%%Ending%
			;Clip(OutputVar . Text . Ending)
			SetStatusMessageAndColor("Done.", "Green")
			return OutputVar . Text . Ending
		}
	;}
}

TextPrefixSuffixRemoveEachLine(UserInput) {
	; Wrap the text with a open/close HTML tag

	Text := UserInput

	;If ("" <> Text := Clip())
	;{
	;Text := Clipboard

		Prompt = "Enter prefix text to remove from each line"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		Prefix := OutputVar
		
		Prompt = "Enter suffix text to remove from each line"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		Suffix := OutputVar

		if ( (StrLen(Prefix) == 0) and (StrLen(Suffix) == 0) )
		{
			;MsgBox % "No data has been entered for Prefix or Suffix. Aborting."
			SetStatusMessageAndColor("No data entered for Prefix or Suffix.","Red")
			return
		}
		
		doPrefix := false
		doSuffix := false
		
		if ( StrLen(Prefix) > 0 )
			doPrefix := true
		
		if ( StrLen(Suffix) > 0 )
			doSuffix := true
			
		NewLine := "`r`n"
		@ := 
		
		Loop, Parse, Text, `n, `r
		{
			; Only surround text, not any indenting non-whitespace characters
			@ .= NewLine
			@ .= RegExReplace(A_LoopField, "(" . Prefix . "([\S].+)" . Suffix . ")", "$2" )
		}
		
		;MsgBox % @
		;Clip(@)
		retValue := SubStr(@, StrLen(NewLine) +1) ; Remove last new line (`r`n) or blank line
		SetStatusMessageAndColor("Done.", "Green")
		return retValue
	;}
}

TextPrefixSuffixEachLine(UserInput){
	; Wrap the text with a open/close HTML tag

	Text := UserInput
	
	;If ("" <> Text := Clip()) {
	;Text := Clipboard

		Prompt = "Enter text to prefix each line"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		Prefix := OutputVar
		
		Prompt = "Enter text to suffix each line"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		Suffix := OutputVar

		if ( (StrLen(Prefix) == 0) and (StrLen(Suffix) == 0) )
		{
			;MsgBox % "No data has been entered for Prefix or Suffix. Aborting."
			SetStatusMessageAndColor("No data entered for Prefix or Suffix.","Red")
			return
		}
		
		if ( (StrLen(Prefix) > 0) or (StrLen(Suffix) > 0) ) {
			
			NewLine := "`r`n"
			@ := 
			
			Loop, Parse, Text, `n, `r
			{
				; Only surround text, not any indenting non-whitespace characters
				@ .= NewLine
				@ .= RegExReplace(A_LoopField, "([^\s][\S].+)", (Prefix . "$1" . Suffix) )
			}
			
			;Clip(@)
			retValue := SubStr(@, StrLen(NewLine) + 1) ; Remove last new line (`r`n) or blank line
			SetStatusMessageAndColor("Done.", "Green")
			return retValue
		}
	;}
}

TextSurroundingHTMLTag(UserInput){
	; Wrap the text with a open/close HTML tag

	Text := UserInput
	
	;If ("" <> Text := Clip()) {

		Prompt = "Enter html tag to wrap"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		Length := StrLen(OutputVar)
		
		if ( Length > 0 ) {
			
			;Send <%OutputVar%>%Text%</%OutputVar%>
			;Clip("<" . OutputVar . ">" . Text . "</" . OutputVar . ">")
			retValue := "<" . OutputVar . ">" . Text . "</" . OutputVar . ">"
			
			SetStatusMessageAndColor("Done.", "Green")
			return retValue
		}
	;}
}

TextSurroundingHTMLTagEachLine(UserInput){
	; Wrap the text with a open/close HTML tag

	Text := UserInput
	
	;If ("" <> Text := Clip()) {
	;Text := Clipboard

		Prompt = "Enter html tag to wrap each line"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		Length := StrLen(OutputVar)
		
		if ( Length > 0 ) {
			
			NewLine := "`r`n"
			@ := 
			
			Loop, Parse, Text, `n, `r
			{
				; Only surround text, not any indenting non-whitespace characters
				@ .= NewLine
				@ .= RegExReplace(A_LoopField, "([^\s][\S].+)", "<" . OutputVar . ">$1</" . OutputVar . ">")
			}
			
			;Clip(@)
			;return @
			retValue := SubStr(@, StrLen(NewLine) +1) ; Remove last new line (`r`n) or blank line
			
			SetStatusMessageAndColor("Done.", "Green")
			return retValue
		}
	;}
}

TextSortAlphabetically(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		Sort, Text, ; Normal sort
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

TextSortAlphabeticallyUnique(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		Sort, Text, C U ; Normal sort
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

TextSortAlphabeticallyUniqueCaseInsensitive(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		Sort, Text, U ; Normal sort
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

TextSortAlphabeticallyWithDelimiter(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	
	;If ("" <> Text := Clip()) {
		
		Prompt = "Enter delimiter separating values"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		Length := StrLen(OutputVar)
		
		if ( Length > 0 ) {
			Sort, Text, D%OutputVar% ; Alphabetically sort for words/text separated by a delimiter
			;Clip(Text)
			
			SetStatusMessageAndColor("Done.", "Green")
			return Text
		}
	;}
}

TextSortNumericAscending(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		Sort, Text, N ; Numeric Sort
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

TextSortNumericDescending(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		Sort, Text, R N ; Reverse Numeric Sort
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

TextSortNumericWithDelimiter(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		
		Prompt = "Enter delimiter separating values"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		Length := StrLen(OutputVar)
		
		if ( Length > 0 ) {
			Sort, Text, N D%OutputVar% ; Numeric Sort
			;Clip(Text)
			
			SetStatusMessageAndColor("Done.", "Green")
			return Text
		}
	;}
}

TextSortReverse(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		Sort, Text, R ; Sort and then Reverse
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

TextSortReverseOrder(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		Sort, Text, F HelperReverseDirection ; Reverse order
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

HelperReverseDirection(a1, a2, offset)
{
    return offset  ; Offset is positive if a2 came after a1 in the original list; negative otherwise.
}

TextSortReverseWithDelimiter(UserInput) {
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		
		Prompt = "Enter delimiter separating values"
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		Length := StrLen(OutputVar)
		
		if ( Length > 0 ) {
			Sort, Text, R D%OutputVar% ; Sort and then Reverse
			;Clip(Text)
			
			SetStatusMessageAndColor("Done.", "Green")
			return Text
		}
	;}
}

TextSortRandom(UserInput) 
{
	; Sort selected text and replace it
	
	Text := UserInput
	;If ("" <> Text := Clip()) {
		Sort, Text, Random ; Random sort text
		;Clip(Text)
		
		SetStatusMessageAndColor("Done.", "Green")
		return Text
	;}
}

UrlDownloadToTextFileAndOpen(UserInput){
	
	;vUrl := "https://www.google.com"
	vUrl := UserInput
	vFilename := "C:\Users\" . A_UserName . "\Desktop\Scrap for Copying.txt"
	
	isHttp := SubStr(vUrl,1,4)
	StringLower, isHttp, isHttp
	
	if ( isHttp != "http")
	{
		MsgBox % "Url is invalid. It must start with http."
		return
	}
	
	try
	{
		URLDownloadToFile, %vUrl%, vFilename
		FileRead, vFileReadOutput, vFilename
		;MsgBox % vFileReadOutput
		
		SetStatusMessageAndColor("Done.", "Green")
		return vFileReadOutput
	}
	catch e
	{
		;MsgBox % "There was an error with URLDownloadToFile"
		MsgBox, 16,, % "Exception thrown!`n`nwhat: " e.what "`nfile: " e.file . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra
		
		return
	}
}

WindowsRun(UserInput){
	Commands := StrSplit(UserInput, ["`r`n", "`n", "`r"])
	;MsgBox % Join(Commands, ", ")
	;MsgBox % Commands[1]
	Command := Commands[1]
	Run, %Command%
	SetStatusMessageAndColor("Done.", "Green")
}