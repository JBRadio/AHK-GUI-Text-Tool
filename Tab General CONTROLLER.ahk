/*SetStatusMessageAndColor(messageText, messageColor)
{
	Gui, Font, c%messageColor%
	GuiControl, Font, StatusMessage
	GuiControl,, StatusMessage, %messageText%
}
*/

ProcessFunction(OutputVarFunctionName, OutputVarInputType){ ; Main processing function for active tab
	;MsgBox % "Function name: " OutputVarFunctionName
	;MsgBox % "Input Type: " OutputVarInputType
	
	GuiControlGet, OutputVarInput,,EditInput,Edit
	
	if ( StrLen(OutputVarInput) = 0  and (OutputVarInputType != "N/A") )
	{
		SB_SetText("User input field is blank. (ProcessFunction)")
		return
	}

	SB_SetText("")
	ReturnFunctionValue := %OutputVarFunctionName%(OutputVarInput)
	GuiControl,, EditOutput, %ReturnFunctionValue%
}

;~ ButtonOKGeneral:
;~ ;ButtonOK:
	;~ Gui, ListView, MyListView
	;~ ;	"To detect when the user has pressed Enter while a ListView has focus, use a default button (which can be hidden if desired)."
	;~ GuiControlGet, OutputVarFocusedVariable, FocusV	; make sure the "v" flag is set on control element. "g" caused this not to work.
	;~ MsgBox % "OutputVarFocusedVariable: " OutputVarFocusedVariable
	;~ if (OutputVarFocusedVariable != "MyListView")
		;~ return

	;~ ;;MsgBox % "Enter was pressed. The focused row number is " . LV_GetNext(0, "Focused")
	
	;~ ; Function|Input|Description|Function Name
	;~ selRow := LV_GetNext(0, "Focused")
	;~ LV_GetText(OutputVarFunctionName, selRow, 4)
	;~ LV_GetText(OutputVarInput, selRow, 2)
	;~ ProcessFunction(OutputVarFunctionName, OutputVarInput)
;~ return

ButtonClear:
	GuiControl,, EditInput,
	GuiControl,, EditOutput,
	;SetStatusMessageAndColor("Cleared all inputs.", "Green")
	SB_SetText("Cleared all inputs.")
return

ButtonCopy:
	GuiControlGet, OutputVarOutput,,EditOutput,Edit
	OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")	; The "Edit" controls only have `n so need to add `r`n on "Copy"
	Clipboard := OutputVarStrReplace
	;SetStatusMessageAndColor("Copied output to Windows Clipboard.", "Green")
	SB_SetText("Copied output to Windows Clipboard.")
return

ButtonMoveUp:
	GuiControlGet, OutputVarOutput,,EditOutput,Edit
	OutputVarStrReplace := StrReplace(OutputVarOutput, "`n", "`r`n")
	GuiControl,, EditInput, %OutputVarStrReplace%
	;SetStatusMessageAndColor("Copied output to input.", "Green")
	SB_SetText("Copied output to input.")
return

ButtonPaste:
	GuiControl,, EditInput, %Clipboard%
	;SetStatusMessageAndColor("Pasted Windows Clipboard content to User Input.", "Green")
	SB_SetText("Pasted Windows Clipboard content to User Input.")
return

ButtonProcess:
	; Overwriting function to reduce duplicate buttons
	GuiControlGet, Current_Tab,,TabVar
	if ( Current_Tab == "General" )
		Gui, ListView, MyListView
	else if ( Current_Tab == "Programmer" )
		Gui, ListView, MyListViewProg
	else
		MsgBox % "Current tab is: " Current_Tab
	
	RowNumber := LV_GetNext(0,"F")
	if (!RowNumber)
	{
		SB_SetText("No selected row found.")
		return
	}
	
	; Function|Input|Description|Function Name
	LV_GetText(OutputVarInput, RowNumber, 2)
	LV_GetText(OutputVarFunctionName, RowNumber, 4)
	ProcessFunction(OutputVarFunctionName, OutputVarInput)

	;~ORIGINAL CODE
	;~ Gui, ListView, MyListView
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
	;~ ProcessFunction(OutputVarFunctionName, OutputVarInput)
return

MyListView:
	GuiControlGet, OutputVarFocusedVariable, FocusV	; make sure the "v" flag is set on control element. "g" caused this not to work.
	
	if (OutputVarFocusedVariable != "MyListView")
	{
		SB_SetText("Wrong elment has focus (MyListView)")
		return
	}
	
	Gui, ListView, MyListView
	
	;MsgBox % "A_GuiEvent: " A_GuiEvent
	if (A_GuiEvent = "DoubleClick")
	{
		;ToolTip You double-clicked row number %A_EventInfo%. Text: %RowText%
		LV_GetText(OutputVarFunction, A_EventInfo, 1)
		LV_GetText(OutputVarInput, A_EventInfo, 2)
		LV_GetText(OutputVarDescription, A_EventInfo, 3)
		LV_GetText(OutputVarFunctionName, A_EventInfo, 4)
		;MsgBox % "OutputVarFunction: " OutputVarFunction "`nOutputVarInput: " OutputVarInput "`nOutputVarDescription: " OutputVarDescription "`nOutputVarFunctionName: " OutputVarFunctionName
		
		/*
		GuiControlGet, OutputVarInput,,EditInput,Edit
		if ( StrLen(OutputVarInput) = 0 )
		{
			;MsgBox % "User input is blank"
			SetStatusMessageAndColor("User input field is blank. (MyListView)", "Red")
			return
		}
		*/
		
		;MsgBox % "MyListView - Function name: " OutputVarFunctionName
		ProcessFunction(OutputVarFunctionName, OutputVarInput)
	}
return

ListViewInit:
	Gui, ListView, MyListView
	; Type, Function, Description
	; 1+ Lines, Convert (Uppercase), Converts text to uppercase, for one or more lines of text
	;LV_Add("", "Cmd.exe (Get IPConfig /all)", "N/A", "Runs IPConfig /all and copies it back", "CommandPromptGetIpconfigAll") - Moved to Programmer
	;LV_Add("", "Cmd.exe (Get NSLookup)", "1 Line", "Runs NSLookup for server/URL and copies it back", "CommandPromptGetNsLookup") - Moved to Programmer
	;LV_Add("", "Cmd.exe (Ping)", "1 Line", "Opens Cmd.exe and runs ping command for server name/URL/IP Address", "CommandPromptPing") - Moved to Programmer
	;LV_Add("", "Cmd.exe (SystemInfo)", "N/A", "Opens Cmd.exe and runs systeminfo command", "CommandPromptSystemInfo") - Moved to Programmer
	;LV_Add("", "Cmd.exe (Tracert)", "1 Line", "Opens Cmd.exe and runs tracert command for input", "CommandPromptTracert") - Moved to Programmer
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
	;LV_Add("", "Char (Ascii)", "Char", "Encodes the leftmost character to ASCII value", "EncodeToAscii")
	LV_Add("", "Date MM/DD/YYYY (Get Full Date)", "1 Line", "Converts MM/DD/YYYY to full date", "TextGetDateFromMMDDYYYY")
	;LV_Add("", "JSON (Format)", "1 Line", "Formats JSON notation to pretty print", "TextFormatJson") - Moved to Programmer
	;LV_Add("", "Number (Char Code)", "Char", "Encodes a number to Char Code value", "EncodeToCharCode") - Moved to Programmer
	LV_Add("", "Number (Telephone)", "1+ Lines", "Formats matched 10 numbers to a telephone format", "TextFormatTelephoneNumber")
	;LV_Add("", "HTML (Encode)", "1+ Lines", "Encodes text to HTML entities (& to &amp;)", "EncodeToHtml") - Moved to Programmer
	;LV_Add("", "URL (Encode)", "1+ Lines", "Encodes text to URI/URL values", "EncodeToUrl") - Moved to Programmer
	;LV_Add("", "URL (Decode)", "1+ Lines", "Decodes URI text", "DecodeUrl") - Moved to Programmer
	;LV_Add("", "URL (Decode, Get Params)", "1 Line", "Decodes URI text and lists URI params", "DecodeUrlAndGetParams") - Moved to Programmer
	;LV_Add("", "URL (Download to Text File)", "1 Line", "Downloads internet file to a text file and opens it.", "UrlDownloadToTextFileAndOpen") - Moved to Programmer
	;
	LV_Add("", "Duplicate X Times", "1 Line", "Duplicates text X times with new lines in between", "TextDuplicateXTimes")
	LV_Add("", "Duplicate Increment Number X Times", "1 Line", "Duplicates text X times with new lines, incrementing the last number found each time by 1", "TextDuplicateIncrementLastNumberXTimes")
	;
	LV_Add("", "CSV (Extract 1 Column)", "CSV", "Pick a column to extract CSV row data from", "ExtractColumnDataForCSV")
	LV_Add("", "CSV (Filter 1 Column)", "CSV", "Pick a column to remove CSV row data from", "FilterColumnDataForCSV")
	;LV_Add("", "CSV (To <table>)", "CSV", "Convert CSV to html table code", "CsvToHtmlTable") - Moved to Programmer
	LV_Add("", "CSV (Sort by Column #)", "CSV", "Pick a column to sort CSV row data from", "SortDataByColumnCSV")
	LV_Add("", "CSV (Swap 1 Column)", "CSV", "Swap the position of one column's data with another column's position", "SwapColumnForCSV")
	LV_Add("", "TSV (Extract 1 Column)", "TSV", "Pick a column to extract TSV row data from", "ExtractColumnDataForTSV")
	LV_Add("", "TSV (Filter 1 Column)", "TSV", "Pick a column to remove TSV row data from", "FilterColumnDataForTSV")
	LV_Add("", "TSV (Swap 1 Column)", "TSV", "Swap the position of one column's data with another column's position", "SwapColumnForTSV")
	LV_Add("", "TSV (Sort by Column #)", "TSV", "Pick a column to sort TSV row data from", "SortDataByColumnTSV")
	LV_Add("", "*SV (Extract 1 Column)", "*SV", "Pick a column to extract User-specific delimiter-SV row data from", "ExtractColumnDataForAnySV")
	LV_Add("", "*SV (Filter 1 Column)", "*SV", "Pick a column to remove User-specific delimiter-SV row data from", "FilterColumnDataForAnySV")
	LV_Add("", "*SV (Swap 1 Column)", "*SV", "Swap the position of one column's data with another column's position", "SwapColumnForAnySV")
	LV_Add("", "*SV (Spaced Columns)", "*SV", "Print delimiter-SV data in spaced columns based on column data length", "SpacedColumnForAnySV")
	LV_Add("", "*SV (Sort by Column #)", "TSV", "Pick a column to sort TSV row data from", "SortDataByColumnForAnySV")
	;
	LV_Add("", "Flatten to 1 Line", "1+ Lines", "Condenses to one line - removing duplicate white space and new lines", "TextFlattenSingleLine")
	LV_Add("", "Flatten each line", "1+ Lines", "Trims and removes duplicate whitespace for each line", "TextFlattenMultiLine")
	;
	LV_Add("", "Indent More", "1+ Lines", "Tabs one or more lines of text", "TextFormatIndentMore")
	LV_Add("", "Indent Less", "1+ Lines", "Decreases tabs by one for one or more lines of text", "TextFormatIndentLess")
	;
	LV_Add("", "Join Lines w/ w/o Char", "1+ Lines", "Joins lines of text to a single line by user entered separator", "TextJoinLinesWithChar")
	LV_Add("", "Join SV Line w/ Char", "1+ Lines", "Split and join a line of text by user entered separator", "TextJoinSplittedLineWithChar")
	LV_Add("", "Join SV Lines w/ Char", "1+ Lines", "Split and join lines of text by user entered separator", "TextJoinSplittedLinesWithChar")
	;
	LV_Add("", "Open Link (AHK RegEx Quick Reference)", "N/A", "Opens AHK RegEx Quick Reference page in default browser", "OpenLinkAHKRegexQuickReference")
	LV_Add("", "Open Link (RegExr.com - RegEx Tester)", "N/A", "Opens AHK RegEx Quick Reference page in default browser", "OpenLinkRegExr")
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
	LV_Add("", "Reduce (Lines that include text)", "1+ Lines", "Reduces the input based on whether or not it includes text", "RemoveLinesThatIncludeText")
	LV_Add("", "Reduce (Lines that exclude text)", "1+ Lines", "Reduces the input based on whether or not it excludes text", "RemoveLinesThatExcludeText")
	LV_Add("", "Reduce (Lines that RegEx match)", "1+ Lines", "Reduces the input based on whether or not it excludes text", "RemoveLinesThatMatchRegEx")
	LV_Add("", "Reduce (Lines that RegEx don't match)", "1+ Lines", "Reduces the input based on whether or not it excludes text", "RemoveLinesThatDontMatchRegEx")
	;
	LV_Add("", "List (Ordinal Number)", "1+ Lines", "Prefixes lines of text with ordinal numbers. (Also see Prefix/Suffix)", "TextPrefixOrdinalNumber")
	LV_Add("", "List (Ordinal Lowercase)", "1+ Lines", "Prefixes lines of text with ordinal lowercase letters. (Also see Prefix/Suffix)", "TextPrefixOrdinalLowercase")
	LV_Add("", "List (Ordinal Uppercase)", "1+ Lines", "Prefixes lines of text with ordinal uppercase letters. (Also see Prefix/Suffix)", "TextPrefixOrdinalUppercase")
	LV_Add("", "List (Unordered Custom Character)", "1+ Lines", "Prefixes lines of text with a custom character. (Also see Prefix/Suffix)", "TextPrefixCustomRepeatingCharacter")
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
	LV_Add("", "Prefix/Suffix Remove (Text or RegEx)", "1+ Lines", "Remove Prefix/Suffix of text from each line", "TextPrefixSuffixRemoveEachLine")
	LV_Add("", "Prefix/Suffix Lines", "1+ Lines", "Adds a custom Prefix/Suffix of text to each line", "TextPrefixSuffixEachLine")
	;
	;LV_Add("", "HTML (Add Tag)", "1+ Lines", "Surround text with user defined HTML tag", "TextSurroundingHTMLTag") - Moved to Programmer
	;LV_Add("", "HTML (Add Tag Each Line)", "1+ Lines", "Surround text on each line with defined HTML tag", "TextSurroundingHTMLTagEachLine") - Moved to Programmer
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
	;LV_Add("", "Windows (Run)", "1 Line", "Use the Windows OS Run command on input.", "WindowsRun") - Moved to Programmer
	;LV_Add("", "Function", "1+ Lines", "Desc", "FunctionName")
return