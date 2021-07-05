TextAnalysisCount(UserInput) {
	; Mango.ahk
	
	vText := UserInput
	
	len := StrLen(vText)
	Spaces:=""
	newline:=""
	chars:=""
	Tabs:=""
	
	Loop, Parse, vText, `n, `r
		newline++
	
	Loop, Parse, vText
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
	
	SB_SetText("Analysis complete.")
	return "Length: " . len . "`nSpaces: " . spaces . "`nTabs: " . Tabs . "`nCharacters: " . chars . "`nLines: " . newline
	
}

ClipboardAppend(UserInput)
{
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly appended data
	
	SB_SetText("Clipboard + Appeneded data in Output.")
	return Clipboard . UserInput
}

ClipboardPrepend(UserInput)
{
	; TESTED?
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly prepended data
	
	SB_SetText("Prepended data + Clipboard to Output.")
	return UserInput . Clipboard
}

ClipboardAppendAsNewLine(UserInput)
{
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly appended data
	
	SB_SetText("Clipboard + Appeneded data + New Line in Output.")
	return Clipboard . "`n" . UserInput
}

ClipboardPrependAsNewLine(UserInput)
{
	; Append user input text to clipboard (after/append)
	; Does not add a new line before or after appeneded data
	
	; Additional Considerations: 
	;  - Ask user to supply a character to fit between existing clipboard data and newly prepended data
	
	SB_SetText("Prepended data + New Line + Clipboard to Output.")
	return UserInput . "`n" . Clipboard
}

ConvertUppercase(UserInput)
{
	; Converts all text to uppercase text
	
	StringUpper, OutputVarStringUpper, UserInput
	
	SB_SetText("Converted input to Uppercase characters.")
	return OutputVarStringUpper
}

ConvertLowercase(UserInput)
{
	; Converts all text to lowercase text
	
	StringLower, OutputVarStringLower, UserInput
	
	SB_SetText("Converted input to Lowercase characters.")
	return OutputVarStringLower
}

ConvertTitleCase(UserInput)
{
	; Converts all text/sentences to title case
	
	StringUpper, OutputVarStringUpper, UserInput, T	; "T" for Title Case
	
	SB_SetText("Converted input to Title case.")
	return OutputVarStringUpper
}

ConvertSentenceCase(UserInput)
{
	; SENTENCE CASE - quick brown fox. went home. = Quick brown fox. Went home.
	
	vText := UserInput
	vText := RegExReplace(vText, "([.?\s!]\s\w)|^(\.\s\b\w)|^(.)", "$U0")
	
	SB_SetText("Converted input to Sentence case.")
	return vText
}

ConvertInvertCase(UserInput) {
	; Thanks JDN ; Mango.ahk
	
	vText := UserInput
	Lab_Invert_Char_Out:= ""

	Loop % Strlen(vText) {
		;Lab_Invert_Char:= Substr(Clipboard, A_Index, 1)
		Lab_Invert_Char:= Substr(vText, A_Index, 1)
		If Lab_Invert_Char is Upper
			Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) + 32)
		Else If Lab_Invert_Char is Lower
			Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) - 32)
		Else
			Lab_Invert_Char_Out:= Lab_Invert_Char_Out Lab_Invert_Char
	}
	
	vText:=Lab_Invert_Char_Out
	
	SB_SetText("Converted input to inverted case.")
	return vText
}

ConvertCarriageReturnLineFeed(UserInput)
{
	; Turn `r to `r`n, in order to allow to work with Notepad and other Windows programs
	
	vText := UserInput
	vText := RegExReplace(vText, "`r", "`r`n")
	
	SB_SetText("Replaced carriage returns (\r) with carriage return line feeds (\r\n or CRLF)")
	return vText
}

ConvertStrikethrough(UserInput)
{
	; Strikeout with Unicode('\u0336') - borrowed from Text Tools (https://chrome.google.com/webstore/detail/text-tools/mpcpnbklkemjinipimjcbgjijefholkd)
	vText := UserInput
	vSplit := StrSplit(vText, "")
	
	for k,v in vSplit
		vSplit[k] := v . "̶"
	vText := Join(vSplit, "")

	SB_SetText("Converted text to have a strikethrough property")
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

	SB_SetText("Converted text to have a underline property")
	return vText
}

HelperAskUserForDelim()
{
	Prompt := "What character separates the data?"
	Height := HelperDynamicallyFindHeightForInputBox(Prompt) ; (vPromptText)
	ErrorLevel =
	
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (strlen(OutputVar) < 1 and not ErrorLevel)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	if ( ErrorLevel )
		return ; User canceled
	else
		return OutputVar
}

ExtractColumnDataForCSV(UserInput)
{
	vText := UserInput
	vDelim := ","
	vText := HelperExtractColumnDataForAnySV(vText, vDelim)

	SB_SetText("Extracted 1 Column's data (CSV).")
	return vText
}

ExtractColumnDataForTSV(UserInput)
{
	vText := UserInput
	vDelim := "`t"
	vText := HelperExtractColumnDataForAnySV(vText, vDelim)
	
	SB_SetText("Extracted 1 Column's data (TSV).")
	return vText
}

ExtractColumnDataForAnySV(UserInput)
{
	
	vDelim := HelperAskUserForDelim()
	
	if ( StrLen(vDelim) == 0 )
	{
		SB_SetText("Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	vText := UserInput
	vText := HelperExtractColumnDataForAnySV(vText, vDelim)
	
	SB_SetText("Extracted 1 Column's data (*SV).")
	return vText
}

HelperGetSVDataIntoRowsArray(data, delim)
{
	arrRows := Object()
	
	if ( delim == "\t" or delim == "`t" )
		delim := "`t"
	
	; Below loop/parse works for all tested scenarios: `r`n, `n, and a single line of text
	Loop, Parse, data, `n, `r
	{
		arrSplitted := StrSplit(A_LoopField, delim)
		
		Loop % arrSplitted.MaxIndex()
		{
			arrSplitted[A_Index] := Trim(arrSplitted[A_Index])
		}
	
		if ( arrSplitted.Length() == 0 or arrSplitted == )
			continue ; Even a non-delimiter value returns something so return if somehow we have nothing.
		
		arrRows[A_Index] := arrSplitted
	}
	
	return arrRows
}

HelperExtractColumnDataForAnySV(data, delim)
{	
	if (!delim)
		delim := HelperAskUserForDelim()
	
	if ( StrLen(delim) == 0 )
	{
		SB_SetText("Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	;arrRows := HelperGetSVDataIntoRowsArray(data, delim)
	myDSVTool.loadText(data, delim)
	arrRows := myDSVTool.getData()

	; Ask user to pick a column to extract
	intColumns := arrRows[1].Length()

	if ( intColumns < 1 )
		return

	;SampleRow := Join(arrRows[1], ",")
	;Prompt := "Extract Column: Pick a number between 1 and " . intColumns . "`r`n`r`nHere is a Sample row: `r`n" . SampleRow
	vSampleRow := HelperBuildVerticalSampleRow(arrRows[1]) ; (vRowFirst)
	Prompt := "Filter Column: " Prompt "`n`nHere is a Sample row: `n" vSampleRow
	Height := HelperDynamicallyFindHeightForInputBox(Prompt) ; (vPromptText)

	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar = " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	myDSVTool.extractColumnData(OutputVar)
	
	return myDSVTool.getResult()
}

SwapColumnForCSV(UserInput)
{
	; Allow the user to swap the position of a columns data with another one.
	; 1.) Ask the user which column numbers to swap
	; 2.) Rebuild the CSV data based on the user's input
	
	vText := UserInput
	vDelim := ","
	vText := HelperSwapColumnDataForAnySV(vText, vDelim)
	
	SB_SetText("Swap 1 Column's data (CSV).")
	return vText
}

SwapColumnForTSV(UserInput)
{
	vText := UserInput
	vDelim := "`t"
	vText := HelperSwapColumnDataForAnySV(vText, vDelim)
	
	SB_SetText("Swap 1 Column's data (TSV).")
	return vText
}

SwapColumnForAnySV(UserInput)
{
	
	vDelim := HelperAskUserForDelim()
	
	if ( StrLen(vDelim) == 0 )
	{
		SB_SetText("Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	if ( vDelim == "\t" or vDelim == "`t" )
		vDelim := "`t"
	
	vText := UserInput
	vText := HelperSwapColumnDataForAnySV(vText, vDelim)
		
	SB_SetText("Swap 1 Column's data (*SV).")
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
	vText := UserInput
	delim := ","
	vText := HelperFilterColumnDataForAnySV(vText, delim)
	
	SB_SetText("Filtered 1 Column's data (CSV).")
	return vText
}

FilterColumnDataForTSV(UserInput)
{
	vText := UserInput
	delim := "`t"
	vText := HelperFilterColumnDataForAnySV(vText, delim)
	
	SB_SetText("Filtered 1 Column's data (TSV).")
	return vText
}

FilterColumnDataForAnySV(UserInput)
{	
	vDelim := HelperAskUserForDelim()
	
	if ( StrLen(vDelim) == 0 )
	{
		SB_SetText("Error: Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	if ( vDelim == "\t" or vDelim == "`t" )
		vDelim := "`t" ; Making sure tab is set. Also convenient for user to have two methods. \t is also used by Regex
	
	vText := UserInput
	vText := HelperFilterColumnDataForAnySV(vText, vDelim)
	
	SB_SetText("Success: Filtered 1 Column's data (*SV).")
	return vText
}

HelperDynamicallyFindHeightForInputBox(vPromptText)
{
	; Count how many lines, then height = 125 + (17 * Lines)
	StringReplace, OutputVar, vPromptText, `n,, UseErrorLevel
	Height := 125 + (17 * (ErrorLevel))
	
	return Height
}

HelperBuildPickAColumn(vRowFirst)
{
	intColumns := vRowFirst.Length()

	if ( intColumns < 1 )
		return 0
	
	MessagePrompt := "Pick a number between 1 and " intColumns "."
	return MessagePrompt
	
}

HelperBuildVerticalSampleRow(vRowFirst)
{
	vSampleRow := 
	Loop % vRowFirst.Length()
		vSampleRow .= "`n" A_Index ": " vRowFirst[A_Index]
	vSampleRow := SubStr(vSampleRow, StrLen("`n") + 1)
	
	return vSampleRow
}

HelperFilterColumnDataForAnySV(data, delim)
{
	arrRows := HelperGetSVDataIntoRowsArray(data, delim)

	; Ask user to pick a column to filter out
	intColumns := arrRows[1].Length()

	if ( intColumns < 1 )
		return
	
	Prompt := HelperBuildPickAColumn(arrRows[1]) ; (vRowFirst)
	
	if ( Prompt == 0 ) {
		SB_SetText("Error: No columns found.")
		return
	}
	
	vSampleRow := HelperBuildVerticalSampleRow(arrRows[1]) ; (vRowFirst)
	Prompt := "Filter Column: " Prompt "`n`nHere is a Sample row: `n" vSampleRow
	Height := HelperDynamicallyFindHeightForInputBox(Prompt) ; (vPromptText)
	
	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar = " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	return HelperFilterColumnDataForColumn(data, delim, OutputVar)
}

HelperFilterColumnDataForColumn(data, delim, columnFilter)
{
	arrRows := HelperGetSVDataIntoRowsArray(data, delim)

	NewLine := "`r`n"
	retValue := 
	
	for indexRow, elementRow in arrRows
	{
		retValue .= NewLine 
		
		for indexColumn, elementColumn in arrRows[indexRow]
		{
			if (indexColumn == columnFilter)
				continue
			else
				retValue .= elementRow[indexColumn] . delim
		}
		
		retValue := Substr(retValue, 1, StrLen(retValue) - StrLen(delim)) ; Remove last delimiter added
	}
	
	retValue := SubStr(retValue, StrLen(NewLine) + 1) ; Remove initial NewLine
	return retValue
}

SortDataByColumnCSV(UserInput)
{
	vText := UserInput
	delim := ","
	vText := HelperSortDataByColumnForAnySV(vText, delim)
	
	SB_SetText("Sorted Column's data (CSV).")
	return vText
}

SortDataByColumnTSV(UserInput)
{
	vText := UserInput
	delim := "`t"
	vText := HelperSortDataByColumnForAnySV(vText, delim)
	
	SB_SetText("Sorted Column's data (TSV).")
	return vText
}

SortDataByColumnForAnySV(UserInput)
{	
	vDelim := HelperAskUserForDelim()
	
	if ( StrLen(vDelim) == 0 )
	{
		SB_SetText("Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	if ( vDelim == "\t" or vDelim == "`t" )
		vDelim := "`t"
	
	vText := UserInput
	vText := HelperSortDataByColumnForAnySV(vText, vDelim)
	
	SB_SetText("Sorted Column data (*SV).")
	return vText
}

; Sort CSV by Column Number chosen
HelperSortDataByColumnForAnySV(data, delim)
{
	vText := data
	vDelim := delim
	arrData := HelperGetSVDataIntoRowsArray(vText, vDelim)
	
	; Ask user to pick a column to filter out
	intColumns := arrData[1].Length()

	if ( intColumns < 1 )
		return

	SampleRow := Join(arrData[1], vDelim)
	Prompt := "Sort Column: Pick a number between 1 and " . intColumns . "`r`n`r`nHere is a Sample row: `r`n" . SampleRow

	ErrorLevel =
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	while (ErrorLevel == 1 or strlen(OutputVar) < 1 or OutputVar = " " or OutputVar > intColumns or OutputVar < 1)
	{
		InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	}
	
	vColumnSort := OutputVar
	
	hasHeader := false
	MsgBox, 4, Data Table, First row header row? (Yes or No)
	IfMsgBox Yes
		hasHeader := true
	
	if ( hasHeader )
		LoopCounter := arrData.MaxIndex()
	else
		LoopCounter := arrData.MaxIndex()-1

	Loop % LoopCounter
	{
		Outer := A_Index
		
		if ( hasHeader )
		{
			Outer += 1
			hasHeader := false
		}
		
		InnerLoopCount := arrData.MaxIndex() - Outer
		Inner := Outer + 1
		
		Loop % InnerLoopCount
		{
			
			if ( arrData[Outer][vColumnSort] > arrData[Inner][vColumnSort])
			{
				temp := arrData[Inner]
				arrData[Inner] := arrData[Outer]
				arrData[Outer] := temp
			}
			
			Inner += 1
		}
	}

	vResults :=
	Loop % arrData.MaxIndex()
	{
		vResults .= "`n"
		vResults .= Join(arrData[A_Index], ",")
	}

	return SubStr(vResults, StrLen("`n")+1)
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
			cValue := arrData[rIndex][cIndex]
			
			if ( rIndex == 1 )
				colMaxWidth[cIndex] := StrLen(cValue)
			else if (colMaxWidth[cIndex] < StrLen(cValue) )
				colMaxWidth[cIndex] := StrLen(cValue)
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

HelperSpacedColumnForAnySV(vText, vDelim, hasHeader := false)
{
	arrData := HelperGetSVDataIntoRowsArray(vText, vDelim)
	colMaxWidth := HelperFindColumnMaxWidth(arrData)
	
	if ( hasHeader )
		return HelperPrintSpacedColumns(arrData, colMaxWidth, true)
	else
		return HelperPrintSpacedColumns(arrData, colMaxWidth, false)
}

SpacedColumnForAnySV(UserInput)
{
	
	vDelim := HelperAskUserForDelim()
	
	if ( StrLen(vDelim) == 0 )
	{
		SB_SetText("Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	if ( vDelim == "\t" or vDelim == "`t" )
		vDelim := "`t"
	
	MsgBox, 4, Data Table, First row header row? (Yes or No)
	IfMsgBox Yes
		hasHeader := true
	
	vText := UserInput
	
	return HelperSpacedColumnForAnySV(vText, vDelim, hasHeader)
}


HelperFindRowMaxWidth(arrData) {
	; Find the max width per column for each row
	rowMaxWidth := Object()

	Loop, % arrData.MaxIndex()
	{ ; Rows[X]
		xIndex := A_Index ; row Index
		
		Loop, % arrData[xIndex].MaxIndex()
		{ ; Row Items [X,Y]
			yIndex := A_Index ; column index
			
			if ( yIndex == 1 ) {
				rowMaxWidth[xIndex] := StrLen(arrData[xIndex][yIndex])
				
			} else if (rowMaxWidth[xIndex] < StrLen(arrData[xIndex][yIndex]) ) {
				rowMaxWidth[xIndex] := StrLen(arrData[xIndex][yIndex])
			}
		}
	}
	
	return rowMaxWidth ; row[#] := max width
}

HelperSpacedColumnVerticallyForAnySV(vText, vDelim, hasHeader:=false){
	; Displays row data as vertical lists, where the first row is normally descriptive text
	;
	; Psuedo Code
	; 1.) Put data into 2-D Array
	; 2.) Loop on number of columns (use row 1 as an example)
	; 3.) Loop on number of rows
	; 4.) Print out data, which the row changes faster than the column (print by column rather than by row)

	arrData := HelperGetSVDataIntoRowsArray(vText, vDelim)
	rowMaxWidth := HelperFindRowMaxWidth(arrData)

	colBetweenSpace := 2
	@ :=
	NewLine := "`r`n"

	; 2-D Array should be a square/rectangle, where positions are X,Y
	
	Loop, % arrData[1].MaxIndex()
	{ ; Columns
		cIndex := A_Index
		@ .= NewLine
		
		Loop, % arrData.MaxIndex()
		{ ; Rows
			rIndex := A_Index
			
			if ( rIndex == arrData.MaxIndex() )
				@ .= arrData[rIndex][cIndex]
			else {
				vLength := rowMaxWidth[rIndex] + colBetweenSpace
				vText := arrData[rIndex][cIndex]
				vDirection := "Right"
				
				if ( hasHeader && rIndex == 1 )
				{
					vLength := vLength - 2
					vDirection := "Left"
					@ .= HelperPadSpaces(vLength, vText, vDirection) ": "
					
				} else {
					@ .= HelperPadSpaces(vLength, vText, vDirection)
				}
			}
				
		}
	}
	
	return SubStr(@, StrLen(NewLine)+1)
}

SpacedColumnVerticalForAnySV(UserInput){
	
	vText := UserInput
	vDelim := HelperAskUserForDelim()
	
	if ( StrLen(vDelim) == 0 )
	{
		SB_SetText("Delimiter is not defined.")
		return ; No delimiter found/given.
	}
	
	if ( vDelim == "\t" or vDelim == "`t" )
		vDelim := "`t"
	
	hasHeader := false
	MsgBox, 4, Data Table, First column header column? (Yes or No)
	IfMsgBox Yes
		hasHeader := true
	
	return HelperSpacedColumnVerticallyForAnySV(vText, vDelim, hasHeader)
}

TextDuplicateXTimes(UserInput){
	; Input: (1) Single line of text, (2) Number of times to duplicate
	; Output: Duplicated text (on multiple lines)
	
	vText := UserInput
	Lines := StrSplit(vText, "`r`n")
	
	if (Lines.Length() > 1) {
		SB_SetText("Select/copy one line of text for processing.")
		return
	}

	Prompt = "Increment how many times?"
	InputBox, varTimes, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

	if varTimes is integer
	{
		if ( varTimes > 0 )
		{
			NewLine := "`r`n"
			@ := vText
			
			Loop % varTimes
				@ .= NewLine vText
			
			SB_SetText("Input duplicated X times.")
			return @
		}
		
	} else {
		SB_SetText("Not a valid number.")
		return
	}
}

TextDuplicateIncrementLastNumberXTimes(UserInput){
	; Input: Single line of text
	; Output: Duplicated text, where the last number is incremented x amount of times.
	
	vText := UserInput
	Lines := StrSplit(vText, "`r`n")
	
	if (Lines.Length() > 1) {
		SB_SetText("Select/copy one line of text for processing")
		return
	}
	
	; DEBUG
	;vText := "C:\Path\To\File\12345\12345_0001.tif"

	; Regular expression used: [1-9]+[0-9]*
	;  - To avoid counting padded zeroes, we'll check for numbers above zero and
	;    then count the zeroes thereafter as part of the pattern to match
	;  - When we replace the number this way, we keep the padded zeroes previously
	;    found in the string
	;  - In this future, we may have to account for padded number/max char length
	Matches := []
	MatchPositions := []
	pos = 1
	While pos := RegExMatch(vText,"[1-9]+[0-9]*", match, pos+StrLen(match))
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
			@ := vText
			
			Loop % varTimes
			{
				Replacement += 1
				strNew := RegExReplace(vText, "[\d]+", Replacement, , , LastPos)
				@ .= NewLine strNew
			}
			
			SB_SetText("Duplicated and incremented input X times.")
			return @
		}
		
	} else {
		SB_SetText("Not a valid number.")
		return
	}
}

HelperFlattenTextSingleLine(vText) {
	vText := Trim(vText)
	;Decide whether or not to leave a space between joined lines.
	vText := RegExReplace(vText, "`r`n", "") ; Remove CRLF and replace with nothing
	vText := RegExReplace(vText, "`n", "")   ; Remove CR
	vText := RegExReplace(vText, "`r", "")   ; Remove LF
	
	/* User should use Join w/ Char to put spaces in between. And Reduce Whitespace for that.
	vText := RegExReplace(vText, "\s", " ") ; Convert any whitespace char (space, tab, newline)
	vText := RegExReplace(vText, "\s+", " ") ; Converts grouped spaces (1 or more)
	*/
	return vText
}

TextFlattenSingleLine(UserInput)
{
	; Flatten to a single line (Trim, More than one space, tabs, New lines, Carriage Returns)
	; 1.) Trim (Preceeding and exceeding white spaces)
	; 2.) RegExReplace (More than one space, tabs, New lines, Carriage Returns)
	;
	; NOTE: Do not change to remove space between. TextFlattenMultiLine + Join Lines should allow without the space.
	
	/*
	vText := Trim(UserInput)
	vText := RegExReplace(vText, "\s", " ") ; Convert any whitespace char (space, tab, newline)
	vText := RegExReplace(vText, "\s+", " ") ; Converts grouped spaces (1 or more)
	*/
	
	vText := HelperFlattenTextSingleLine(UserInput)
	
	SB_SetText("Flattened input to one line.")
	return vText
}

TextFlattenMultiLine(UserInput) {
	; Flatten each line (Trim, More than one space, tabs, New lines, Carriage Returns)
	; 1.) Trim (Preceeding and exceeding white spaces)
	; 2.) RegExReplace (More than one space, tabs, New lines, Carriage Returns)
	
	vText := UserInput
	NewLine := "`r`n"
	@ := ""
	
	Loop, Parse, vText, `n, `r
	{
		vText := TextFlattenSingleLine(A_LoopField)
		@ .= NewLine vText
	}		
	
	SB_SetText("Flattened input each line.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextFormatIndentMore(UserInput)
{
	; https://autohotkey.com/board/topic/70404-clip-send-and-retrieve-text-using-the-clipboard/
	; Indent a block of text. (This introduces a small delay for typing a normal tab; to avoid this you can use ^Tab / ^+Tab as a hotkey instead.)
	
	vText := UserInput
	TabChar := A_Tab ; this could be something else, say, 4 spaces
	NewLine := "`r`n"
	
	@ := ""
	Loop, Parse, vText, `n, `r
		@ .= NewLine (InStr(A_ThisHotkey, "+") ? SubStr(A_LoopField, (InStr(A_LoopField, TabChar) = 1) * StrLen(TabChar) + 1) : TabChar A_LoopField)
	
	SB_SetText("Indented more.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextFormatIndentLess(UserInput){
	
	vText := UserInput
	TabChar := A_Tab
	NewLine := "`r`n"
	
	@ := ""
	Loop, Parse, vText, `n, `r
	{
		; If we find a tab at the beginning of the string, substring remove it to "indent less"
		FoundPos := RegExMatch(A_LoopField, "O)(^\t)", RegExMatchOutputVar)
		
		if (RegExMatchOutputVar.Count() > 0)
			@ .= NewLine SubStr(A_LoopField, StrLen(A_Tab)+1)
		else
			@ .= NewLine A_LoopField
	}

	SB_SetText("Indented less.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextFormatTelephoneNumber(UserInput)
{
	; Formats the Clipboard value to (###) ###-#### format
	; Ex: 1234567890 --> (123) 456-7890
	
	vText := UserInput
	vFormatted := RegExReplace(vText, "(\d{1})(\d{3})(\d{3})(\d{3})", "$1-($2) $3-$4")
	vFormatted := RegExReplace(vFormatted, "(\d{3})(\d{3})(\d{3})", "($1) $2-$3")
	
	SB_SetText("Formatted numbers (Telephone).")
	return vFormatted
}

TextGetDateFromMMDDYYYY(UserInput){
	; Get the full date from mm/dd/yyyy
	; Ex: 1/1/2010 = Friday, January 1, 2010
	;datestring := "1/1/2010"
	
	vText := UserInput

	FoundPos := RegExMatch(vText, "O)(\d+\/\d+\/\d+)", RegExMatchOutputVar)
	
	if ( RegExMatchOutputVar.Count() > 0 ) {
		;MsgBox % RegExMatchOutputVar.Value(0)
		SetFormat, float, 02.0
		StringSplit, d, strDate, / 
		FormatTime, FullDate, % d3 . d1+0. . d2+0., dddd, MMMM d, yyyy
		;MsgBox % FullDate " has been added to the Windows Clipboard."
		
		SB_SetText("Full Date returned.")
		return FullDate
	}
	else
		SB_SetText("No date found in Windows Clipboard to process. Aborted.")	
}

TextGenerateLoremIpsum()
{
	SB_SetText("Generated text (Lorem Ipsum).")
	return "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
}


; -----------------------------------------------------------------------------------------------------------

TextJoinBrokenLines(UserInput) {
	/* Combine pairs to one line
	 *  - Input:
	 * 	Every X lines, combine them
	 * 	Text to perform operation on
	 *  - Output:
	 * 	Reduced lines
	 *  - Suggestion:
	 * 	If the X does not divide evenly, reduce the remaining lines together
	 * 	 So if X=3 and there are 5 lines, combine 1-3 and 4-5
	 *  - Pseudo Code
	 * 	Ask if combining every 2 or more lines
	 * 	Split input text into lines
	 * 	Loop through lines
	 * 	  Build output text based on input, use "`n" after X lines
	 * 	Output results
	 */

	vInputText = %UserInput%
	vLines := StrSplit(vInputText, "`n")
	vLineCount := vLines.Count()
	
	; DEBUG
	;MsgBox % vInputText
	;MsgBox % "Line Count: " vLineCount
	
	Prompt := "Combine every x lines into one line."
	Dft := 2
	InputBox, OutputVar, Enter value for x, %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

	OutputVar := Trim(OutputVar)

	if ( StrLen(OutputVar) == 0 ) {
		SB_SetText("Not a valid response (No valude entered)")
		return ; Not a valid response (No valude entered)
	}

	if ( isNaN(OutputVar) ) {
		SB_SetText("Not a valid response (Not a number)")
		return ; Not a valid response (Not a number)
	}

	if ( OutputVar >= vLineCount ) {
		SB_SetText("Not a valid response (Number is greater than or equal to the total number of lines)")
		return ; Not a valid response (Number is greater than or equal to the total number of lines)
	}

	if ( OutputVar <= 1 ) {
		SB_SetText("Not a valid response (Number is less than or greater than zero)")
		return ; Not a valid response (Number is less than or greater than zero)
	}

	Prompt := "Character between each line? (Nothing is acceptable)"
	Dft := 
	InputBox, joinChar, Enter value for x, %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

	vInputText := Trim(vInputText)

	if ( StrLen(vInputText) == 0 ) {
		SB_SetText("Not a valid input (No input to process)")
		return ; Not a valid input (No input to process)
	}
	
	if ( vLineCount == 1 ) {
		SB_SetText("Not a valid input (Already one line)")
		return ; Not a valid input (Already one line)
	}

	vResults := ""
	Loop % vLines.Count() {
		
		if ( (A_Index != vLines.Count()) and (A_Index != 1) )
			vResults := vResults . joinChar
		
		vResults := vResults . vLines[A_Index]
		
		if ( Mod(A_Index, OutputVar) == 0 )
			vResults := vResults . "`n"
	}

	return vResults
}

; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------

isNaN(x){    
	if x is number
		return false 
	else	  
		return true 
} 

; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------

; Join lines of text to a single line, separated by a user-defined character
TextJoinLinesWithChar(UserInput) {
	; 1.) Get content (Text) and put into clipboard
	; 2.) Prompt user for split delimiter
	; 3.) Prompt user for join delimiter
	; 4.) Loop by new line
	;      - For each line, split and join
	
	vText := UserInput
	
	Prompt = "Join lines by..."
	InputBox, varJoiner, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	SB_SetText("Joined lines w/ or w/o character.")
	return HelperTextJoinLinesWithChar(vText, varJoiner)
}

HelperTextJoinLinesWithChar(vText, vJoinChar)
{
	vResults := 
	
	Loop, Parse, vText, `n, `r 
	{
		if (A_Index > 1)
			vResults .= vJoinChar A_LoopField
		else
			vResults = %A_LoopField% ; Join character starts being used after the first item
	}
	
	if (StrLen(varJoiner) > 0)
		vResults := Substr(vResults, StrLen(vJoinChar) + 1) ; Remove excess "varJoiner" from the beginning of the first line
	
	return vResults
}

; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
; Join the items of an array into a text ilne, separated by custom separator
Join(Array, Sep)
{
	for k, v in Array
		out .= Sep . v
	return SubStr(out, 1+StrLen(Sep))
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
; *SV line of text can be recompiled with a different separator
TextJoinSplittedLineWithChar(UserInput) {
	; 1.) Get content (Text) and put into clipboard
	; 2.) Prompt user for split delimiter
	; 3.) Prompt user for join delimiter
	; 4.) Loop by new line
	;      - For each line, split and join
	
	vText := UserInput
		
	Prompt = "Split text by..."
	InputBox, varSplitter, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(varSplitter) == 0 ) {
		SB_SetText("Missing defined character that splits the data.")
		return
	}
		
	Prompt = "Join lines by..."
	InputBox, varJoiner, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

	;if ( StrLen(varJoiner) > 0 ) { ; Lets allow the user to either join with a character or not
		
		TextArray := StrSplit(vText, varSplitter)
		@ := Join(TextArray, varJoiner)
		
		SB_SetText("Joined splitted line with character.")
		return @
	;}
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
; *SV lines of text can be recompiled with a different separator
TextJoinSplittedLinesWithChar(UserInput) {
	; 1.) Get content (Text) and put into clipboard
	; 2.) Prompt user for split delimiter
	; 3.) Prompt user for join delimiter
	; 4.) Loop by new line
	;      - For each line, split and join
	
	vText := UserInput
	
	Prompt = "Split text by..."
	InputBox, varSplitter, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(varSplitter) > 0 ) {
		
		Prompt = "Join lines by..."
		InputBox, varJoiner, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
		if ( StrLen(varJoiner) > 0 ) {
			
			NewLine := "`r`n"
			@ :=
			
			Loop, Parse, vText, `n, `r 
			{
				TextArray := StrSplit(A_LoopField, varSplitter)
				@ .= NewLine Join(TextArray, varJoiner)
			}
			
			
			SB_SetText("Joined splitted lines with character.")
			return SubStr(@, StrLen(NewLine) + 1)
		}
	}
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
OpenLinkAHKRegexQuickReference()
{
	Run, https://www.autohotkey.com/docs/misc/RegEx-QuickRef.htm
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
OpenLinkRegExr()
{
	Run, https://regexr.com/
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
; Replaces unfriendly Windows OS Filesystem characters with an underscore
; "A file name can't contain any of the following characters: \ / : * ? " < > |
TextMakeWindowsFileFriendly(UserInput) {
	vText := UserInput
	vNewLine := "`r`n"
	@ := ""
	
	Loop, Parse, vText, `n, `r 
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
		@ .= vNewLine tempText
	}
	
	vResult := SubStr(@, StrLen(vNewLine) + 1)

	SB_SetText("Characters are now Windows Filesystem friendly.")
	return vResult
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
TextPrefixNestedList(UserInput){
	NewLine := "`r`n"
	charPrefixLevel := "\_" ; Each line
	charPrefixList := "|_" ; Each tabbed-in line
	countLevel := 0
	
	; Level 0
	; \_Level 1
	;  \_Level 2
	; ....
	
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r 
	{
		
		; First see if there is text. Then, prefix the text part only (for example, indented text)
		MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar) ; Testing to see if the line is not empty by having characters other than invisible ones
		if ( RegExMatchOutputVar.Count() > 0 )
		{	
			if (countLevel == 0)
			{
				@ .= NewLine
				@ .= A_LoopField
				countLevel += 1
				continue
			}			
			
			vLine := A_LoopField
			vLength := countLevel + 1 + StrLen(vLine) ; 2 for "\_"
			vDirection := "Left"
			
			@ .= NewLine
			@ .= HelperPadSpaces(vLength, "\_" . A_LoopField, vDirection)
			
			countLevel += 1
		}
	}
	
	SB_SetText("Created a nested list.")
	return SubStr(@, StrLen(NewLine) + 1)
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
; Create an unordered text list, using a custom defined "bullet" character that won't change each line
TextPrefixCustomRepeatingCharacter(UserInput){
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
	
	Prompt = "Enter custom character for Text List conversion"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(OutputVar) < 1 or OutputVar == )
	{
		SB_SetText("Invalid user input (TextPrefixCustomRepeatingCharacter).")
		return
	}
	
	NewLine := "`r`n"
	vCustomChar := OutputVar
	OrdinalText := vCustomChar . A_Space
	vText := UserInput
	@ := ""
	
	Loop, Parse, vText, `n, `r 
	{
		; First see if there is text. Then, prefix the text part only (for example, indented text)
		MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar)
		if ( RegExMatchOutputVar.Count() > 0 )
		{	
			@ .= NewLine
			@ .= RegExReplace(A_LoopField, "([\S].+)", OrdinalText . "$1" )
		}
	}
	
	SB_SetText("Created list with Unordered Custom Character.")
	return SubStr(@, StrLen(NewLine) + 1)
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
	
	; Ask the user if indented items should be renumbered
	; 1 -> 2 -> 3 vs. 1 -> 1.1 -> 1.2
	renumber := false
	MsgBox, 4, Text to List, Renumber indented lines of text? (Yes or No)
	IfMsgBox Yes
		renumber := true

	NewLine := "`r`n"
	Numeral := 1 ; Ordinal number listing starts with 1
	
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r 
	{
		; First see if there is text. Then, prefix the text part only (for example, indented text)
		MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar)
		if ( RegExMatchOutputVar.Count() > 0 )
		{	
			OrdinalText := Numeral . ". "
			@ .= NewLine
			@ .= RegExReplace(A_LoopField, "([\S].+)", OrdinalText . "$1" )
			Numeral += 1
		}
	}
	
	SB_SetText("Created list with ordered numbers.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextPrefixOrdinalNumberRenumberIndented(UserInput){

	; Line 1				->	1. Line 1
	; 	Indented Line 2		->		1.1. Indented Line 2
	;   Indented Line 3		->  	1.2. Indented Line 3
	;		Indented Line 4 ->			1.2.1. Indented Line 4
	
	; 1.) First organize lines into an array so we know how to number
	; 2.) Loop through array and prefix lines with proper numbering
	; 3.) Print results
	
	vText := UserInput
	;vText := "Line 1`nLine 2`nLine 3" ; 1 Level
	;vText := "Line 1`n`tLine 2`n`tLine 3" ; 2 Levels
	;vText := "Line 1`n`tLine 2`n`tLine 2`n`t`tLine 3`n`t`tLine 3" ; 3 Levels
	;vText := "Line 1`n`tLine 2`n`tLine 2`n`t`tLine 3`n`t`tLine 3`nLine 1`n`tLine 2`n`tLine 2`n`t`tLine 3`n`t`tLine 3" ; 3 Levels - 2 Lists
	;vText := "Line 1`n`tLine 2`n`t`tLine 3`n`t`tLine 3`n`tLine 2`n`tLine 2`n`t`tLine 3`n`t`tLine 3" ; in and out levels
	
	vResult :=
	
	NewLine := "`r`n"
	
	vArrIndentCounter := Array()
	; vArrIndentCounter[]
	; First position is 0 level and cannot be cleared
	; Second position is tab level 1, and is cleared when a line with no tabs is found
	; Third position is tab level 2, and is cleared when a line has no tab or only 1 tab level
	; Fourth position is tab level 3, and is cleared when a line has no tab or only 1/2 tab levels
	; ....
	
	; 1.) First organize lines into an array so we know how to number, based on tabs present
	Loop, Parse, vText, `n, `r
	{
		vLine := A_LoopField
		vPrefix := 
		
		if ( SubStr(vLine, 1, 1) != A_Tab )
		{
			
			if ( !vArrIndentCounter.MaxIndex() ) {
				vArrIndentCounter.Push(1)
			} else {
				vArrIndentCounter[1] := vArrIndentCounter[1] + 1
			
				Loop % (vArrIndentCounter.MaxIndex() - 1)
					vArrIndentCounter.Pop() ; Remove the levels after the first
			}
		
			vPrefix := vArrIndentCounter[1] . ". "
			
			if ( vPrefix != "1. " )
				vResult .= NewLine
			
			vResult .= NewLine vPrefix A_LoopField
			
		} else {
		
			vIndentCounter := 1 ; Not zero based ; 1 is first level or not indented
		
			Loop % StrLen(vLine)
			{
				if ( SubStr(vLine, A_Index, 1) == A_Tab )
					vIndentCounter += 1
				else
					break
			}
			
			; Build vArrIndentCounter
			if ( !vArrIndentCounter[vIndentCounter] )
				vArrIndentCounter.Push(1)
			else
				vArrIndentCounter[vIndentCounter] := vArrIndentCounter[vIndentCounter] + 1
			
			; Close lists not directly under the current one (say new line is indented one less than the previous line)
			if ( vIndentCounter < vArrIndentCounter.MaxIndex() )
			{
				Loop % (vArrIndentCounter.MaxIndex() - vIndentCounter)
					vArrIndentCounter.Pop()
			}
			
			vPrefix := 
			
			; Loop through vArrIndentCounter to build prefix for indented lines
			Loop % vArrIndentCounter.MaxIndex()
				vPrefix := vPrefix . vArrIndentCounter[A_Index] . "."
			
			vResult .= NewLine RegExReplace(A_LoopField, "([^\s][\S].+)", (vPrefix . " $1") ) ; Put prefix after indentation
			
		}
	}
	
	SB_SetText("Created indented level list - Ordinal Numbers")
	return SubStr(vResult, StrLen(NewLine) + 1) ; Remove Newline at the beginning
}

; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
	
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r 
	{
		; First see if there is text. Then, prefix the text part only (for example, indented text)
		MatchedPosition := RegExMatch(A_LoopField, "O)([\S].+)", RegExMatchOutputVar)
		if ( RegExMatchOutputVar.Count() > 0 )
		{	
			OrdinalText := Chr(Numeral) . ". "
			@ .= NewLine
			@ .= RegExReplace(A_LoopField, "([\S].+)", OrdinalText . "$1" )
			Numeral += 1
		}
	}
	
	SB_SetText("Created list with lowercase letters.")
	return SubStr(@, StrLen(NewLine) + 1)
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
	
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r 
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
	
	SB_SetText("Created list with uppercase letters.")
	return SubStr(@, StrLen(NewLine) + 1)
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
TextReplaceCharsNewLine(UserInput){

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
	vText := UserInput
	
	Prompt = "Enter character(s) to find/replace with new line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%

	vUserFind := OutputVar
	vAltText := StrReplace(vText, vUserFind, "`r`n") ; Symbols are free to use in StringReplace
	
	If ErrorLevel
	{
		SB_SetText("(TextReplaceRegex) There was an issue.")
		return
	}
	
	SB_SetText("Replaced text with new line.")
	return vAltText
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
HelperTextReplaceRegex(vText, vRegExSearch, vRegExReplace){ ; Modified
	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string

	if ( !vRegExReplace )
		vRegExReplace := ""

	vAltText := RegExReplace(vText, vRegExSearch, vRegExReplace)

	If ErrorLevel
	{
		SB_SetText("(TextReplaceRegex) There was an issue.")
		return
	}
	
	return vAltText
}

HelperTextReplaceRegexEachLine(vText, vRegExSearch, vRegExReplace){ ; Modified

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string

	if ( !vRegExReplace )
		vRegExReplace := ""

	NewLine := "`r`n"
	@ := ""
	Loop, Parse, vText, `n, `r
	{
		@ .= NewLine RegExReplace(A_LoopField, vRegExSearch, vRegExReplace)
		If ErrorLevel
		{
			SB_SetText("(TextReplaceRegexEachLine) There was an issue.")
			return
		}
	}
	
	return SubStr(@, StrLen(NewLine) + 1)
}

HelperTextReplaceStringReplaceEachLine(vText, vSearch, vReplace){ ; Modified

	; Find/Replace with StringReplace (versus RegExReplace)

	if ( !vReplace )
		vReplace := ""

	NewLine := "`r`n"
	@ := ""
	Loop, Parse, vText, `n, `r
	{
		vReplacedText := StrReplace(A_LoopField, vSearch, vReplace)
		@ .= NewLine vReplacedText
	}
	
	return SubStr(@, StrLen(NewLine) + 1)
}

TextReplaceRegex(UserInput){

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
	
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
	
	If ErrorLevel
	{
		SB_SetText("(TextReplaceRegex) There was an issue.")
		return
	}
	
	SB_SetText("Replaced text with string (RegEx).")
	return vAltText
}

TextReplaceRegexEachLine(UserInput){

	; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
	
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

	SB_SetText("Replaced text on each line (RegEx).")
	return HelperTextReplaceRegexEachLine(vText, vUserRegex, vUserReplacement)
}

TextReplaceStringReplace(UserInput){
	; Find/Replace with StringReplace
	vText := UserInput
	Prompt = "Enter search text for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserSearch := OutputVar
	
	if (!vUserSearch){
		SB_SetText("User did not enter text to search input for.")
		return ; Cannot replace if search criteria is not defined.
	}

	; replace double quotes with "\x22"

	Prompt = "Enter replacement text for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserReplacement := OutputVar

	if ( !vUserReplacement )
		vUserReplacement := ""
	
	vAltText := StrReplace(vText, vUserSearch, vUserReplacement)
	
	SB_SetText("Replaced all occurrences of search text with replacement text (StringReplace).")
	return vAltText
}

TextReplaceStringReplaceEachLine(UserInput){
	; Replace 'search text' with 'replacement text'
	vText := UserInput
	Prompt = "Enter text for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserSearch := OutputVar

	if (!vUserSearch){
		SB_SetText("User did not enter text to search input for.")
		return ; Cannot replace if search criteria is not defined.
	}
	
	; replace double quotes with "\x22"
	Prompt = "Enter replacement string for find/replace"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vUserReplacement := OutputVar

	if ( !vUserReplacement )
		vUserReplacement := ""

	SB_SetText("Replaced text on each line (StringReplace).")
	return HelperTextReplaceStringReplaceEachLine(vText, vUserSearch, vUserReplacement) ; (vText, vSearch, vReplace)
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

	If ErrorLevel
	{
		SB_SetText("(TextReplaceStringXTimes) There was an issue.")
		return
	}
	
	vStatusText := "String replaced " vUserNumber " time(s)."
	
	SB_SetText(vStatusText)
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
	
	return RegExReplace(vText, vFind, vReplace,,1) ; Replaces the first occurrence as "limit" is set to 1
}

TextReplaceLastStringOccurrenceEachLine(UserInput)
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
	
	vResult := HelperLoopEachLineCallFunction(vText, "CallbackReplaceLastOccurrence", vFind, vReplace)
	
	SB_SetText("Replaced last occurrence (Each Line).")
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
	
	SB_SetText("Replaced first occurrence (Each Line).")
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
	
	if ( vFoundPos ) {
		SB_SetText("Replaced last occurrence.")
		return RegExReplace(vText, vFind, vReplace, 0, 1, vFoundPos) ; 
		
	} else {
		SB_SetText("Could not find text.")
		return vText
	}
}

TextReplaceFirstStringOccurrence(UserInput)
{
	; Replace last occurrence in a string - Using variadic variables and callback functions
	vText := UserInput
	
	Prompt = "What do you want to replace?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vFind := OutputVar
	
	if ( StrLen(vFind) == 0 ) {
		SB_SetText("Nothing entered to replace. Stopped.")
		return
	}
	
	Prompt = "What do you want to replace it with?"
	InputBox, OutputVar, Find/Replace String X Times..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	vReplace := OutputVar
	
	SB_SetText("Replaced first occurrence.")
	return RegExReplace(vText, vFind, vReplace,,1)
}


TextTrim(UserInput) {

	NewLine := "`r`n"
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r
		@ .= NewLine vText := Trim(A_LoopField)
	
	SB_SetText("Trimmed text.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextTrimLeft(UserInput) {

	NewLine := "`r`n"
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r
		@ .= NewLine vText := LTrim(A_LoopField)
	
	SB_SetText("Trimmed left of Text.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextTrimRight(UserInput) {

	NewLine := "`r`n"
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r
		@ .= NewLine vText := RTrim(A_LoopField)
	
	SB_SetText("Trimmed right of Text.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextRemoveSpaces(UserInput) {
	NewLine := "`r`n"
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r
		@ .= NewLine vText := RegExReplace(A_LoopField, " ", "")
	
	SB_SetText("Removed spaces.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextReplaceSpacesUnderscore(UserInput) {
	NewLine := "`r`n"
	vText := UserInput
	
	@ := ""
	Loop, Parse, vText, `n, `r
		@ .= NewLine vText := RegExReplace(A_LoopField, " ", "_")
	
	SB_SetText("Replaced spaces with underscores.")
	return SubStr(@, StrLen(NewLine) + 1)
}

TextReplaceSpacesDash(UserInput) {
	NewLine := "`r`n"
	vText := UserInput	
	@ := ""
	
	Loop, Parse, vText, `n, `r
	{
		@ .= NewLine vText := RegExReplace(A_LoopField, " ", "-")
	}
	
	SB_SetText("Replaced spaces with dashes.")
	return SubStr(@, StrLen(NewLine) + 1)
}

RemoveBlankLines(UserInput)
{	
	vText := UserInput
	
	; Reduce double CRLF (`r`n`r`n)
	Loop
	{
		; When the last parameter is UseErrorLevel, ErrorLevel is given the number occurrences replaced (0 if none).
		; Otherwise, ErrorLevel is set to 1 if SearchText is not found within InputVar, or 0 if it is found.
		;StringReplace, vText, vText, `r`n`r`n, `r`n, UseErrorLevel
		;if (ErrorLevel = 0)  ; No more replacements needed.
		;	break
		
		StringReplace, vText, vText, `r`n`r`n, `r`n, All
		if (ErrorLevel = 1)
			break
	}
	
	; Reduce double LF (`n`n)
	Loop
	{
		StringReplace, vText, vText, `n`n, `n, All
		if (ErrorLevel = 1)
			break
	}
	
	SB_SetText("Removed blank lines.")
	return vText
}

RemoveDuplicateLines(UserInput)
{
	arrLines := StrSplit(UserInput, ["`r`n", "`n", "`r"])
	trimmedArray := HelperTrimArray(arrLines)
	joinLines := Join(trimmedArray, "`n")
	
	SB_SetText("Removed duplicate lines.")
	return SubStr(joinLines, 1, StrLen(joinLines) - StrLen("`n") + 1)
}

; https://stackoverflow.com/questions/46432447/how-do-i-remove-duplicates-from-an-autohotkey-array
HelperTrimArray(arr) { ; Hash O(n)

    hash := {}, newArr := []

    for e, v in arr
        if (!hash.Haskey(v))
            hash[(v)] := 1, newArr.push(v)

    return newArr
}

RemoveRepeatingDuplicateLines(UserInput)
{
	vText := UserInput . "`n"
	
	Loop, Parse, vText, `n, `r
	{
		delim := A_LoopField . "`n"
		vText := st_removeDuplicates(vText, delim)
	}
	
	TextMinusDelim := StrLen(Text) - StrLen("`n")
	
	SB_SetText("Removed repeating lines.")
	return SubStr(vText, 1, TextMinusDelim)
}

RemoveRepeatingDuplicateChars(UserInput)
{
	Prompt = "Repeating Character?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if (StrLen(OutputVar) == 0)
		return
	
	vText := UserInput
	
	SB_SetText("Removed repeating characters.")
	return st_removeDuplicates(vText, OutputVar)
}

HelperRemoveLinesIncludeExcludeText(vText, vNeedle, vChoice)
{
	; 1 = Include Text (remove)
	; 2 = Exclude Text (remove)
	
	if (vChoice < 1)
		vChoice := 1
	
	if (vChoice > 2)
		vChoice := 2
	
	if (vText == "" or StrLen(Trim(vText)) == 0)
		return ; No data to process
	
	NewLine := "`r`n"
	@ :=
	Loop, Parse, vText, `n, `r
	{
		if (vChoice == 1){ ; Skip if line includes text
			If InStr(A_LoopField, vNeedle)
				continue
			Else
				@ .= NewLine A_LoopField
			
		} else if (vChoice == 2){ ; Skip line if it excludes text
			If InStr(A_LoopField, vNeedle)
				@ .= NewLine A_LoopField
			Else
				continue
		}
	}
	
	return SubStr(@, StrLen(NewLine) + 1)
}

RemoveLinesThatIncludeText(UserInput)
{
	Prompt = "Remove text lines that include what characters?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(OutputVar) == 0 ) {
		SB_SetText("Character(s) input cannot be nothing.")
		return
	}
	
	if ( InStr(UserInput, OutputVar) == false ) {
		SB_SetText("Character(s) should be in string.")
		return
	}
	
	vText := UserInput
	vChoice := 1
	vNeedle := OutputVar
	vResult := HelperRemoveLinesIncludeExcludeText(vText, vNeedle, vChoice)
	
	if ( StrLen(vResult) == 0 )
		SB_SetText("All lines included the text.")
	else
		SB_SetText("Removed lines that included the text.")

	return vResult
}

RemoveLinesThatExcludeText(UserInput)
{
	Prompt = "Remove text lines that exclude what characters?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(OutputVar) == 0 )
	{
		SB_SetText("Character(s) input cannot be nothing.")
		return
	}
	
	vText := UserInput
	vChoice := 2
	vNeedle := OutputVar
	vResult := HelperRemoveLinesIncludeExcludeText(vText, vNeedle, vChoice)
	
	if ( StrLen(vResult) == 0 )
		SB_SetText("All lines excluded the text.")
	else
		SB_SetText("Removed lines that excluded the text.")
	
	return vResult
}

HelperRemoveLinesMatchDontMatchRegEx(vText, vRegEx, vChoice)
{
	; 1 = Match RegEx (remove)
	; 2 = Not Match RegEx (remove)
	
	if (vChoice < 1)
		vChoice := 1
	
	if (vChoice > 2)
		vChoice := 2
	
	if (vText == "" or StrLen(Trim(vText)) == 0)
		return ; No data to process
	
	NewLine := "`r`n"
	@ :=
	Loop, Parse, vText, `n, `r
	{
		
		if (vChoice == 1){ ; Skip if line includes text
			If RegExMatch(A_LoopField, vRegEx)
				continue
			Else
				@ .= NewLine A_LoopField
			
		} else if (vChoice == 2){ ; Skip line if it excludes text
			If RegExMatch(A_LoopField, vRegEx)
				@ .= NewLine A_LoopField
			Else
				continue
		}
	}
	
	return SubStr(@, StrLen(NewLine) + 1)
}

RemoveLinesThatMatchRegEx(UserInput)
{
	Prompt := "Remove text lines that match what Regular Expression?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(OutputVar) == 0 )
	{
		SB_SetText("Regular Expression input cannot be nothing.")
		return
	}
	
	vText := UserInput
	vChoice := 1
	vNeedle := OutputVar
	vResult := HelperRemoveLinesMatchDontMatchRegEx(vText, vNeedle, vChoice)
	
	if ( StrLen(vResult) == 0 )
		SB_SetText("All lines matched the regular expression.")
	else
		SB_SetText("Removed lines that matched RegEx.")
	
	return vResult
}

RemoveLinesThatDontMatchRegEx(UserInput)
{
	Prompt := "Remove text lines that don't match what regular expression?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ( StrLen(OutputVar) == 0 )
	{
		SB_SetText("Regular Expression input cannot be nothing.")
		return
	}
	
	vText := UserInput
	vChoice := 2
	vNeedle := OutputVar
	vResult := HelperRemoveLinesMatchDontMatchRegEx(vText, vNeedle, vChoice)
	
	if ( StrLen(vResult) == 0 )
		SB_SetText("All lines didn't match the regular expression.")
	else
		SB_SetText("Removed lines that don't match RegEx.")
	
	return vResult
}

RemoveWhitespace(UserInput)
{
	; User should use Join w/ Char to put spaces in between. And Reduce Whitespace for that.
	vText := UserInput
	vText := Trim(vText)
	vText := RegExReplace(vText, "\s", " ") ; Convert any whitespace char (space, tab, newline)
	vText := RegExReplace(vText, "\s+", " ") ; Converts grouped spaces (1 or more)
	
	return vText
}

RemoveWhitespaceEachLine(UserInput)
{
	vText := UserInput
	vCallBack := "RemoveWhitespace"
	vResult := HelperLoopEachLineCallFunction(vText, vCallBack, "")
	
	return vResult
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
		
		SB_SetText("Invalid value for pad length.")
		return
	}
	
	vText := HelperPadSpacesLeftOrRight(OutputVar, UserInput, "Left")
	
	SB_SetText("Padded with spaces (Left).")
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
	
	SB_SetText("Padded with spaces (Right).")
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
		
		SB_SetText("Character(s) input cannot be nothing.")
		return
	}
	
	if ( InStr(UserInput, OutputVar) == false )
	{
		
		SB_SetText("Character(s) should be in string.")
		return
	}
	
	vText := UserInput
	vFindChar := OutputVar
	vMaxChar := HelperFindHighestPositionInTextLines(vFindChar, vText)
	
	if (vMaxChar == 0 )
	{
		
		SB_SetText("MaxChar is 0; No need to process or bug.")
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
	
	SB_SetText("Padded lines to align text to character.")
	return SubStr(@, StrLen(NewLine) + 1)
}

HelperTextCharLineUp(vText, vFindChar)
{
	NewLine := "`r`n"
	MaxChar := 0
	
	; Loop to determine the line with the most amount of characters before special character (Item1: Fun = 5 characters before ":")
	Loop, Parse, vText, `n, `r
	{
		; Returns the position of an occurrence of the string Needle in the string Haystack.
		; Position 1 is the first character; this is because 0 is synonymous with "false", making it an intuitive "not found" indicator.
		FoundPos := InStr(A_LoopField, vFindChar)
		
		; Compare to our placeholder variable for spacing, MaxChar
		; Only overwrite MaxChar if the newly found position is greater than our previously found, and furthest found, position
		If ( FoundPos <> 0 and FoundPos > MaxChar)
			MaxChar := FoundPos
	}
	
	If ( MaxChar > 0 )
	{
		; Loop again to precede line with spaces until all lines match (line 1 has 5 characters/spaces before ":" and so done line 2, 3, 4, etc.)
		@ := ""
		Loop, Parse, vText, `n, `r
		{
			; If less than desired length, pad with spaces
			PaddedSpace :=
			if ( InStr(A_LoopField, ":") < MaxChar )
			{
				PlaceDiff := MaxChar - InStr(A_LoopField, vFindChar)
				Loop, %PlaceDiff%
				{
					PaddedSpace .= A_Space
				}
				@ .= NewLine PaddedSpace A_LoopField
			}
			else
			{
				@ .= NewLine A_LoopField
			}
		}
		
		SB_SetText("Padded text to align to character.")
		return SubStr(@, StrLen(NewLine) + 1)
	}
}

TextCharLineUp(UserInput) {
	; Maybe allow user to enter a special character or have default (":", "-", ...)
	; Loop to determine the line with the most amount of characters before special character (Item1: Fun = 5 characters before ":")
	; Loop again to precede line with spaces until all lines match (line 1 has 5 characters/spaces before ":" and so done line 2, 3, 4, etc.)
	
	Prompt := "Which (1st occurence) character to line up text with?"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
		
	if ( OutputVar <> "" )
		vFindChar := OutputVar
	else
	{
		SB_SetText("No character entered by user.")
		return
	}
	
	vText := UserInput
	
	return HelperTextCharLineUp(vText, vFindChar)
}

TextSay(UserInput)
{
	vText := UserInput

	Length := StrLen(vText)
	
	if ( Length > 0 ) {
		
		AnnaVoice := TTS_CreateVoice("Microsoft Anna")
		TTS(AnnaVoice, "SpeakWait", vText)
		
		SB_SetText("Text spoken.")
	}
	else
		SB_SetText("No text available in the clipboard to say")
}

TextSayStop(UserInput){
	AnnaVoice := TTS_CreateVoice("Microsoft Anna")
	TTS(AnnaVoice,"Stop")
}

TextSurroundingChar(UserInput) {

	vText := UserInput
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
		
		;Send %OutputVar%%vText%%Ending%
		;Clip(OutputVar . vText . Ending)
		
		SB_SetText("Surrounded Text with character.")
		return OutputVar . vText . Ending
	}
}

TextPrefixSuffixRemoveEachLine(UserInput) {

	vText := UserInput

	Prompt = "Enter prefix text to remove from each line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	Prefix := OutputVar
	
	Prompt = "Enter suffix text to remove from each line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	Suffix := OutputVar

	if ( (StrLen(Prefix) == 0) and (StrLen(Suffix) == 0) )
	{
		SB_SetText("No data entered for Prefix or Suffix.")
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
	
	Loop, Parse, vText, `n, `r
	{
      line := A_LoopField
		
      ; Prefix: Left-to-Right search
      if ( doPrefix && (InStr(line, Prefix) = 1) )
         line := SubStr(line, (StrLen(Prefix)+1) )
      
      ; Suffix: Right-to-Left search; Make sure it's at the end of the line
      if ( doSuffix && InStr(line, Suffix, 0) && ((InStr(line, Suffix, 0) - 1 + StrLen(Suffix)) = StrLen(line)) )
         line := SubStr(line, 1, (InStr(line, Suffix, 0) - 1))
      
		@ .= NewLine
		@ .= line
	}
	
	retValue := SubStr(@, StrLen(NewLine) +1) ; Remove last new line (`r`n) or blank line
	
	SB_SetText("Removed Prefix/Suffix (Text/RegEx).")
	return retValue
}

TextPrefixSuffixRemoveEachLineRegex(UserInput) {
	; Wrap the text with a open/close HTML tag

	vText := UserInput

	Prompt = "Enter prefix text to remove from each line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	Prefix := OutputVar
	
	Prompt = "Enter suffix text to remove from each line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	Suffix := OutputVar

	if ( (StrLen(Prefix) == 0) and (StrLen(Suffix) == 0) )
	{
		SB_SetText("No data entered for Prefix or Suffix.")
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
	
	Loop, Parse, vText, `n, `r
	{
		; Only surround text, not any indenting non-whitespace characters
		@ .= NewLine
		@ .= RegExReplace(A_LoopField, "(" . Prefix . "([\S].+)" . Suffix . ")", "$2" )
	}
	
	retValue := SubStr(@, StrLen(NewLine) +1) ; Remove last new line (`r`n) or blank line
	
	SB_SetText("Removed Prefix/Suffix (Text/RegEx).")
	return retValue
}

TextPrefixSuffixEachLine(UserInput){
	
	vText := UserInput	

	Prompt = "Enter text to prefix each line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	Prefix := OutputVar
	
	Prompt = "Enter text to suffix each line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	Suffix := OutputVar

	if ( (StrLen(Prefix) == 0) and (StrLen(Suffix) == 0) )
	{
		;MsgBox % "No data has been entered for Prefix or Suffix. Aborting."
		
		SB_SetText("No data entered for Prefix or Suffix.")
		return
	}
	
	if ( (StrLen(Prefix) > 0) or (StrLen(Suffix) > 0) )
		return HelperTextPrefixSuffixEachLine(vText, Prefix, Suffix)	; (data, Prefix, Suffix)
	else
		return
}

HelperTextPrefixSuffixEachLine(data, Prefix, Suffix)
{
	NewLine := "`r`n"
	@ := 
	
	Loop, Parse, data, `n, `r
	{
		; Only surround text, not any indenting non-whitespace characters
		@ .= NewLine
		@ .= RegExReplace(A_LoopField, "([^\s][\S].+)", (Prefix . "$1" . Suffix) )
	}
	
	retValue := SubStr(@, StrLen(NewLine) + 1) ; Remove last new line (`r`n) or blank line
	return retValue
}

TextSortRegexCaptureGroup1(UserInput) {
	
	Prompt = "Enter character(s) to find/replace with new line"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	if ErrorLevel
		return ; Cancel was pressed
	
	if ( StrLen(Trim(OutputVar)) == 0 ) {
		SB_SetText("User hit enter with nothing in the box.")
		return ; User hit enter with nothing in the box
	}
	
	if ( RegExMatch(OutputVar, "\(.+\)") < 1 ) {
		SB_SetText("No capture groups found.")
		return ; No capture groups found
	}
		
	vText := UserInput
	vRegExSearch := OutputVar
	
	SB_SetText("Text sorted by capture group 1 of Regular Expression")
	return HelperTextSortRegexCaptureGroup1(vText, vRegExSearch)
}

HelperTextSortRegexCaptureGroup1(vText, vRegExSearch) {
	; Sort Lines by RegEx Capture Group #1
	; - Maybe allow user to specify capture group for comparison

	;vRegExSearch := "\{[a-zA-Z0-9\-]+\|(\d+)\}"
	arrLines := StrSplit(vText, "`n")

	Loop % arrLines.MaxIndex()
	{
		Outer := A_Index
		InnerLoopCount := arrLines.MaxIndex() - Outer
		Inner := Outer + 1
		;MsgBox % "InnerLoopCount: " InnerLoopCount
		
		Loop % InnerLoopCount
		{
			; MsgBox % "Outer: " Outer ", Inner: " Inner
			
			
			FoundPos := RegExMatch(arrLines[Outer], vRegExSearch, First)
			
			if ErrorLevel
				continue ; There was a RegEx processing error
			
			if FoundPos == 0
				First1 := "ZZZ"
			
			FoundPos := RegExMatch(arrLines[Inner], vRegExSearch, Second)
			
			if ErrorLevel {
				continue ; There was a RegEx processing error
			}
			
			if FoundPos == 0
				Second1 := "ZZZ" ; Sort to the bottom of the list if the pattern couldn't be found
			
			
			; MsgBox % "Outer Match: " First1 ", Inner Match: " Second1
			
			if ( First1 > Second1 )
			{
				temp := arrLines[Inner]
				arrLines[Inner] := arrLines[Outer]
				arrLines[Outer] := temp
				;MsgBox % "Outer: " arrData[Outer][vColumnSort] ", Inner: " arrData[Inner][vColumnSort]
			}
			
			Inner += 1
		}
	}

	vResult := Join(arrLines, "`n")
	return vResult
}

TextSortAlphabetically(UserInput) {
	; Sort selected text and replace it	
	vText := UserInput
	Sort, vText, ; Normal sort
	
	SB_SetText("Sorted alphabetically.")
	return vText
}

TextSortAlphabeticallyUnique(UserInput) {
	; Sort selected text and replace it	
	vText := UserInput
	Sort, vText, C U ; Normal sort
	
	SB_SetText("Unique Sort Alphabetically.")
	return vText
}

TextSortAlphabeticallyUniqueCaseInsensitive(UserInput) {
	; Sort selected text and replace it	
	vText := UserInput
	Sort, vText, U ; Normal sort
	
	SB_SetText("Unique Sort Alphabetically Case Insensitive.")
	return vText
}

TextSortAlphabeticallyWithDelimiter(UserInput) {
	; Sort selected text and replace it
	vText := UserInput		
	Prompt = "Enter delimiter separating values"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	Length := StrLen(OutputVar)
	
	if ( StrLen(Length) == 0 )
	{
		SB_SetText("Delimiter not entered.")
		return
	}
	
	Sort, vText, D%OutputVar% ; Alphabetically sort for words/text separated by a delimiter
	
	SB_SetText("Sorted alphabetically with delimiter.")
	return vText
}

TextSortNumericAscending(UserInput) {
	; Sort selected text and replace it	
	vText := UserInput
	Sort, vText, N ; Numeric Sort
	
	SB_SetText("Sorted Numeric Ascending.")
	return vText
}

TextSortNumericDescending(UserInput) {
	; Sort selected text and replace it	
	vText := UserInput
	Sort, vText, R N ; Reverse Numeric Sort
	
	SB_SetText("Sorted Numeric Descending.")
	return vText
}

TextSortNumericWithDelimiter(UserInput) {
	; Sort selected text and replace it
	vText := UserInput		
	Prompt = "Enter delimiter separating values"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	Length := StrLen(OutputVar)
	
	if ( Length > 0 ) {
		Sort, vText, N D%OutputVar% ; Numeric Sort
		
		SB_SetText("Sorted Delimiter Numeric Data.")
		return vText
	}
}

TextSortReverse(UserInput) {
	; Sort selected text and replace it
	vText := UserInput
	Sort, vText, R ; Sort and then Reverse
	
	SB_SetText("Reverse sorted.")
	return vText
}

TextSortReverseOrder(UserInput) {
	; Sort selected text and replace it
	
	vText := UserInput
	Sort, vText, F HelperReverseDirection ; Reverse order
	
	SB_SetText("Reverse ordered.")
	return vText
}

HelperReverseDirection(a1, a2, offset)
{
    return offset  ; Offset is positive if a2 came after a1 in the original list; negative otherwise.
}

TextSortReverseWithDelimiter(UserInput) {
	; Use AHK's standard sort to reverse sort values separated by a delimiter
	
	vText := UserInput
		
	Prompt = "Enter delimiter separating values"
	InputBox, OutputVar, Enter Data for ..., %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%, , %Timeout%, %Dft%
	
	Length := StrLen(OutputVar)
	
	if ( Length == 0 ) {
		SB_SetText("User did not enter a delimiter.")
		return
	}
	
	Sort, vText, R D%OutputVar% ; Sort and then Reverse
	
	SB_SetText("Sort reversed delimiter text.")
	return vText
}

TextSortRandom(UserInput) 
{
	; Use AHK's standard Sort function to randomly sort the text given - 
	
	vText := UserInput
	Sort, vText, Random
	
	SB_SetText("Text sorted randomly.")
	return vText
}
