; My AHK Text Tool Class
; Author: James Burns
;
; Purpose:
;  To Provide a simple to use class/library of methods for text processing.
; 
; Usage:
; vMyTextTool := New MyTextTool
; MsgBox % vMyTextTool.removeBlankLines(vText)
;
; Dependencies:
;  Join()
;
; Tools:
;  Convert - CR to CRLF, 
;  Format - Uppercase, Lowercase, Sentence case, Inverse case
;  Remove - CFLF, 
;  Text - Append (w/ w/o Newline), Prepend (w/ w/o NewLine)
;

#SingleInstance Force

; -----------------------------------------------------------------------------------------------------------
/*
; Join the items of an array into a text ilne, separated by custom separator
Join(Array, Sep)
{
	for k, v in Array
		out .= Sep . v
	return SubStr(out, 1+StrLen(Sep))
}
*/
; -----------------------------------------------------------------------------------------------------------


Class MyTextTool
{
	vNewLine := "`r`n"
	
	cannedTextLoremIpsum(){
		return "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
	}
	
	convertCharToNewLine(vText, vFind){
		; Find/Replace with Regex: Allow user to enter a regular expression and replacement string
		return StrReplace(vText, vFind, "`r`n") ; Symbols are free to use in StringReplace
	}
	
	convertCRtoCRLF(vText){
		; Turn `r to `r`n, in order to allow to work with Notepad and other Windows programs
		return RegExReplace(vText, "`r", "`r`n")
	}
	
	convertDateFromMMDDYYYY(vText){
		; Get the full date from mm/dd/yyyy
		; Ex: 1/1/2010 = Friday, January 1, 2010
		;datestring := "1/1/2010"

		FoundPos := RegExMatch(vText, "O)(\d+\/\d+\/\d+)", RegExMatchOutputVar)
		
		if ( RegExMatchOutputVar.Count() > 0 ) {
			SetFormat, float, 02.0
			StringSplit, d, strDate, / 
			FormatTime, FullDate, % d3 . d1+0. . d2+0., dddd, MMMM d, yyyy
			
			return FullDate
		}
		else
			return "Error"
	}
	
	convertDelimiter(vText, vOldDelimiter, vNewDelimiter){ ; Works with multiple lines of text
		TextArray := StrSplit(vText, vOldDelimiter)
		return Join(TextArray, vNewDelimiter) ; Dependency
	}
	
	convertLinesToListNested(vText){
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
		
		return SubStr(@, StrLen(NewLine) + 1)
	}
	
	convertToWindowsFilename(vText){
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
			@ .= tempText
		}

		return SubStr(@, StrLen(tempText) + 1)
	}
	
	convertLinesToListOrdinalLowercase(vText){
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
		
		return SubStr(@, StrLen(NewLine) + 1)
	}
	
	convertLinesToListOrdinalNumber(vText){
		; Line 1				->	1. Line 1
		; 	Indented Line 2		->		1.1. Indented Line 2
		;   Indented Line 3		->  	1.2. Indented Line 3
		;		Indented Line 4 ->			1.2.1. Indented Line 4
		
		; 1.) First organize lines into an array so we know how to number
		; 2.) Loop through array and prefix lines with proper numbering
		; 3.) Print results
		
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
	
	convertLinesToListUnorderedCustom(vText, vCustomChar){
		NewLine := "`r`n"
		vCustomChar := OutputVar
		OrdinalText := vCustomChar . A_Space
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
	
	convertLinesToListOrdinalUppercase(vText){
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
		
		return SubStr(@, StrLen(NewLine) + 1)
	}
	
	formatIndentLess(vText){
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

		return SubStr(@, StrLen(NewLine) + 1)
	}
	
	formatIndentMore(vText){
		; https://autohotkey.com/board/topic/70404-clip-send-and-retrieve-text-using-the-clipboard/
		; Indent a block of text. (This introduces a small delay for typing a normal tab; to avoid this you can use ^Tab / ^+Tab as a hotkey instead.)
		
		TabChar := A_Tab ; this could be something else, say, 4 spaces
		NewLine := "`r`n"
		
		@ := ""
		Loop, Parse, vText, `n, `r
			@ .= NewLine (InStr(A_ThisHotkey, "+") ? SubStr(A_LoopField, (InStr(A_LoopField, TabChar) = 1) * StrLen(TabChar) + 1) : TabChar A_LoopField)
		
		return SubStr(@, StrLen(NewLine) + 1)
	}
	
	formatInversecase(vText){
		; Thanks JDN ; Mango.ahk
		Lab_Invert_Char_Out:= ""
		Loop % Strlen(vText) {
			Lab_Invert_Char:= Substr(vText, A_Index, 1)
			If Lab_Invert_Char is Upper
				Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) + 32)
			Else If Lab_Invert_Char is Lower
				Lab_Invert_Char_Out:= Lab_Invert_Char_Out Chr(Asc(Lab_Invert_Char) - 32)
			Else
				Lab_Invert_Char_Out:= Lab_Invert_Char_Out Lab_Invert_Char
		}
		
		return Lab_Invert_Char_Out
	}
	
	formatLowercase(vText){
		StringLower, vText, vText
		return vText
	}
	
	formatSentencecase(vText){
		; SENTENCE CASE - quick brown fox. went home. = Quick brown fox. Went home.
		return RegExReplace(vText, "([.?\s!]\s\w)|^(\.\s\b\w)|^(.)", "$U0")
	}
	
	formatStrikethrough(vText){
		; Strikeout with Unicode('\u0336') - borrowed from Text Tools (https://chrome.google.com/webstore/detail/text-tools/mpcpnbklkemjinipimjcbgjijefholkd)
		vSplit := StrSplit(vText, "")
		for k,v in vSplit
			vSplit[k] := v . "?"
		
		return Join(vSplit, "")
	}
	
	formatTelephoneNumber(vText){
		; Formats the Clipboard value to (###) ###-#### format
		; Ex: 1234567890 --> (123) 456-7890
		vFormatted := RegExReplace(vText, "(\d{1})(\d{3})(\d{3})(\d{3})", "$1-($2) $3-$4")
		vFormatted := RegExReplace(vFormatted, "(\d{3})(\d{3})(\d{3})", "($1) $2-$3")
		
		return vFormatted
	}
	
	formatTitlecase(vText){
		StringUpper, vText, vText, T	; "T" for Title Case
		return vText
	}
	
	formatUnderline(vText){
		; Underscore with Unicode ('\u0332') - borrowed from TextTools (https://chrome.google.com/webstore/detail/text-tools/mpcpnbklkemjinipimjcbgjijefholkd)
		vSplit := StrSplit(vText, "")
		for k,v in vSplit
			vSplit[k] := v . "_"
		return Join(vSplit, "") ; Dependency
	}
	
	formatUppercase(vText){
		StringUpper, vText, vText
		return vText
	}
	
	reduceWhitespace(vText){
		return RegExReplace(vText, " {2,}", " ")
	}
	
	removeBlankLines(vText){
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
		
		return vText
	}
	
	removeCRLF(vText){
		vText := RegExReplace(vText, "`r`n", "") ; Remove CRLF and replace with nothing
		vText := RegExReplace(vText, "`n", "")   ; Remove CR
		vText := RegExReplace(vText, "`r", "")   ; Remove LF
		
		return vText
	}
	
	removeWhitespace(vText){
		return StrReplace(vText, " ", "")
	}
	
	textAppend(vText, vAppendText, vNewline:=""){
		if (vNewLine)
			return (vText . "`r`n" . vAppendText)
		else
			return (vText . vAppendText)
	}
	
	textJoinLinesWithChar(vText, vJoinChar){
		vResults := 
	
		Loop, Parse, vText, `n, `r 
			vResults .= vJoinChar A_LoopField
		
		if (StrLen(varJoiner) > 0)
			vResults := Substr(vResults, StrLen(vJoinChar) + 1) ; Remove excess "varJoiner" from the beginning of the first line
		
		return vResults
	}
	
	textPrepend(vText, vPrependText, vNewLine:=""){
		if (vNewLine)
			return (vPrependText . "`r`n" . vText)
		else
			return (vPrependText . vText)
	}
	
	textTrim(vText, vSide:=""){
		if ( !vSide )
			return Trim(vText)
		else if ( vSide == "Left" or vSide == "left" )
			return LTrim(vText)
		else
			return RTrim(vText)
	}
	
}