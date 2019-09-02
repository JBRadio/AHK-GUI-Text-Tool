; Text Manipulation for Programmer
;
; Purpose: To provide ready-to-use text scultping tools related to Programmers data/information in efforts to increase productivity, visibilty, and general use
;

; Function List
; -------------
; ** OPEN URL Considerations
; 	* Development material
;	* Troubleshooting material
;	* Social (Google Group, ...)
; 
; Browser
;  Open browser to temporary html code
;	Ex: C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe "data:text/html,<textarea style='font-size: 1.5em; width: 100%; height: 100%; border: none; outline: none;' placeholder='Type here' autofocus />"
; CSS
; JavaScript
; JSON
; HTML
; Windows
;  Command Prompt
;  Windows Run

;							TEMPLATE
; -----------------------------------------------------------------------------------------------------------
; Description
; Input: 
; Output: 
FunctionName(UserInput)
{
	vResult :=
	SB_SetText("Status Message")
	return vResult
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
; Browser HTML Hack - Turn a browser tab into a text pad
BrowserOpenTabToHtmlNotepad()
{
	vProgram := "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"
	vArgument := """data:text/html,<textarea style='font-size: 1.5em; width: 100%; height: 100%; border: none; outline: none;' placeholder='Type here' autofocus />"""
	Run % vProgram A_Space vArgument
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
BrowserOpenTabToHtml(UserInput)
{
	;vTextHtml := "<table><thead><tr><th colspan=""2"">The table header</th></tr></thead><tbody><tr><td>The table body</td><td>with two columns</td></tr></tbody></table>"
	vTextHtml := TextFlattenMultiLine(UserInput)
	vTextHtml := HelperTextJoinLinesWithChar(vTextHtml, "")	; (vText, vJoinChar)
	vProgram := "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"
	vArgument := """data:text/html,<REPLACEME>"""
	vArgument := RegExReplace(vArgument, "<REPLACEME>", vTextHtml)
	
	Run % vProgram A_Space vArgument
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
CommandPromptSystemInfo()
{
	Command := "systeminfo"
	HelperCommandPromptClip(Command, 0)
	;SetStatusMessageAndColor("Ran systeminfo in Command Prompt.", "Green")
	SB_SetText("Ran systeminfo in Command Prompt.")
	return
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
CommandPromptTracert(UserInput)
{
	Command := "tracert " . UserInput
	HelperCommandPromptClip(Command, 0)
	;SetStatusMessageAndColor("Ran tracert in Command Prompt.", "Green")
	SB_SetText("Ran tracert in Command Prompt.")
	return
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
CommandPromptPing(UserInput)
{
	; Command Prompt: Ping
	Command := "ping " . UserInput
	HelperCommandPromptClip(Command, 0)
	;SetStatusMessageAndColor("Ran ping in Command Prompt.", "Green")
	SB_SetText("Ran ping in Command Prompt.")
	return
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
CommandPromptGetIpconfigAll()
{
	; Command Prompt: IPCONFIG info
	Command := "ipconfig /all"
	HelperCommandPromptClip(Command, 1)
	;MsgBox % clipboard
	;SetStatusMessageAndColor("Ran Ipconfig in Command Prompt.", "Green")
	SB_SetText("Ran Ipconfig in Command Prompt.")
	return clipboard
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
CommandPromptGetNsLookup(UserInput)
{
	Command := "nslookup " . UserInput
	HelperCommandPromptClip(Command, 1)
	;SetStatusMessageAndColor("Ran NSLookup in Command Prompt.", "Green")
	SB_SetText("Ran NSLookup in Command Prompt.")
	return clipboard
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
	arrCsvData := HelperGetSVDataIntoRowsArray(data, delim) ; From General Model

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
	;SetStatusMessageAndColor("Converted CSV to <table>.", "Green")
	SB_SetText("Converted CSV to <table>.")
	return htmlTable
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
EncodeToAscii(UserInput)
{
	Transform, Text, ASc, %UserInput% ; Encode ASCII
	;SetStatusMessageAndColor("Encoded input (ASCII).", "Green")
	SB_SetText("Encoded input (ASCII).")
	return Text
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
EncodeToHtml(UserInput)
{
	Transform, Text, HTML, %UserInput% ; Encode HTML
	;SetStatusMessageAndColor("Encoded input (HTML).", "Green")
	SB_SetText("Encoded input (HTML).")
	return Text
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
EncodeToCharCode(UserInput) {
	
	Text := UserInput
	
	if Text is not Integer
	{
		;MsgBox % "User input must be in the form of an integer."
		;SetStatusMessageAndColor("User input must be in the form of an integer. (EncodeToCharCode)", "Red")
		SB_SetText("User input must be in the form of an integer. (EncodeToCharCode)")
		return
	}
		
	Chr := Chr(Text)
	;SetStatusMessageAndColor("Encoded input (CharCode).", "Green")
	SB_SetText("Encoded input (CharCode).")
	return Chr
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
uriDecode(str) {
;https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/
	Loop
		If RegExMatch(str, "i)(?<=%)[\da-f]{1,2}", hex)
			StringReplace, str, str, `%%hex%, % Chr("0x" . hex), All
		Else Break
	
	;SetStatusMessageAndColor("Decoded input (URI).", "Green")
	SB_SetText("Decoded input (URI).", "Green")
	Return, str
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
UriEncode(Uri, RE="[0-9A-Za-z]"){
;https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/
	VarSetCapacity(Var,StrPut(Uri,"UTF-8"),0),StrPut(Uri,&Var,"UTF-8")
	While Code:=NumGet(Var,A_Index-1,"UChar")
		Res.=(Chr:=Chr(Code))~=RE?Chr:Format("%{:02X}",Code)
	
	;SetStatusMessageAndColor("Encoded input (URI).", "Green")
	SB_SetText("Encoded input (URI).")
	Return,Res  
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
EncodeToUrl(UserInput)
{
	Text := UserInput
	Text := uriEncode(Text) ; Encode URI
	
	;SetStatusMessageAndColor("Encoded input (URL).", "Green")
	SB_SetText("Encoded input (URL).")
	return Text
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
DecodeUrl(UserInput)
{
	Text := UserInput
	Text := uriDecode(Text) ; Decode URI
	
	;SetStatusMessageAndColor("Decoded input (URL).", "Green")
	SB_SetText("Decoded input (URL).")
	return Text
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
DecodeUrlAndGetParams(UserInput) {
	;Decode URL and Wrap on parameters
	;http://the-automator.com/parse-url-parameters/
	
	Text := UserInput
	Text := uriDecode(Text) 
	StringReplace,Text,Text,?,`r`n`t?,All ;Line break and tab indent <strong>parse URL parameters</strong>.
	StringReplace,Text,Text,&,`r`n`t`t&,All ;Line break and double tab indent
	
	;SetStatusMessageAndColor("Decoded input and got params (URL).", "Green")
	SB_SetText("Decoded input and got params (URL).")
	return Text
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
			
			;SetStatusMessageAndColor("Wrapped Text with HTML tag.", "Green")
			SB_SetText("Wrapped Text with HTML tag.")
			return retValue
		}
	;}
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
			
			;SetStatusMessageAndColor("HTML Wrapped each line.", "Green")
			SB_SetText("HTML Wrapped each line.")
			return retValue
		}
	;}
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
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
		
		;SetStatusMessageAndColor("Downloaded to text file.", "Green")
		SB_SetText("Downloaded to text file.")
		return vFileReadOutput
	}
	catch e
	{
		;MsgBox % "There was an error with URLDownloadToFile"
		MsgBox, 16,, % "Exception thrown!`n`nwhat: " e.what "`nfile: " e.file . "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra
		
		return
	}
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
WindowsRun(UserInput){
	Commands := StrSplit(UserInput, ["`r`n", "`n", "`r"])
	;MsgBox % Join(Commands, ", ")
	;MsgBox % Commands[1]
	Command := Commands[1]
	Run, %Command%
	;SetStatusMessageAndColor("Sent text to Windows Run.", "Green")
	SB_SetText("Sent text to Windows Run.")
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
TextFormatJson(UserInput)
{	
	;vFormatted := JSON_Beautify(UserInput)
	vFormatted := JSON_Beautify(UserInput, "  ")
	;MsgBox % vFormatted
	
	;SetStatusMessageAndColor("Formatted (JSON).", "Green")
	SB_SetText("Formatted (JSON).")
	return vFormatted
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
HtmlTableToCSV(UserInput){
	; 1.) Flatten Each Line - remove indents
	; 2.) Join each line with no character - line up all 
	; 3.) Match and replace
	;     <tr><th> -> ""
	;     </th><th> -> ","
	;     </th></tr> -> ""
	;
	;     <tr><td> -> "\n"
	;     </td><td> -> ","
	;     </td></tr> -> ""
	;
	;  - Since we need the text to be one continuous line, we may have issues with <pre> content, where spaces are supposed to exist.
	;  - If colspan is used, data may need to be duplicated in order to match column data
	
	vTextHtml := TextFlattenMultiLine(UserInput)
	vTextHtml := HelperTextJoinLinesWithChar(vTextHtml, "")	; (vText, vJoinChar)
	;vTextHtml := "<table><thead><tr><th colspan=""2"">The table header</th></tr></thead><tbody><tr><td>The table body</td><td>with two columns</td></tr></tbody></table>"
	vResults :=

	RegExMatch(vTextHtml, "O)<thead>(<tr>.*</tr>)</thead>", vTableHeadFields)
	if ( vTableHeadFields.Count() > 0 )
	{
		vTemp := vTableHeadFields.Value(1)
		vTemp := RegExReplace(vTemp, "<tr><th.*?>", "")
		vTemp := RegExReplace(vTemp, "</th><th.*?>", ",")
		vTemp := RegExReplace(vTemp, "</th></tr>", "")
		vResults .= vTemp
	}

	RegExMatch(vTextHtml, "O)<tbody>(<tr>.*</tr>)</tbody>", vTableBodyRows)
	if ( vTableBodyRows.Count() > 0)
	{
		vTemp := vTableBodyRows.Value(1)
		vTemp := RegExReplace(vTemp, "<tr><td>", "`n")
		vTemp := RegExReplace(vTemp, "</td><td.*?>", ",")
		vTemp := RegExReplace(vTemp, "</td></tr>", "")
		vResults .= vTemp
	}

	return vResults
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------
FileGetProperties(){
	SetBatchLines, -1							; Originally 10ms - SetBatchLines, 10ms
	FileSelectFile, FilePath					; Select a file to use for this example.
	
	if (FilePath = "") {
		SB_SetText("The user didn't select a file.")
		return
	}

	/*
	PropName := FGP_Name(0)						; Gets a property name based on the property number.
	PropNum  := FGP_Num("Size")					; Gets a property number based on the property name.
	PropVal1 := FGP_Value(FilePath, PropName)	; Gets a file property value by name.
	PropVal2 := FGP_Value(FilePath, PropNum)	; Gets a file property value by number.
	PropList := FGP_List(FilePath)				; Gets all of a file's non-blank properties.

	MsgBox, % FilePath
	. "`n" PropName ":`t" PropVal1			; Display the results.
	. "`n" PropNum ":`t" PropVal2
	. "`n`nList:`n" PropList.CSV
	*/

	PropName := FGP_Name(0)						; Gets a property name based on the property number.
	PropVal1 := FGP_Value(FilePath, PropName)	; Gets a file property value by name.
	PropList := FGP_List(FilePath)				; Gets all of a file's non-blank properties.

	vResult := PropName ": " PropVal1	; "Name: " {Name of Selected File}
	. "`nPath: " FilePath					; "Path: " {Path of Selected File}
	. "`n`nProperties:`n" PropList.CSV

	SetBatchLines, 10ms
	return vResult
}
; -----------------------------------------------------------------------------------------------------------

; -----------------------------------------------------------------------------------------------------------

;GUIName := "MyGui"
/*
ShowListOfWindowsExtendedProperties(){
	
	;global GUIName
	
	SetBatchLines, -1
	Gui, MyGui: New
	Gui, MyGui: Margin, 20, 20
	Gui, MyGui: Add, ListView, w600 r20 Grid vLV +E0x010000, #|Index|Name
	
	Details := GetDetails()
	For Index, Name In Details
	   LV_Add("", A_Index, Index, Name)
	Loop, % LV_GetCount("Column")
	   LV_ModifyCol(A_Index, "AutoHdr")
	Gui, %GUIName%:Show, , Extended Properties
	
	SetBatchLines, 10ms
	SB_SetText("Launched Gui")
	Return
}

GetDetails() {
   Static MaxGap := 11 ; on WIn 8.1 MU all gaps between named details are less than 4
   Shell := ComObjCreate("Shell.Application")
   Folder := Shell.NameSpace(0)
   Details := [], Gap := 0
   While (Gap < MaxGap)
	  If (Name := Folder.GetDetailsOf(0, A_Index - 1))
		 Details[A_Index - 1] := Name, Gap := 0
	  Else
		 Gap++
   Return Details
}
*/
