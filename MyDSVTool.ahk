; MyDSVTool Delimiter Separated Value Parser
; author: James Burns

#SingleInstance Force

#Include MyPrintTool.ahk

Class MyDSVTool 
{
	isLoaded := false				; Used to determine if we should perform a function other than loadText() / loadFile() / loadURL()
	dataLoaded := 					; Data to be used for subsequent functions after loadText() / loadFile() / loadURL ()
	delimiterLoaded :=				; Delimiter used with the loaded data, in case different than a comma (,)
	defaultDelimiter := ","			; Default delimiter to comma (,) if none is provided
	processedResult :=				; Data result from a process on the loaded data
	NewLine := "`r`n"				; Carriage Return Line Feed
	
	; Getters / Setters
	clearDataLoaded(){
		this.dataLoaded :=
		this.isLoaded := false
	}
	
	getTotalRows(){
		if (this.isLoaded)
			return this.dataLoaded.MaxIndex() 
		else
			return 0
	}
	
	getTotalColumns(){
		if (this.isLoaded)
			return this.dataLoaded[1].MaxIndex()
		else
			return 0
	}
	
	getData(){
		if (this.isLoaded)
			return this.dataLoaded
		else
			return 0
	}
	
	getDelimiter(){
		this.isLoaded ? return this.delimiterLoaded : return this.defaultDelimiter
	}
	
	getResult(){
		return this.processedResult ? this.processedResult : 0
	}
	
	setResult(value){
		this.processedResult := value
	}
	
	loadText(data:="",delim:="") {
		
		Local Rows
		Local Columns
		
		if ( !data or data == "" ) {
			MsgBox % "No data was provided. Cannot load data"
			return
		}
		
		if ( !delim or delim == "" )
			delim := this.defaultDelimiter
		
		Rows := Object()
	
		Loop, Parse, data, `n, `r
		{
			if ( delim == "\t" or delim == "`t" )
				Columns := StrSplit(A_LoopField, "`t")
			else
				Columns := StrSplit(A_LoopField, delim)
			
			if ( Columns.Length() == 0 or Columns == )
				continue ; Even a non-delimiter value returns something so return if somehow we have nothing.
			
			Rows[A_Index] := Columns
		}
		
		this.dataLoaded := Rows
		this.delimiterLoaded := delim
		this.isLoaded := true
	}
	
	; Methods
	
	extractColumnData(column:=1)
	{
		Local data, rIndex, row, result
		
		; Validation
		if ( column < 1 or column > this.getTotalColumns() )
			column := 1
		
		; Processing
		data := this.getData()
		for rIndex, row in data
        {
			result .= this.NewLine row[column]
        }
		result := SubStr(result, StrLen(this.NewLine) + 1)
		
		; Return
		this.setResult(result)
	}
	
	swapColumnData(columnMove, columnSwap)
	{
		Local Rows, rIndex, row
		Local cIndex, column
		Local result, delim
		
		; Validation
		if ( !this.isLoaded )
			return
		
		if ( columnMove < 1 )
			columnMove := 1
		
		if ( columnMove > this.getTotalColumns() )
			columnMove := this.getTotalColumns()
		
		if ( columnSwap < 1 )
			columnSwap := 1
		
		if ( columnSwap > this.getTotalColumns() )
			columnSwap := this.getTotalColumns()
		
		if ( columnMove == columnSwap ) {
			this.setResult(0)
			return ; Cannot move column to its original position
		}
		
		; Processing
		Rows := this.getData()
		delim := this.delimiterLoaded
		
		for rIndex, row in Rows
		{
			result .= this.NewLine 
			for cIndex, column in Rows[rIndex]
			{
				if (cIndex == columnMove )
					result .= row[columnSwap] . delim
				else if (cIndex == columnSwap)
					result .= row[columnMove] . delim
				else 
					result .= row[cIndex] . delim
			}
			result := Substr(result, 1, StrLen(result) - StrLen(delim)) ; Remove last delimiter added
		}
		
		result := SubStr(result, StrLen(this.NewLine) + 1) ; Remove initial NewLine
		
		; Return
		this.setResult(result)
	}
	
	filterColumnData(columnFilter:=1)
	{
		Local Rows, rIndex, row
		Local cIndex, column
		Local result, delim
		
		; Validation
		if ( !this.isLoaded )
		{
			MsgBox % "First load DSV data."
			this.setResult(0)
			return
		}
		
		if ( this.getTotalColumns() == 1 )
		{
			MsgBox % "Cannot filter only remaining column."
			this.setResult(0)
			return
		}
		
		if ( columnFilter < 1 )
			columnFilter := 1
		
		if (columnFilter > this.getTotalColumns() )
			columnFilter := this.getTotalColumns()
		
		; Processing
		Rows := this.getData()
		delim := this.delimiterLoaded
		
		for rIndex, row in Rows
		{
			result .= this.NewLine 
			for cIndex, column in Rows[rIndex]
			{
				if (cIndex == columnFilter)
					continue ; skip or filter out column from result data
				else
					result .= row[cIndex] . delim
			}
			result := Substr(result, 1, StrLen(result) - StrLen(delim)) ; Remove last delimiter added
		}
		result := SubStr(result, StrLen(this.NewLine) + 1) ; Remove initial NewLine
		
		; Return
		this.setResult(result)
	}
	
	filterRowData(rowFilter:=1)
	{
		Local Rows, rIndex, row
		Local cIndex, column
		Local result, delim
		
		; Validation
		if ( !this.isLoaded )
		{
			MsgBox % "First load DSV data."
			this.setResult(0)
			return
		}
		
		if ( this.getTotalRows() == 1 )
		{
			MsgBox % "Cannot filter only remaining row."
			this.setResult(0)
			return
		}
		
		if ( rowFilter < 1 )
			rowFilter := 1
		
		if (rowFilter > this.getTotalRows() )
			rowFilter := this.getTotalRows()
		
		; Processing
		Rows := this.getData()
		delim := this.delimiterLoaded
		
		for rIndex, row in Rows
		{
			if (rIndex == rowFilter)
				continue ; skip or filter out column from result data
			else {
				result .= this.NewLine 
				for cIndex, column in Rows[rIndex]
						result .= row[cIndex] . delim
				result := Substr(result, 1, StrLen(result) - StrLen(delim)) ; Remove last delimiter added
			}
		}
		result := SubStr(result, StrLen(this.NewLine) + 1) ; Remove initial NewLine
		
		; Return
		this.setResult(result)
	}
	
	printAsSpacedColumns(hasHeader := false){
		Local data
		
		; Validation
		if (!this.isLoaded)
			return
		
		;Processing & Return
		data := this.getData()
		
		if ( hasHeader )
			this.setResult( MyPrintTool.HelperPrintSpacedColumns(data, true) )
		else
			this.setResult( MyPrintTool.HelperPrintSpacedColumns(data, false) )
	}
}

;myPrintTool := New MyPrintTool

;myData := [1,2,3]
;myData := "1,2,3`n4,5,6`n7,8,9`napple,banana,cookies"
;myDSVTool := New MyDSVTool
;myDSVTool.loadText(myData)
;myPrintTool.print1DArrayToMsgBox(myData)
;myPrintTool.print2DArrayToMsgBox(myDSVTool.dataLoaded)

;myDSVTool.extractColumnData(1)
;myDSVTool.SwapColumnData(1, 2) ; (From, To)
;myDSVTool.filterColumnData(4) ; (Column to filter out)
;myDSVTool.filterRowData() ; (Row to filter out)
;myDSVTool.printAsSpacedColumns(false)
;myDSVTool.printAsSpacedColumns(true)
;MsgBox % myDSVTool.getResult()
;Clipboard := myDSVTool.getResult()

;MsgBox % "Total Rows in Loaded Data: " myDSVTool.getTotalRows()
;MsgBox % "Total Columns in Loaded Data: " myDSVTool.getTotalColumns()

;DEBUG
;myDSVTool.print2DArrayToMsgBox(myDSVTool.dataLoaded)
