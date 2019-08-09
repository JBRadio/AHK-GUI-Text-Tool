; MyPrintTool		Helpful Printer/Formatting Tool
; Author James Burns

Class MyPrintTool{
	
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

	HelperPrintSpacedColumns(arrData, hasHeader)
	{
		colMaxWidth := this.HelperFindColumnMaxWidth(arrData)
		
		colBetweenSpace := 2
		@ :=
		NewLine := "`r`n"

		Loop, % arrData.MaxIndex()
		{ ; Rows
			rIndex := A_Index ; row Index
			@ .= NewLine
			
			if ( hasHeader and A_Index == 2 )
			{
				; Print ----- in column spot to separate headers from body data
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
						@ .= MyPrintTool.HelperPadSpaces(vLength, vText, vDirection)
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
					@ .= MyPrintTool.HelperPadSpaces(vLength, vText, vDirection)
				}
			}
		}

		return SubStr(@, StrLen(NewLine)+1)
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
			@ .= this.HelperPadSpaces(vLength, A_LoopField, vDirection)
		}
		
		return SubStr(@, StrLen(NewLine) +1)
	}
	
	getTextFor1DArray(data){
		Local index, value, result
		
		for index, value in data
			result .= "Index " index ": " value "`r`n"
		return result
	}
	
	getTextFor2DArray(data){
		Local rIndex, row
		Local cIndex, column
		Local result
		
		for rIndex, row in data
			for cIndex, column in data[rIndex]
				result .= "Row: " rIndex ", Column: " column "`n"
		return result
	}
	
	print1DArrayToMsgBox(data){		
		MsgBox % this.getTextFor1DArray(data)
	}
	
	print2DArrayToMsgBox(data){
		MsgBox % this.getTextFor2DArray(data)
	}
}