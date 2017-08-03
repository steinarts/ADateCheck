$subroutine ADateCheck [Lock="N",Author="stsk",Written="2017-08-03 14:18:04",Text="checks if the date input is in the dateformat used"]
'should work with MM-MON,YY-YYYY and "",-
Function ADateCheck(strDate as String, strDateFormat as String) as Boolean
	Dim lngDayPos as long
	Dim lngMonthPos as long
	Dim lngYearPos as long
	Dim blnMM as Boolean
	Dim blnMON as Boolean
	Dim blnYYYY as Boolean
	Dim lngYear as long
	Dim strSeperator as String
	Dim blnRet as Boolean
	Dim strDay as String
	Dim strMonth as String

	'find seperator 
	if Instr(strDateFormat,".") > -1 then
		strSeperator = "."
	elseif Instr(strDateFormat,"-") > -1 then
		strSeperator = "-"
	Else 
		strSeperator = ""
	End If

	'check for correct seperator
	Dim lngIndex as Long
	If strSeperator <> "" then
		lngIndex = Instr(strDate,strSeperator)
		if lngIndex > -1 then
			If lngIndex <> StrLastIndexOf(strDate,strSeperator)-1 then
				blnRet = true
			Else
				blnRet = false
				Return
			End If
		Else
			Return
		End If
	Else
		If Instr(strDate,".") > -1 or Instr(strDate,"-") > -1 then
			Return
		End IF
	End If

	'find date (DD) position
	lngDayPos = Instr(strDateFormat,"DD")

	'check correct day
	lngDayPos = Instr(strDateFormat,"DD") 
	strDay = mid(strDate,lngDayPos, 2)
	If clng(strDay) > 0 and clng(StrDay) <32 then
		blnRet = true
	Else
		blnRet = false
		Return
	End IF

	'month: is it MM or MON
	If Instr(strDateFormat,"MM") > -1 then
		blnMM = true
		'blnMON = false
	elseIf Instr(strDateFormat,"MON") > -1 then
		blnMM = false
		'blnMON = true
	Else
		return
	End If

	'find month position
	If blnMM then
		lngMonthPos = Instr(strDateFormat,"MM")
	Else
		lngMonthPos = Instr(strDateFormat,"MON")
	End If

	'check correct month
	'find month position
	If blnMON then
		strMonth = mid(strDate,lngMonthPos, 3)
		If strMonth = "JAN" or strMonth = "FEB" OR strMonth = "MAR" or strMonth = "APR" or strMonth = "MAY"
		Or strMonth = "JUN" Or strMonth = "JUL" Or strMonth = "AUG" Or strMonth = "SEP"
		Or strMonth = "OKT" Or strMonth = "NOV" Or strMonth = "DEC" then
			blnRet = true
		Else
			blnRet = false
			Return
		End If
	ElseIf blnMM then
		strMonth = mid(strDate,lngMonthPos, 2)
		If clng(strMonth) > 0 and clng(strMonth) < 13 then
				blnRet = true
			Else
				blnRet = false
				Return
			End IF
	Else
		'error
	End If
	
	'is YYYY?
	If Instr(strDateFormat,"YYYY") > -1 then
		blnYYYY = true
	Else
		blnYYYY = false
	End If

	'find year position
	If blnYYYY then
		lngYearPos = Instr(strDateFormat,"YYYY")
	Else
		lngYearPos = Instr(strDateFormat,"YY")
	End If

	'check correct year
	if blnYYYY then
		lngYear = mid(strDate,lngYearPos, 4)
		If lngYear > 1900 and lngYear < 2200 then
			blnRet = true
		Else
			blnRet = false
			Return
		End IF
	Else
		lngYear = mid(strDate,lngYearPos, 2)
		If lngYear > 0 and lngYear < 99 then
			blnRet = true
		Else
			blnRet = false
			Return
		End IF

	End IF
	Return blnRet

End Function
$endobj

