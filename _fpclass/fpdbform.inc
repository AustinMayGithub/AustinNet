<%
Function FP_SaveFormFields(rs, rgFormFields, rgDBFields)
	On Error Resume Next
	Err.Clear

	Dim i
	For i = 0 To UBound(rgFormFields)
		FP_SaveFormField rs, rgFormFields(i), rgDBFields(i)
	Next 
	Err.Clear
End Function

Function FP_SaveFormField(rs, strField, strDBField)
	On Error Resume Next
	Err.Clear

	If (Request.Form(strField)) = "" And rs(strDBField).Type <> adBoolean Then
		Exit Function
	End If

	FP_SaveFieldToDB rs, Request.Form(strField), strDBField	
	Err.Clear
End Function

Function FP_SaveFieldToDB(rs, strField, strDBField)
	On Error Resume Next
	Err.Clear

	Select Case rs(strDBField).Type
		Case adInteger Or adBigInt Or adUnsignedTinyInt Or adUnsignedSmallInt Or  adUnsignedInt Or adUnsignedBigInt
			rs(strDBField) = CInt(strField)
		Case adSingle Or adDecimal Or adNumeric 
			rs(strDBField) = CSng(strField)
		Case adDouble
			rs(strDBField) = CDbl(strField)
		Case adCurrency
			rs(strDBField) = CCur(strField)
		Case adBoolean
			rs(strDBField) = CBool(strField)
		Case adDate Or adDBDate Or adDBTime or adDBTimeStamp
			rs(strDBField) = CDate(strField)
		Case Else
			rs(strDBField) = CStr(strField)
	End Select

	strError = "Cannot save value """ & strField & """ to database field """ & strDBField & """"

	FP_DumpError strErrorUrl, strError 
	Err.Clear
End Function

Function FP_EncodeOutput(str)
	FP_EncodeOutput = str
	FP_EncodeOutput = Replace(FP_EncodeOutput,"&","^^@^^")
	FP_EncodeOutput = Server.HTMLEncode(FP_EncodeOutput)
	FP_EncodeOutput = Replace(FP_EncodeOutput,"^^@^^","&")
End Function

Function FP_FormConfirmation(szCharset, szTitle, szmsg1, szUrl, szMsg2)
	On Error Resume Next
	Response.Write "<html>" & vbCrLf
	Response.Write "<head>" & vbCrLf
	Response.Write "<meta http-equiv=""Content-Type"" content=""" & szCharSet & """>" & vbCrLf
	Response.Write "<title>" & szTitle & "</title>" & vbCrLf
	Response.Write "</head>" & vbCrLf
	Response.Write "<body>" & vbCrLf
	Response.Write "<h1>" & szTitle & "</h1>" & vbCrLf
	Response.Write "<p>" & szMsg1 &"</p>" & vbCrLf
	Response.Write "<p>" & vbCrLf
	For Each item in Request.Form 
		If item <> "VTI-GROUP" Then
			Response.Write "<b>" & item & "</b>" & ": " & FP_EncodeOutput(Request.Form(item)) & "<BR>" & vbCrLf
		End If
	Next 
	Response.Write "</p>" & vbCrLf
	Response.Write "<p><a href=""" &  szUrl & """>" & szMsg2 & "</a></p>" & vbCrLf
	Response.Write "</body>" & vbCrLf
	Response.Write "</html>" & vbCrLf
	Session.CodePage = Session("FP_OldCodePage")
	Session.LCID = Session("FP_OldLCID")
	Response.End
End Function

Function FP_FormConfirmationFromArrays(szCharset, szTitle, szmsg1, szUrl, szMsg2, rgDBFields, rgDBValues)
	On Error Resume Next
	Response.Write "<html>" & vbCrLf
	Response.Write "<head>" & vbCrLf
	Response.Write "<meta http-equiv=""Content-Type"" content=""" & szCharSet & """>" & vbCrLf
	Response.Write "<title>" & szTitle & "</title>" & vbCrLf
	Response.Write "</head>" & vbCrLf
	Response.Write "<body>" & vbCrLf
	Response.Write "<h1>" & szTitle & "</h1>" & vbCrLf
	Response.Write "<p>" & szMsg1 &"</p>" & vbCrLf
	Response.Write "<p>" & vbCrLf
	For j = 0 To UBound(rgDBFields) - 1
		Response.Write "<b>" & rgDBFields(j) & "</b>" & ": " & FP_EncodeOutput(rgDBValues(j)) & "<BR>" & vbCrLf
	Next 
	Response.Write "</p>" & vbCrLf
	Response.Write "<p><a href=""" &  szUrl & """>" & szMsg2 & "</a></p>" & vbCrLf
	Response.Write "</body>" & vbCrLf
	Response.Write "</html>" & vbCrLf
	Session.CodePage = Session("FP_OldCodePage")
	Session.LCID = Session("FP_OldLCID")
	Response.End
End Function

Function FP_DumpError(strErrorUrl, strMsg)
	On Error Resume Next
	If Err.Number <> 0 Then
		fp_conn.Close
		Session.CodePage = Session("FP_OldCodePage")
		Session.LCID = Session("FP_OldLCID")
		If strErrorUrl <> "" Then
			Response.Redirect strErrorUrl
		Else
			Response.Write "<b> " & strMsg & "</b><p>"
			Response.Write "Error Description: " & Err.Description & "<p>"
			Response.Write "Error Number: " & Err.Number  & "<p>"
			Response.Write "Error Source: " & Err.Source & "<p>"
			Response.End
		End If
	End if
End Function
%>
