strInput = InputBox("Please enter a valid landline phone number," & vbNewLine & "mobile phone number, email address" & vbNewLine & "or National Insurance Number.", "Contact Details") 

OutputType = validateInput(strInput)

FormattedOutput = FormatPhoneNumber(strInput, OutputType)


if OutputType = "Landline" then
	MsgBox("You have entered a valid landline number")
	
	elseif OutputType = "Mobile" then
	MsgBox("You have entered a valid mobile number")
	
	elseif OutputType = "Email" then
	MsgBox("You have entered a valid email address")
	
	elseif OutputType = "NINumber" then
	MsgBox("You have entered a valid NI Number")
	
	elseif OutputType = "Invalid" then
	MsgBox("Invalid Input. Please input valid details")
	
	Else
	MsgBox("An unknown error has occoured. Please contact IT")

End if

Function validateInput(strInput)

Set objRegPhone = New RegExp
    objRegPhone.IgnoreCase = True
    objRegPhone.Global = True
    objRegPhone.Pattern = "^(((\+44\s?\d{4}|\(?0\d{4}\)?)\s?\d{3}\s?\d{3})|((\+44\s?\d{3}|\(?0\d{3}\)?)\s?\d{3}\s?\d{4})|((\+44\s?\d{2}|\(?0\d{2}\)?)\s?\d{4}\s?\d{4}))(\s?\#(\d{4}|\d{3}))?$"
		  
Set objRegEmail = New RegExp
    objRegEmail.IgnoreCase = True
    objRegEmail.Global = True
    objRegEmail.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"


Set objRegNINumber = New RegExp
	objRegNINumber.IgnoreCase = True
	objRegNINumber.Global = True
	objRegNINumber.Pattern = "^\s*([a-zA-Z]){2}\s*([0-9]){1}\s*([0-9]){1}\s*([0-9]){1}\s*([0-9]){1}\s*([0-9]){1}\s*([0-9]){1}\s*([a-zA-Z]){1}?$"
		  

		  
if left(strInput, 4) = "+447" Then
	If objRegPhone.Test(strInput) then
		validateInput = "Mobile"
		Else 
	End IF 
	
	elseif left(strInput, 2) = "07" Then
		If objRegPhone.Test(strInput) then
		validateInput = "Mobile"
		Else 
	End IF 

	elseif objRegPhone.Test(strInput) then
	validateInput = "Landline"
	
	elseif objRegEmail.Test(strInput) then
	validateInput = "Email"
	
	elseif objRegNINumber.Test(strInput) then
	validateInput = "NINumber"

	Else
	validateInput = "Invalid"
	
End if

Set objRegPhone = nothing
Set objRegEmail = nothing
Set objRegNINumber = nothing
End Function

Function FormatPhoneNumber(strInput, OutputType)

RequiresModification = False
NumberModified = False
strInput = Replace(strInput,"+44","0")	
strInput = Replace(strInput, " ","")

If OutputType = "Landline" Then 
	RequiresModification = True
	
	ElseIF  OutputType = "Mobile" Then 
	RequiresModification = True
	
	Else RequiresModification = False
	
End IF

IF RequiresModification = True Then 
	Select Case left(strInput, 3) 
		Case 020, 023, 024, 028, 029
			FormatPhoneNumber = Mid(strInput,1,3) & " " & Mid(strInput,4)			
			NumberModified = True
			MsgBox(FormatPhoneNumber)
	End Select 
	
	Select Case left(strInput, 4)
		Case 0113, 0114, 0115, 0116, 0117, 0118, 0121, 0131, 0141, 0151, 0161
			FormatPhoneNumber = Mid(strInput,1,4) & " " & Mid(strInput,5)		
			NumberModified = True			
			MsgBox(NumberModified)
	End Select
	
	If NumberModified = False Then 
		FormatPhoneNumber = Mid(strInput,1,5) & " " & Mid(strInput,6)		
		MsgBox("Bilbo")
		Else
	End IF
	
	Else 
End IF

End Function