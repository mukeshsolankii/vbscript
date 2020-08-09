'This script create User Account in AD in correct OU with the all the attributes.

'What are we doing in this script?
' 1. We are getting data from excel sheet like eg. Unid , First name etc.
' 2. We are checking if the data is blank or not.
' 3. We are checking if the user account with that Unid is already exist in AD or not. (UserDN)
' 4. if User Account not already exist in AD we will than proceed further.
' 5. We will check the Location whether the user is from NA , EU or AS and we will call the function accordingly.
'    like NA_CreateUser() , EU_CreateUser() or AS_CreateUser().
' 6. According to the location we will create the user in sub OU using arrays.

'****************************************************'

'Variables.
Dim Unid , SAP_num , FN , LN , Password , Location , Task_num , Start_date , UserFound , strUserDN , Types , Manual


'Constants.
Const ADS_NAME_INITTYPE_GC = 3
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_1779 = 1
Const ADS_PROPERTY_APPEND = 3


'Starting from Here.
wscript.echo "Start"

'Getting data from excel sheet.
path = "C:\ScriptingSpace\Sol New Hire\SolutionRecords.xlsx"
Set objexl = CreateObject("Excel.Application")
Set owb = objexl.workbooks.open(path)
Set sheet = owb.worksheets("sheet1")


'Assigning value to variables.
row = 2

'Checking if the data from excel sheet is blank or not.
for i = 1 to 9 :
	if Trim(sheet.Cells(row, i).value = "") Then
		wscript.echo "Value of some column is empty in excel sheet!"
		owb.close
		objexl.quit
		wscript.quit
	End if
next

Unid = Trim(sheet.Cells(row,2).value)
Types = Trim(sheet.Cells(row,5).value)
SAP_num = Trim(sheet.Cells(row,1).value)
FN = Trim(sheet.Cells(row,3).value)
LN = Trim(sheet.Cells(row,4).value)
Password = Trim(sheet.Cells(row,8).value)
Location = Trim(sheet.Cells(row,6).value)
Manual = Trim(sheet.Cells(row,11).value)
Task_num = Trim(sheet.Cells(row,9).value)
Start_date = Trim(sheet.Cells(row,7).value)



'wscript.echo "Excel Data Unid : " & Password

'Here we are checking user exist in AD or not.
UserFound = "Yes"
userDN()
if UserFound = "Yes" Then
	wscript.echo "User already exist in AD!"
	owb.close
	objexl.quit
	wscript.quit
Else
	'We will call the createUser Function according to location here.
	if UCase(Left(Location,1)) = "N" Then
		NA_CreateUser()
	Elseif UCase(Left(Location,1)) = "E" Then
		EU_CreateUser()
	Elseif UCase(Left(Location,1)) = "A" Then
		AS_CreateUser()
	Else
		wscript.echo "Location doesn't exist!" &vbCrLf &"Please enter the location in excel sheet like for:"_
		&vbCrLf &"USA = NAmerica" &vbCrLf &"Europe = Europe" &vbCrLf &"Asia = Asia"
	End if
End if

	


owb.close
objexl.quit

wscript.quit




'Function userDN By Using Name Translate.
Function userDN()

Set objNetwork = createobject("wscript.Network")
strNetBiosDomain = objNetwork.UserDomain

'Specifing the NetBios name = myDomain/Unid (chemtura\Lvksj).
strNTName = strNetBiosDomain &"\" &Unid

Set objNT = createObject("NameTranslate")

'Initializing NameTranslate by locating global catalog.
objNT.Init ADS_NAME_INITTYPE_GC, ""

'Use the Set method to specify the NT Format of the object name.
on error resume next
objNT.Set ADS_NAME_TYPE_NT4 , strNTName

'Get method is used for getting the distinguished name.
on error resume next
strUserDN = objNT.get(ADS_NAME_TYPE_1779)
if err.number <> 0 Then
	UserFound = "No"
End if

' Escape any "/" characters with backslash escape character.
' All other characters that need to be escaped will be escaped.
strUserDN = Replace(strUserDN, "/", "\/")

' Now bind that Distinguished name with LDAP provider while creating user object.
on error resume next
Set objUser = GetObject("LDAP://" & strUserDN)
'objUser.get("sn")
'wscript.echo "User DN: " & strUserDN
End Function



'Fucntion NA_CreateUser.
Function NA_CreateUser()
	Dim NA_OU_index , NA_OU
	
	if UCase(Left(Manual,1)) = "Y" Then
		NA_OU = inputBox("Enter the OU where the user account to be created" & vbCrLf _
		&"Eg: OU=3rd Party Users,OU=USED,OU=NA,OU=Location,DC=chemtura,DC=com","Manualy Filling OU",_
		"OU=3rd Party Users,OU=USED,OU=NA,OU=Location,DC=chemtura,DC=com")
	
	else
		'There are 16 Location in NA.
		NA_Loc_arr = Array(" NA CA Elmira"," NA US Shelton CT","NA US El Dorado AR Central","NA US El Dorado AR South"_
		,"NA US El Dorado AR West","NA CA West Hill","NA CA Guelph","NA US East Hanover NJ","NA US Fords NJ Chem",_
		"NA US Gastonia NC Chem","NA US Mapleton IL","NA US Middlebury CT","NA US Naugatuck CT","NA US Newtown Square PA",_
		"NA US Perth Amboy NJ","NA US West Lafayette IN")
		
		NA_OU_arr = Array("CAEL","USSH","USED","USES","USEW","CAGT","CAGU","USEH","USFO","USGA","USMA","USMI","USNA",_
		"USNS","USPE","USWL")
		
		'Here we are mapping the OU according to the location.
		for i = 0 to UBound(NA_Loc_arr):
			if Lcase(Location) = Trim(Lcase(NA_Loc_arr(i))) Then
				NA_OU_index = i
				'wscript.echo "The OU is : " &NA_OU_arr(NA_OU_index)
				Exit For
			End if
			if i = UBound(NA_Loc_arr) Then
				wscript.echo "Location not found!"
				NA_OU = inputBox("Enter the OU where the user account to be created" & vbCrLf _
				&"Eg: OU=3rd Party Users,OU=USED,OU=NA,OU=Location,DC=chemtura,DC=com","Manualy Filling OU",_
				"OU=3rd Party Users,OU=USED,OU=NA,OU=Location,DC=chemtura,DC=com")
			End if
		next
		
		'We will check if user is external or regular.
		if UCase(Left(Types,1)) = "E" Then
			'Exceptional case of Mapilton becouse there is only Users OU inside this and it contain both external and regular users.
			If Ucase(Location) = "NA US MAPLETON IL" Then
				NA_OU = "OU=Users,OU="& NA_OU_arr(NA_OU_index) &",OU=NA,OU=Location,DC=chemtura,DC=com"
			Else
				NA_OU = "OU=3rd Party Users,OU="& NA_OU_arr(NA_OU_index) &",OU=NA,OU=Location,DC=chemtura,DC=com"
			End if
		elseif UCase(Left(Types,1)) = "R" Then
			NA_OU = "OU=Users,OU="& NA_OU_arr(NA_OU_index) &",OU=NA,OU=Location,DC=chemtura,DC=com"
		
		else 
			wscript.echo "Check excel sheet if user is external or regular!"
			owb.close
			objexl.quit
			wscript.quit
		End if
	End if
	
	'Here we are creating the account.
	on error resume next
	Set objOU = GetObject("LDAP://"& NA_OU)
	if err.number <> 0 Then
		NA_OU = "OU=3rd party Users,OU="& NA_OU_arr(NA_OU_index) &",OU=NA,OU=Location,DC=chemtura,DC=com"
		on error resume next
		Set objOU = GetObject("LDAP://"& NA_OU)
		if err.number <> 0 Then
			wscript.echo "OU not found inside Location!"
			owb.close
			objexl.quit
			wscript.quit
		End if
	End if
	
	Set objUser = objOU.Create("User","cn=" &FN &" " &LN)
	objUser.put "givenName", FN
	objUser.Put "sAMAccountName", Unid
	objUser.put "sn", LN
	objUser.put "description", Task_num &" - New User ID Created!"
	objUser.put "displayName", LN &", " &FN
	objUser.put "userPrincipalName", Unid & "@chemtura.com"
	objUser.put "employeeNumber", SAP_num

	'objUser.Put "userAccountControl", 512
	objUser.SetInfo
	
	'Setting Password to never expire.
	'objUser.Put "userAccountControl", intUAC XOR &h10000
        
	'Setting Password to change at next logon.
	objUser.Put "pwdLastSet", CLng(0)
	objUser.SetPassword(Password)
	
	objUser.AccountDisabled = false
	objUser.SetInfo
	
	
	'Adding Groups to user account.
	NA_External_Groups = Array("CN=!chemtura Contractor,OU=Groups,DC=chemtura,DC=com","CN=Default Group,OU=Groups,DC=chemtura,DC=com")
	NA_Regular_Groups = Array("CN=!chemtura Employees,OU=Groups,DC=chemtura,DC=com","CN=Default Group,OU=Groups,DC=chemtura,DC=com")
	
	strUserDN = "cn=" &FN &" " &LN &"," &NA_OU
	
	'Checking if user is external or regular and adding group accordingly.
	if UCase(Left(Types,1)) = "E" Then
		For i = 0 to UBound(NA_External_Groups):
			on error resume next
			Set objGroup = GetObject("LDAP://" & NA_External_Groups(i))
			if err.number <> 0 Then
				wscript.echo "We can't add this group this doesn't exist anymore Group Name: " & NA_External_Groups(i)
			else 
				objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(strUserDN)
				objGroup.SetInfo
			End if	
		next
	else
		For i = 0 to UBound(NA_External_Groups):
			on error resume next
			Set objGroup = GetObject("LDAP://" & NA_Regular_Groups(i))
			if err.number <> 0 Then
				wscript.echo "We can't add this group this doesn't exist anymore Group Name: " & NA_Regular_Groups(i)
			else 
				objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(strUserDN)
				objGroup.SetInfo
			End if	
		next
	End if
	
	wscript.echo "User Account Created Successfully inside NA in this OU : " &NA_OU_arr(NA_OU_index)
End Function




'Function EU_CreateUser.
Function EU_CreateUser()

End function




'Function AS_CreateUser.
Function AS_CreateUser()

End function

