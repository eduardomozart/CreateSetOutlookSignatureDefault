'-----------------------------------------------------------------------------
' CreateSetOutlookSignatureDefault
' 
' This script will pull the user info from Active Directory and create a signature for Outlook. 
' This signature will be set as default and can be able to be modified or changed by the user (optional).
'
' Version:
'
' 	1.0.0
'
' Author:
'
'	* Eduardo Mozart de Oliveira (2021)
'
' See Also:
' 
' 	<Create and Set Outlook Signature as Default: https://community.spiceworks.com/scripts/show/1739-create-and-set-outlook-signature-as-default>
'	<Configuring Outlook for the signatures within the users registry: https://ifnotisnull.wordpress.com/automated-outlook-signatures-vbscript/configuring-outlook-for-the-signatures-within-the-users-registry/>
'	<.NET Regex Tester: http://regexstorm.net/tester>
Option Explicit

' Define global variables
Dim boolVerbose, boolDebug
boolVerbose = False
boolDebug = False

Dim boolForceSignature
boolForceSignature = False

' Parse command-line arguments
Dim arrArgs, strArg
Set arrArgs = Wscript.Arguments
For Each strArg In arrArgs
  	Select Case LCase(strArg)
		Case "/force"
			' User can change his/her signature, but it will be overridden into the next reboot
			boolForceSignature = True
		Case "/debug"
			' Enable Debug and Verbose messages on "Debug" flag
			boolVerbose = True
			boolDebug = True
			ShowDebugMessage("[debug] Enable Debug mode." & vbCrLf)
		Case "/verbose"
			' Enable Verbose messages only
			boolVerbose = True
			ShowVerboseMessage("Enable Verbose mode." & vbCrLf)
	End Select
Next

' Ignore errors during execution (do not throw errors to the end user)
On Error Resume Next

' Variables to detect current script path location
' The script path is mounted to a local drive letter if it's a UNC path (e.g. \\server2008\NETLOGON)
Dim strScriptPath, strScriptSignaturesPath, strMapNetworkDriveLetter
strScriptPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName)-1) 
strMapNetworkDriveLetter = GetFirstFreeDriveLetter & "\"

' Create File System object (FSO)
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create ADODB object
' Object used to create UTF-8 signatures files
Dim objStream : Set objStream = CreateObject("ADODB.Stream")

' Create MSWord object
Const wdFormatText = 2, wdFormatRTF = 6, wdFormatHTML = 8

strScriptPath = RemoveTrailingBackslash(strScriptPath)

' If UNC Path (e.g. \\server2008\NETLOGON) Then
If Left(strScriptPath, 2) = "\\" Then
	' Create Drive Mapping (e.g. J: - \\server2008\NETLOGON)
	Dim WshNetwork
	Set WshNetwork = WScript.CreateObject("WScript.Network")
	
	' MapNetworkDrive with first available drive letter (D:, E:, G:, etc.)
	' Example: J: - \\server2008\NETLOGON\CreateSetOutlookSignatureDefault
	WshNetwork.MapNetworkDrive RemoveTrailingBackslash(strMapNetworkDriveLetter), strScriptPath
	' Check if network drive was mounted succesfully listing if strMapNetworkDriveLetter drive letter (D:, E:, G:, etc.) exists
	If objFSO.DriveExists(strMapNetworkDriveLetter) Then
		Dim objDrives, boolMapNetworkDrive, i
		Set objDrives = WshNetwork.EnumNetworkDrives
		boolMapNetworkDrive = False
		For i = 0 to objDrives.Count - 1 Step 2
			ShowDebugMessage("[debug] " & objDrives.Item(i) & " = " & objDrives.Item(i+1))
			If objDrives.Item(i) = RemoveTrailingBackslash(strMapNetworkDriveLetter) Then
				If objDrives.Item(i+1) = strScriptPath Then
					ShowDebugMessage("[debug] Drive " & objDrives.Item(i) & " is now connected to " & objDrives.Item(i+1) & vbCrLf)
					' UNC path was mounted successfully
					boolMapNetworkDrive = True
					Exit For
				End If
			End If
		Next
	End If
	
	' If failed to mount UNC path, exits script
	If Not boolMapNetworkDrive Then
		ShowVerboseMessage("The network connection (" & strScriptPath & ") could not be found.")
		WScript.Quit(2)
	End If
End If

' List scripting information into Verbose/Debug mode
Dim strScriptInformation
strScriptInformation = "Script Information:" & vbCrLf &_
						"------------------" & vbCrLf
If boolMapNetworkDrive Then
	strScriptInformation = strScriptInformation &_
							"strScriptPath: " & strMapNetworkDriveLetter & " (" & strScriptPath & ")" & vbCrLf
	strScriptPath = strMapNetworkDriveLetter
Else
	strScriptInformation = strScriptInformation &_
							"strScriptPath: " & strScriptPath & vbCrLf
End If
strScriptSignaturesPath = objFSO.BuildPath(strScriptPath, "Signatures")
strScriptInformation = strScriptInformation &_
						"strScriptSignaturesPath: " & strScriptSignaturesPath
ShowVerboseMessage(strScriptInformation & vbCrLf)

' Create StdRegProv (Registry) object
Dim objReg, strComputer
strComputer = "."
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Const HKEY_CLASSES_ROOT    = &H80000000
Const HKEY_CURRENT_USER    = &H80000001
Const HKEY_LOCAL_MACHINE   = &H80000002

' Create ADSystemInfo object
' Object used to query user information from Active Directory (AD)
Dim objSysInfo, strUserDomain, strUserDN, objUser
Set objSysInfo = CreateObject("ADSystemInfo")
strUserDomain = objSysInfo.DomainDNSName
strUserDN = objSysInfo.UserName
' objSysInfo.RefreshSchemaCache ' User must belong to "Schema admins" group
Set objUser = GetObject("LDAP://" & strUserDN)

' Create global variables to store Active Directory (AD) user information
' https://www.howto-outlook.com/howto/corporatesignatures.htm
' https://andrelaudari.wordpress.com/2013/05/19/assinatura-de-e-mail-microsoft-exchange-server-2010/
Dim strsAMAccountName, strDisplayName, strFirstName, strInitials, strLastName
strsAMAccountName = objUser.samaccountname
strDisplayName = objUser.displayname
strFirstName = objUser.givenname
strInitials = objUser.initials
strLastName = objUser.sn

Dim strTelephoneNumber, strOtherTelephoneNumber, strHomePhoneNumber, strPagerNumber, strMobileNumber, strFaxNumber, strOtherFaxNumber, strEmail, strEmailDomain
strTelephoneNumber = objUser.telephonenumber
strOtherTelephoneNumber = objUser.othertelephone
strHomePhoneNumber = objUser.HomePhoneNumber
strPagerNumber = objUser.PagerNumber
strMobileNumber = objUser.mobile
strFaxNumber = objUser.facsimiletelephonenumber
strOtherFaxNumber = objUser.otherfacsimiletelephonenumber
strEmail = objUser.mail

Dim strStreet, strPOBox, strCity, strState, strZipCode, strCountry, strCountryAbbreviation
strStreet = objUser.streetaddress
strPOBox = objUser.postofficebox
strCity = objUser.l
strState = objUser.st
strZipCode = objUser.postalcode
strCountry = objUser.co
strCountryAbbreviation = objUser.c

Dim strTitle, strDepartment, strManager, strOffice, strCompany
strTitle = objUser.title
strDepartment = objUser.department
strManager = objUser.manager
strOffice = objUser.physicaldeliveryofficename
strCompany = objUser.company

Dim strWhenChanged
strWhenChanged = objUser.whenchanged

' Print Active Directory (AD) user information into Verbose/Debug mode
ShowVerboseMessage("User Information:" & vbCrLf &_
			"-----------------" & vbCrLf &_
			"strUserDomain: " & strUserDomain & vbCrLf &_
			"strDisplayName: " & strDisplayName & vbCrLf &_
			"strsAMAccountName: " & strsAMAccountName & vbCrLf &_
			"strFirstName: " & strFirstName & vbCrLf &_
			"strInitials: " & strInitials & vbCrLf &_
			"strLastName: " & strLastName & vbCrLf &_
			"strTelephoneNumber: " & strTelephoneNumber & vbCrLf &_
			"strOtherTelephoneNumber: " & strOtherTelephoneNumber & vbCrLf &_
			"strHomePhoneNumber: " & strHomePhoneNumber & vbCrLf &_
			"strPagerNumber: " & strPagerNumber & vbCrLf &_
			"strMobileNumber: " & strMobileNumber & vbCrLf &_
			"strFaxNumber: " & strFaxNumber & vbCrLf &_
			"strOtherFaxNumber: " & strOtherFaxNumber & vbCrLf &_
			"strEmail: " & strEmail & vbCrLf &_
			"strStreet: " & strStreet & vbCrLf &_
			"strPOBox: " & strPOBox & vbCrLf &_
			"strCity: " & strCity & vbCrLf &_
			"strState: " & strState & vbCrLf &_
			"strZipCode: " & strZipCode & vbCrLf &_
			"strCountry: " & strCountry & vbCrLf &_
			"strCountryAbbreviation: " & strCountryAbbreviation & vbCrLf &_
			"strTitle: " & strTitle & vbCrLf &_
			"strDepartment: " & strDepartment & vbCrLf &_
			"strManager: " & strManager & vbCrLf &_
			"strOffice: " & strOffice & vbCrLf &_
			"strCompany: " & strCompany & vbCrLf &_
			"strWhenChanged: " & strWhenChanged & " UTC" & vbCrLf)

' TODO: Outlook must be closed to edit current Outlook account signatures? Answer: No.
' If Not IsOutlookRunning Then
	Dim strRootProfilesKey, strProfilesKey, arrProfilesKeys, strProfilesSubKey, arrProfilesSubKeys, strCurrentAccountKey
	' Dim arrSigValue
	
	' If Office edition >= Office 2013 Then
	' https://stackoverflow.com/questions/29153766/how-do-i-determine-if-user-has-a-default-outlook-2013-profile
	If FirstVersionSupOrEqualToSecondVersion(GetOfficeVersion, "15.0") Then
		strRootProfilesKey = "Software\Microsoft\Office\" & _
		GetOfficeVersion & "\Outlook\Profiles"
	Else
		strRootProfilesKey = "Software\Microsoft\Windows NT\" & _
		"CurrentVersion\Windows " & _
		"Messaging Subsystem\Profiles"
	End If
	
	' Print Office edition into Verbose/Debug mode
	ShowVerboseMessage("GetOfficeVersion: " & GetOfficeVersion)
	
	' If any Office edition is installed
	If (Len(GetOfficeVersion)) > 0 Then
		Dim arrSignaturesCSVLines
		' Parse signature rules file to determine if E-mail Account match any RegExp from this file.
		arrSignaturesCSVLines = ParseSignatureFileVariables(objFSO.BuildPath(strScriptPath, "signatures.csv"))
		
		ShowVerboseMessage("strRootProfilesKey: HKEY_CURRENT_USER\" & strRootProfilesKey & vbCrLf)
		
		' For Each Outlook Profiles
		objReg.EnumKey HKEY_CURRENT_USER, strRootProfilesKey, arrProfilesKeys
		For Each strProfilesKey In arrProfilesKeys
			
			' For Each Outlook Accounts into Each Outlook Profile
			objReg.EnumKey HKEY_CURRENT_USER, strRootProfilesKey & "\" & strProfilesKey & "\9375CFF0413111d3B88A00104B2A6676", arrProfilesSubKeys
			For Each strProfilesSubKey In arrProfilesSubKeys
			
				strCurrentAccountKey = strRootProfilesKey & "\" & strProfilesKey & _
										"\9375CFF0413111d3B88A00104B2A6676\" & strProfilesSubKey
				
				' Outlook >= 2016 store Account information as String values,
				' Outlook <= 2013 store them as DWORD values			
				Dim strEmailAccountNewSignatureValue, strEmailAccountReplyForwardSignatureValue
				Dim arrEmailAccountEmailValue, arrEmailAccountNewSignatureValue, arrEmailAccountReplyForwardSignatureValue
				If FirstVersionSupOrEqualToSecondVersion(GetOfficeVersion, "16.0") Then
					objReg.GetStringValue HKEY_CURRENT_USER, strCurrentAccountKey, "Email", strEmail
					objReg.GetStringValue HKEY_CURRENT_USER, strCurrentAccountKey, "New Signature", strEmailAccountNewSignatureValue
					objReg.GetStringValue HKEY_CURRENT_USER, strCurrentAccountKey, "Reply-Forward Signature", strEmailAccountReplyForwardSignatureValue
				Else
					objReg.GetBinaryValue HKEY_CURRENT_USER, strCurrentAccountKey, "Email", arrEmailAccountEmailValue
					strEmail = RegBinaryToString(arrEmailAccountEmailValue)
					
					objReg.GetBinaryValue HKEY_CURRENT_USER, strCurrentAccountKey, "New Signature", arrEmailAccountNewSignatureValue
					strEmailAccountNewSignatureValue = RegBinaryToString(arrEmailAccountNewSignatureValue)
					
					objReg.GetBinaryValue HKEY_CURRENT_USER, strCurrentAccountKey, "Email", arrEmailAccountReplyForwardSignatureValue
					strEmailAccountReplyForwardSignatureValue = RegBinaryToString(arrEmailAccountReplyForwardSignatureValue)
				End If
				
				' RegExp to check if E-mail Account contains a valid e-mail.
				Dim boolValidEmail
				boolValidEmail = False
				strEmailDomain = Join(GetRegExMatches(strEmail, "@((?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9]))"))
				If UBound(GetRegExMatches(strEmail, "([a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)")) >= 0 Then
					boolValidEmail = True
				End If
				
				' List Registry Key information For Each Outlook Account.
				ShowVerboseMessage("Script Main function:" & vbCrLf &_
							"---------------------" & vbCrLf &_
							"strCurrentAccountKey: HKEY_CURRENT_USER\" & strCurrentAccountKey & vbCrLf &_
							"strEmail: " & strEmail & vbCrLf &_
							"strEmailDomain: " & strEmailDomain & vbCrLf &_
							"strEmailAccountNewSignatureValue: " & strEmailAccountNewSignatureValue & vbCrLf &_
							"strEmailAccountReplyForwardSignatureValue: " & strEmailAccountReplyForwardSignatureValue & vbCrLf &_
							"boolValidEmail: " & boolValidEmail & vbCrLf)
				
				' For Each Account with a valid mail, call SetProfileSignature() function.
				If boolValidEmail Then
					SetProfileSignature(Array(strCurrentAccountKey, _
												strEmail, _
												strEmailAccountNewSignatureValue, _
												strEmailAccountReplyForwardSignatureValue))
				End If
			Next
		Next
	Else
		' If Outlook is not installed, print message into Verbose/Debug mode and exits the script.
		ShowVerboseMessage("Please install Outlook before " & _
			"running this script.")
	End If
' Else
	' If Outlook is running, print message to user and exits the script.
	' MsgBox "Please shut down Outlook before " & _
	'		"running this script.", vbExclamation, "SetDefaultSignature"
' End If

On Error Goto 0

' Unmount UNC Drive Mapping (e.g. J: - \\server2008\NETLOGON) after script execution.
If boolMapNetworkDrive Then
	ShowVerboseMessage("WshNetwork.RemoveNetworkDrive " & RemoveTrailingBackslash(strMapNetworkDriveLetter))
	WshNetwork.RemoveNetworkDrive RemoveTrailingBackslash(strMapNetworkDriveLetter), True
	If objFSO.DriveExists(strMapNetworkDriveLetter) Then
		ShowVerboseMessage("Error on WshNetwork.RemoveNetworkDrive.")
		WScript.Quit(3)
	End If
End If

'
' INTERNAL FUNCTIONS
'

' Function: SetProfileSignature
'
'   This function is used to check if creating/setting the signature to an Outlook E-mail account is necessary by comparing:
' 
'   - 1) If Current System Date is within a Signature ruleset date range (see <ParseSignatureFileVariables>).
'   - 2) If Current Outlook E-mail Account matches the RegExp from Signature ruleset above. 
'   
'   Parameters:
'
'      arrEmailAccountInfo - Outlook E-mail Account information (e-mail account registry path, e-mail account and current signature filename).
'
'   Returns:
'
'      None.
'
'   See Also:
'
'      <ParseSignatureFileVariables>
'
Sub SetProfileSignature(arrEmailAccountInfo)
	ShowVerboseMessage("SetProfileSignature sub procedure:" & vbCrLf &_
				"-----------------------------" & vbCrLf)
	'For i = LBound(arrEmailAccountInfo) To UBound(arrEmailAccountInfo)
	'	ShowVerboseMessage("arrEmailAccountInfo(" & i & "): " &_
	'					arrEmailAccountInfo(i))
	'Next
	
	Dim i, j
	' If signatures.csv file isn't empty.
	If UBound(arrSignaturesCSVLines) >= 0 Then
		' Parse Each signatures.csv file lines.
		For i = LBound(arrSignaturesCSVLines) To UBound(arrSignaturesCSVLines)
			ShowVerboseMessage("Line #" & i+1 & ":")
			
			Dim arrSignatureCSVValues
			' Split each line with ';' to arrSignatureCSVValues
			arrSignatureCSVValues = Split(arrSignaturesCSVLines(i), ";")
			For j = LBound(arrSignatureCSVValues) To UBound(arrSignatureCSVValues)
				ShowVerboseMessage(vbTab & "arrSignatureCSVValues(" & j & "): " & arrSignatureCSVValues(j))
			Next
			
			' Set local sub procedures variables with splitted signatures.csv line content. Example:
			' 01/12/1995;01/13/1995;Default;(\w+@example.com)
			Dim dtmSigCSVStartDate, dtmSigCSVEndDate, strSigCSVSubFolderName, strSigCSVEmailRegEx
			dtmSigCSVStartDate = "#" & CDate(arrSignatureCSVValues(0)) & "#" ' 01/12/1995
			dtmSigCSVEndDate = "#" & CDate(arrSignatureCSVValues(1)) & "#" ' 01/13/1995
			strSigCSVSubFolderName = arrSignatureCSVValues(2) ' Default
			strSigCSVEmailRegEx = arrSignatureCSVValues(3) ' (\w+@example.com)
			
			Dim strSigSubFolderRemotePath, strSigHTMLRemoteFilePath
			Dim strSigHTMLLocalFileBaseName
			' Example: strSigSubFolderRemotePath = objFSO.BuildPath(J:\Signatures, Default)
			strSigSubFolderRemotePath = objFSO.BuildPath(strScriptSignaturesPath, strSigCSVSubFolderName)
			' strSigHTMLRemoteFilePath = objFSO.BuildPath(strSigSubFolderRemotePath, arrEmailAccountInfo(1) & ".htm")
			' Example: strSigCSVSubFolderName = Default - it@contoso.com
			strSigHTMLLocalFileBaseName = strSigCSVSubFolderName & " - " & arrEmailAccountInfo(1)
			
			Dim boolSetCustomProfileSignature
			boolSetCustomProfileSignature = False
			' Example: If J:\Signatures\Default folder exists.
			If objFSO.FolderExists(strSigSubFolderRemotePath) Then
				' If Current System Date between Signatures.csv ruleset date range. Examples:
				' If 01/12/1995 between (01/12/1995 and 01/13/1995) = True
				' If 02/24/2021 between (01/12/1995 and 01/13/1995) = False
				Dim dtmNowDate
				dtmNowDate = "#" & CDate(Month(Now) & "/" & Day(Now) & "/" & Year(Now)) & "#"
				' https://devblogs.microsoft.com/scripting/how-can-i-tell-if-a-date-falls-within-a-specified-time-period/
				' https://stackoverflow.com/questions/15195089/find-if-date-is-more-than-another-date/15222901#15222901
				If dtmNowDate >= dtmSigCSVStartDate And dtmNowDate <= dtmSigCSVEndDate Then
					' If signatures.csv 'RegExp' field matches with Outlook Email Account 'Email' registry value.
					If UBound(GetRegExMatches(arrEmailAccountInfo(1), strSigCSVEmailRegEx)) >= 0 Then
						boolSetCustomProfileSignature = True
					Else
						ShowVerboseMessage("Email " & arrEmailAccountInfo(1) & " doesn't match RegExp " & strSigCSVEmailRegEx & ". [Continue]")
					End If
				Else
					ShowVerboseMessage("Date " & dtmNowDate & " isn't within the specified range: " & dtmSigCSVStartDate & " - " & dtmSigCSVEndDate & ". [Continue]")
				End If
			Else
				ShowVerboseMessage("Remote signature subfolder """ & strSigSubFolderRemotePath & """ doesn't exists. [Continue]")
			End If
			
			ShowVerboseMessage("boolSetCustomProfileSignature: " & boolSetCustomProfileSignature & vbCrLf)
			
			' If current Outlook E-mail Account matches the current Signature ruleset conditions Then:
			' 1) Create signature file.
			' 2) Set created signature by default when compose/reply e-mail messages using the current Outlook E-mail Account.
			If (boolSetCustomProfileSignature) Then
				Dim strSigNewSignatureName, strSigNewSignatureRemotePath
				
				ShowVerboseMessage(vbCrLf & vbTab & "Setting signature: " & strSigCSVSubFolderName & "." & vbCrLf)
				strSigNewSignatureName = strSigHTMLLocalFileBaseName
				strSigNewSignatureRemotePath = strSigSubFolderRemotePath
				
				' If the signature file was created successfully.
				If (CreateSignatureFiles(arrEmailAccountInfo, strSigNewSignatureRemotePath, strSigNewSignatureName)) Then
					' We're creating signatures file based on HTML files instead of creating signatures through MS Outlook.
					' Images from signatures created outside Outlook aren't send with e-mail messages by default.
					' The symptom is that the user can see the images within the signature when composing/replying e-mail messages,
					' but the recipient doesn't receive them.
					' We'll need to create the DWORD value 'Send Pictures With Document' with the value '1' to fix this issue.
					' https://support.microsoft.com/en-gb/help/2779191/inline-images-may-display-as-a-red-x-in-outlook
					Dim strOutlookOptionsMailKey, intSendPicturesWithDocument, boolSetSendPicturesWithDocumentsValue
					strOutlookOptionsMailKey = "Software\Microsoft\Office\" & GetOfficeVersion & "\Outlook\Options\Mail"
					Dim iRet : iRet = objReg.GetDWORDValue(HKEY_CURRENT_USER, strOutlookOptionsMailKey, "Send Pictures With Document", intSendPicturesWithDocument)
					ShowDebugMessage(vbTab & "objReg.GetDWORDValue(HKEY_CURRENT_USER, """ & strOutlookOptionsMailKey & """, ""Send Pictures With Document"", intSendPicturesWithDocument)" & vbCrLf &_
									vbTab & "iRet = " & iRet & vbCrLf &_
									vbTab & "intSendPicturesWithDocument = " & intSendPicturesWithDocument)
					
					boolSetSendPicturesWithDocumentsValue = True
					If iRet = 0 Then
						If intSendPicturesWithDocument = 1 Then
							boolSetSendPicturesWithDocumentsValue = False
						End If
					End If
					
					If boolSetSendPicturesWithDocumentsValue Then
						Dim arrOutlookOptionsMailKeys
						iRet = objReg.EnumKey(HKEY_CURRENT_USER, strOutlookOptionsMailKey, arrOutlookOptionsMailKeys)
						If iRet <> 0 Then
							objReg.CreateKey HKEY_CURRENT_USER, strOutlookOptionsMailKey
						End If
						
						ShowVerboseMessage(vbTab & "objReg.SetDWORDValue HKEY_CURRENT_USER, """ & strOutlookOptionsMailKey & """, ""Send Pictures With Document"", 1")
						objReg.SetDWORDValue HKEY_CURRENT_USER, strOutlookOptionsMailKey, "Send Pictures With Document", 1
					End If
					
					' Update Outlook E-mail account information into the Windows registry to use the new signature file by default.
					If FirstVersionSupOrEqualToSecondVersion(GetOfficeVersion, "16.0") Then
						ShowVerboseMessage(vbTab & "objReg.SetStringValue HKEY_CURRENT_USER, " & arrEmailAccountInfo(0) & ", ""New Signature"", """ & strSigNewSignatureName & """")
						objReg.SetStringValue HKEY_CURRENT_USER, arrEmailAccountInfo(0), "New Signature", strSigNewSignatureName
						ShowVerboseMessage(vbTab & "objReg.SetStringValue HKEY_CURRENT_USER, " & arrEmailAccountInfo(0) & ", ""Reply-Forward Signature"", """ & strSigNewSignatureName & """" & vbCrLf)
						objReg.SetStringValue HKEY_CURRENT_USER, arrEmailAccountInfo(0), "Reply-Forward Signature", strSigNewSignatureName
					Else
						ShowVerboseMessage(vbTab & "objReg.SetBinaryValue HKEY_CURRENT_USER, " & arrEmailAccountInfo(0) & ", ""New Signature"", """ & Join(StringToByteArray(strSigNewSignatureName, True), ",") & """")
						objReg.SetBinaryValue HKEY_CURRENT_USER, arrEmailAccountInfo(0), "New Signature", StringToByteArray(strSigNewSignatureName, True)
						ShowVerboseMessage(vbTab & "objReg.SetBinaryValue HKEY_CURRENT_USER, " & arrEmailAccountInfo(0) & ", ""Reply-Forward Signature"", """ & Join(StringToByteArray(strSigNewSignatureName, True), ",") & """" & vbCrLf)
						objReg.SetBinaryValue HKEY_CURRENT_USER, arrEmailAccountInfo(0), "Reply-Forward Signature", StringToByteArray(strSigNewSignatureName, True)
					End If
				End If
			End If
		Next
	End If
End Sub

' Function: SetProfileSignature
'
'   This function is called by SetProfileSignature() sub procedure to create signature HTML, RTF and TXT signature files
'   to a Outlook e-mail account.
' 
'   Parameters:
'
'      arrEmailAccountInfo - Outlook E-mail Account information (e-mail account registry path, e-mail account and current signature filename).
'      
'
'   Returns:
'
'      None.
'
'   See Also:
'
'      <ParseSignatureFileVariables>
'
Function CreateSignatureFiles(arrEmailAccountInfo, strSigNewSignatureRemotePath, strSigNewSignatureName)
	ShowVerboseMessage(vbTab & "CreateSignatureFiles function:" & vbCrLf &_
				vbTab & "------------------------------" & vbCrLf &_
				vbTab & "strSigNewSignatureRemotePath: " & strSigNewSignatureRemotePath & vbCrLf &_
				vbTab & "strSigNewSignatureName: " & strSigNewSignatureName & vbCrLf)
	
	Dim boolRemoteHTMLSignatureFileExists
	boolRemoteHTMLSignatureFileExists = False
	
	Dim boolRemoteTXTSignatureFileExists
	boolRemoteTXTSignatureFileExists = False
	
	Dim arrRemoteSignatureFiles
	arrRemoteSignatureFiles = Array(objFSO.BuildPath(strSigNewSignatureRemotePath, strEmail), _
										objFSO.BuildPath(strSigNewSignatureRemotePath, strEmailDomain), _
										objFSO.BuildPath(strSigNewSignatureRemotePath, "default"))
	
	' arrRemoteSignatureFiles(0) = E:\Signatures\Christmas\example@example.com.[htm|txt]
	' arrRemoteSignatureFiles(1) = E:\Signatures\Christmas\example.com.[htm|txt]
	' arrRemoteSignatureFiles(2) = E:\Signatures\Christmas\default.[htm|txt]
	
	Dim strOutlookHTMLSignatureRemotePath, strOutlookTXTSignatureRemotePath
	Dim strOutlookHTMLSignatureRemoteFilePath, strOutlookTXTSignatureRemoteFilePath
	For i = LBound(arrRemoteSignatureFiles) To UBound(arrRemoteSignatureFiles)
		If Not boolRemoteHTMLSignatureFileExists Then
			If objFSO.FileExists(arrRemoteSignatureFiles(i) & ".htm") Then
				strOutlookHTMLSignatureRemotePath = objFSO.GetAbsolutePathName(arrRemoteSignatureFiles(i) & ".htm")
				strOutlookHTMLSignatureRemoteFilePath = arrRemoteSignatureFiles(i) & ".htm"
				boolRemoteHTMLSignatureFileExists = True
			End If
		End If
		
		If Not boolRemoteTXTSignatureFileExists Then
			If objFSO.FileExists(arrRemoteSignatureFiles(i) & ".txt") Then
				strOutlookTXTSignatureRemotePath = objFSO.GetAbsolutePathName(arrRemoteSignatureFiles(i) & ".txt")
				strOutlookTXTSignatureRemoteFilePath = arrRemoteSignatureFiles(i) & ".txt"
				boolRemoteTXTSignatureFileExists = True
			End If
		End If
	Next
	
	Dim objShell, strUserAppDataPath, strOutlookLocalizedSignatureFolderName, strOutlookSignaturesPath
	Set objShell = CreateObject("WScript.Shell")
	strUserAppDataPath = objShell.ExpandEnvironmentStrings("%APPDATA%")
	objReg.GetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Office\" & GetOfficeVersion & "\Common\General", "Signatures", strOutlookLocalizedSignatureFolderName
	strOutlookSignaturesPath = objFSO.BuildPath(strUserAppDataPath, "Microsoft\" & strOutlookLocalizedSignatureFolderName)
	
	ShowVerboseMessage(vbTab & "strOutlookSignaturesPath: " & strOutlookSignaturesPath)
	
	If Not (objFSO.FolderExists(strOutlookSignaturesPath)) Then
		objFSO.CreateFolder(strOutlookSignaturesPath)
	End If
	
	Dim strLocalHTMLFilePath, strLocalHTMLConnectedFilesRelativePath, strLocalHTMLConnectedFilesPath
	Dim strLocalRTFFilePath
	Dim strLocalTXTFilePath
	strLocalHTMLFilePath = objFSO.BuildPath(strOutlookSignaturesPath, strSigNewSignatureName & ".htm")
	strLocalHTMLConnectedFilesRelativePath = strSigNewSignatureName & "_files"
	strLocalHTMLConnectedFilesPath = objFSO.BuildPath(strOutlookSignaturesPath, strLocalHTMLConnectedFilesRelativePath)
	strLocalRTFFilePath = objFSO.BuildPath(strOutlookSignaturesPath, strSigNewSignatureName & ".rtf")
	strLocalTXTFilePath = objFSO.BuildPath(strOutlookSignaturesPath, strSigNewSignatureName & ".txt")
	
	ShowVerboseMessage(vbTab & "strLocalHTMLFilePath: " & strLocalHTMLFilePath & vbCrLf &_
				vbTab & "strLocalHTMLConnectedFilesRelativePath: " & strLocalHTMLConnectedFilesRelativePath & vbCrLf &_
				vbTab & "strLocalHTMLConnectedFilesPath: " & strLocalHTMLConnectedFilesPath & vbCrLf &_
				vbTab & "strLocalRTFFilePath: " & strLocalRTFFilePath & vbCrLf &_
				vbTab & "strLocalTXTFilePath: " & strLocalTXTFilePath & vbCrLf)
	
	ShowVerboseMessage(vbTab & "strOutlookHTMLSignatureRemotePath: " & strOutlookHTMLSignatureRemotePath & vbCrLf &_
				vbTab & "strOutlookHTMLSignatureRemoteFilePath: " & strOutlookHTMLSignatureRemoteFilePath & vbCrLf &_
				vbTab & "boolRemoteHTMLSignatureFileExists: " & boolRemoteHTMLSignatureFileExists & vbCrLf)
				
	ShowVerboseMessage(vbTab & "strOutlookTXTSignatureRemotePath: " & strOutlookTXTSignatureRemotePath & vbCrLf &_
				vbTab & "strOutlookTXTSignatureRemoteFilePath: " & strOutlookTXTSignatureRemoteFilePath & vbCrLf &_
				vbTab & "boolRemoteTXTSignatureFileExists: " & boolRemoteTXTSignatureFileExists & vbCrLf)
		
	If boolRemoteHTMLSignatureFileExists Or boolRemoteTXTSignatureFileExists Then
		Dim arrRemoteTXTFileContents
		If boolRemoteTXTSignatureFileExists Then
			arrRemoteTXTFileContents = ParseSignatureFileVariables(strOutlookTXTSignatureRemoteFilePath)
		End If
		
		Dim strRemoteHTMLFileContents, arrRemoteHTMLFileContents
		Dim arrRemoteHTMLFileContentsURIs
		If boolRemoteHTMLSignatureFileExists Then
			arrRemoteHTMLFileContents = ParseSignatureFileVariables(strOutlookHTMLSignatureRemoteFilePath)
		Else
			arrRemoteHTMLFileContents = arrRemoteTXTFileContents
		End If
		
		Dim boolReplaceHTMLLocalSignatureFile
		boolReplaceHTMLLocalSignatureFile = False
		
		If objFSO.FileExists(strLocalHTMLFilePath) Then
			' Allow the user customize the HTML signature, replacing it only if Remote signature is modified (newer than Local signature).
			If FirstFileAreNewer(strOutlookHTMLSignatureRemoteFilePath, strLocalHTMLFilePath) Then
				ShowVerboseMessage(vbTab & "Remote HTML signature file """ & strOutlookHTMLSignatureRemoteFilePath & """ is newer than """ & strLocalHTMLFilePath & """." & vbCrLf)
				boolReplaceHTMLLocalSignatureFile = True
			Else
				ShowVerboseMessage(vbTab & "Remote HTML signature file """ & strOutlookHTMLSignatureRemoteFilePath & """ isn't newer than """ & strLocalHTMLFilePath & """." & vbCrLf)
				Dim strLocalHTMLFileDateLastModified
				strLocalHTMLFileDateLastModified = objFSO.GetFile(strLocalHTMLFilePath).DateLastModified
				If CDate(ConvertToUTCTime(strLocalHTMLFileDateLastModified)) > CDate(strWhenChanged) Then
					ShowVerboseMessage(vbTab & "AD User atribute 'whenChanged' (" & strWhenChanged & " UTC) isn't newer than """ & strLocalHTMLFilePath & """ (" & ConvertToUTCTime(strLocalHTMLFileDateLastModified) & " UTC)." & vbCrLf)
				Else
					boolReplaceHTMLLocalSignatureFile = True
					ShowVerboseMessage(vbTab & "AD User atribute 'whenChanged' (" & strWhenChanged & " UTC) is newer than """ & strLocalHTMLFilePath & """ (" & ConvertToUTCTime(strLocalHTMLFileDateLastModified) & " UTC)." & vbCrLf)
				End If
			End If
		Else
			ShowVerboseMessage(vbTab & "Local HTML signature file """ & strLocalHTMLFilePath & """ doesn't exist." & vbCrLf)
			boolReplaceHTMLLocalSignatureFile = True
		End If
		
		' Parse and download HTML Remote attachments (src tag) if they are newer than the Local ones
		' E.g. <img src="example@example.com.png" = example@example.com.png
		If Not objFSO.FolderExists(strLocalHTMLConnectedFilesPath) Then
			objFSO.CreateFolder(strLocalHTMLConnectedFilesPath)
		End If
		
		Dim strRegExPattern
		strRegExPattern = "src\s*=\s*""(.+?)"""
		'                 src\s*=\s*"(.+?)"
			
		strRemoteHTMLFileContents = Join(arrRemoteHTMLFileContents, vbCrLf)
		ShowVerboseMessage(vbTab & "GetRegExMatches(strRemoteHTMLFileContents, " & strRegExPattern & ")")
		arrRemoteHTMLFileContentsURIs = RemoveDupsArray(GetRegExMatches(strRemoteHTMLFileContents, strRegExPattern))
		ShowVerboseMessage(vbTab & "UBound(arrRemoteHTMLFileContentsURIs): " & UBound(arrRemoteHTMLFileContentsURIs) & vbCrLf)
						
		Dim strURIMatch, strURIRemoteFilePath, strURILocalFilePath
		Dim boolURIMatchFileExists
		boolURIMatchFileExists = False
		For i = LBound(arrRemoteHTMLFileContentsURIs) To UBound(arrRemoteHTMLFileContentsURIs)
			' src="example@example.com.png = example@example.com.png
			strURIMatch = arrRemoteHTMLFileContentsURIs(i)
			ShowVerboseMessage(vbTab & "arrRemoteHTMLFileContentsURIs(" & i & "): " & arrRemoteHTMLFileContentsURIs(i))
							
			' Check if FileExists as Full Path
			If objFSO.FileExists(strURIMatch) Then
				strURIRemoteFilePath = strURIMatch
				boolURIMatchFileExists = True
							
			' Check if FileExists as Relative Path
			ElseIf objFSO.FileExists(objFSO.BuildPath(strSigNewSignatureRemotePath, strURIMatch)) Then
				strURIRemoteFilePath = objFSO.BuildPath(strSigNewSignatureRemotePath, strURIMatch)
				boolURIMatchFileExists = True
							
			End If
			ShowVerboseMessage(vbTab & "strURIRemoteFilePath: " & strURIRemoteFilePath)
			ShowVerboseMessage(vbTab & "boolURIMatchFileExists: " & boolURIMatchFileExists & vbCrLf)
						
			' Remote file exists. External Web files (E.g: "HTTP(S)", "FTP(S)" protocols) are not checked and returns "False".
			If boolURIMatchFileExists Then
				' example@example.com.png = Natal - example@example.com\example@example.com.png
				strRemoteHTMLFileContents = RegExReplace(Join(arrRemoteHTMLFileContents, vbCrLf), _
															arrRemoteHTMLFileContentsURIs(i), _
															objFSO.BuildPath(strLocalHTMLConnectedFilesRelativePath, _
																				arrRemoteHTMLFileContentsURIs(i)))
				ShowVerboseMessage(vbTab & "strRemoteHTMLFileContents: " & vbCrLf & strRemoteHTMLFileContents & vbCrLf)
				arrRemoteHTMLFileContents = Split(strRemoteHTMLFileContents, vbCrLf)
					
				' Check if Remote file are newer than the Local one
				strURILocalFilePath = objFSO.BuildPath(strLocalHTMLConnectedFilesPath, strURIMatch)
				If FirstFileAreNewer(strURIRemoteFilePath, strURILocalFilePath) Then
					ShowVerboseMessage(vbTab & "objFSO.CopyFile """ & strURIRemoteFilePath & """, """ & strURILocalFilePath & """")
					objFSO.CopyFile strURIRemoteFilePath, strURILocalFilePath, True
					' The line below will recreate the RTF signature file.
					boolReplaceHTMLLocalSignatureFile = True
				Else
					ShowVerboseMessage(vbTab & "Local file """ & strURILocalFilePath & """ is newer than """ & strURIRemoteFilePath & """.")
				End If
			End If
		Next
		
		Dim boolReplaceTXTLocalSignatureFile
		boolReplaceTXTLocalSignatureFile = True
		
		If objFSO.FileExists(strLocalTXTFilePath) Then
			If FirstFileAreNewer(strOutlookTXTSignatureRemoteFilePath, strLocalTXTFilePath) Then
				ShowVerboseMessage(vbTab & "Remote TXT signature file """ & strOutlookTXTSignatureRemoteFilePath & """ is newer than """ & strLocalTXTFilePath & """." & vbCrLf)
				boolReplaceTXTLocalSignatureFile = True
			Else
				ShowVerboseMessage(vbTab & "Remote TXT signature file """ & strOutlookTXTSignatureRemoteFilePath & """ isn't newer than """ & strLocalTXTFilePath & """." & vbCrLf)
				Dim strLocalTXTFileDateLastModified
				strLocalTXTFileDateLastModified = objFSO.GetFile(strLocalTXTFilePath).DateLastModified
				If CDate(ConvertToUTCTime(strLocalTXTFileDateLastModified)) > CDate(strWhenChanged) Then
					ShowVerboseMessage(vbTab & "AD User attribute 'whenChanged' (" & strWhenChanged & " UTC) isn't newer than """ & strLocalTXTFilePath & """ (" & ConvertToUTCTime(strLocalTXTFileDateLastModified) & " UTC)." & vbCrLf)
				Else
					boolReplaceTXTLocalSignatureFile = True
					ShowVerboseMessage(vbTab & "AD User attribute 'whenChanged' (" & strWhenChanged & " UTC) is newer than """ & strLocalTXTFilePath & """ (" & ConvertToUTCTime(strLocalTXTFileDateLastModified) & " UTC)." & vbCrLf)
				End If
			End If
		Else
			ShowVerboseMessage(vbTab & "Local TXT signature file """ & strLocalTXTFilePath & """ doesn't exist." & vbCrLf)
			boolReplaceTXTLocalSignatureFile = True
		End If
		
		If boolForceSignature Then
			ShowVerboseMessage(vbTab & "**INFO** boolForceSignature: Forcing Remote signature." & vbCrLf)
			boolReplaceHTMLLocalSignatureFile = True
			boolReplaceTXTLocalSignatureFile = True
		End If
		
		ShowVerboseMessage(vbTab & "boolForceSignature: " & boolForceSignature & vbCrLf &_
					vbTab & "boolReplaceHTMLLocalSignatureFile: " & boolReplaceHTMLLocalSignatureFile & vbCrLf &_
					vbTab & "boolReplaceTXTLocalSignatureFile: " & boolReplaceTXTLocalSignatureFile & vbCrLf)
		
		If boolReplaceHTMLLocalSignatureFile Or boolReplaceTXTLocalSignatureFile Then
			
			Dim objFile
			
			If boolReplaceHTMLLocalSignatureFile Then
				ShowVerboseMessage(vbTab & "**INFO** boolReplaceHTMLLocalSignatureFile: Updating HTML and RTF Local signatures files." & vbCrLf)
				
				' HTML File
				If objFSO.FileExists(strLocalHTMLFilePath) Then
					objFSO.DeleteFile strLocalHTMLFilePath
				End If
				
				ShowVerboseMessage(vbTab & "objStream.SaveToFile: " & strLocalHTMLFilePath & vbCrLf)
				objStream.CharSet = "utf-8"
				objStream.Open
				objStream.WriteText Join(arrRemoteHTMLFileContents, vbCrLf)
				objStream.SaveToFile strLocalHTMLFilePath
				objStream.Close

				' RTF File
				WordSaveAs strOutlookHTMLSignatureRemoteFilePath, strLocalRTFFilePath, wdFormatRTF
			End If
			
			If boolReplaceTXTLocalSignatureFile Then
				If objFSO.FileExists(strLocalTXTFilePath) Then
					objFSO.DeleteFile strLocalTXTFilePath
				End If
				
				' TXT File
				ShowVerboseMessage(vbTab & "objStream.SaveToFile: " & strLocalTXTFilePath & vbCrLf)
				objStream.CharSet = "_autodetect"
				objStream.Open
				objStream.WriteText Join(arrRemoteTXTFileContents, vbCrLf)
				objStream.SaveToFile strLocalTXTFilePath
				objStream.Close
			End If
		Else
			ShowVerboseMessage(vbTab & "**INFO** Local signature is newer than Remote signature.")
		End If
		
		ShowVerboseMessage(vbTab & "CreateSignatureFiles: " & True & vbCrLf)
		CreateSignatureFiles = True
	Else
		ShowVerboseMessage(vbTab & "**WARNING** Remote HTML and TXT files doesn't exists. Aborting.")
		ShowVerboseMessage(vbTab & "CreateSignatureFiles: " & False & vbCrLf)
		CreateSignatureFiles = False
	End If
End Function

Function WordSaveAs(strOpenFilePath, strSaveAsFilePath, wdSaveFormat)
	ShowVerboseMessage(vbTab & "WordSaveAs function:" & vbCrLf &_
				vbTab & "--------------------" & vbCrLf &_
				vbTab & "strOpenFilePath: " & strOpenFilePath & vbCrLf &_
				vbTab & "strSaveAsFilePath: " & strSaveAsFilePath)
				
	Select Case wdSaveFormat
		Case wdFormatHTML
			ShowVerboseMessage(vbTab & "wdSaveFormat: wdFormatHTML (" & wdSaveFormat & ")" & vbCrLf)
		Case wdFormatRTF
			ShowVerboseMessage(vbTab & "wdSaveFormat: wdFormatRTF (" & wdSaveFormat & ")" & vbCrLf)
		Case wdFormatText
			ShowVerboseMessage(vbTab & "wdSaveFormat: wdFormatText (" & wdSaveFormat & ")" & vbCrLf)
		Case Else
			ShowVerboseMessage(vbTab & "wdSaveFormat: " & wdSaveFormat & vbCrLf)
	End Select
	
	Dim objWord, objDoc
	
	If objFSO.FileExists(strOpenFilePath) Then
		If objFSO.FolderExists(objFSO.GetParentFolderName(strSaveAsFilePath)) Then
			Set objWord = CreateObject("Word.Application")
			
			' https://stackoverflow.com/questions/8807153/vbscript-to-convert-word-doc-to-pdf
			If Err.Number <> 0 Then
				Select Case Err.Number
					Case &H80020009
						ShowVerboseMessage("**ERROR** Word not installed properly.")
					Case Else
						ShowDefaultErrorMsg
				End Select
				objWord.Quit
			Else
				Set objDoc = objWord.Documents.Open(strOpenFilePath)
				objDoc.SaveAs strSaveAsFilePath, wdSaveFormat
				objDoc.Close
				objWord.Quit
			End If
			
		Else
			ShowVerboseMessage("**ERROR** The folder """ & strSaveAsFilePath & """ doesn't exist." & vbCrLf)
		End If
	Else
		ShowVerboseMessage("**ERROR** The file """ & strOpenFilePath & """ doesn't exist." & vbCrLf)
	End If
End Function

Sub ShowDefaultErrorMsg
    ShowVerboseMessage("Error #" & CStr(Err.Number) & vbNewLine & vbNewLine & Err.Description)
End Sub

Function ParseSignatureFileVariables(strSignatureFilePath)
	ShowVerboseMessage(vbTab & "ParseSignatureFileVariables function:" & vbCrLf &_
				vbTab & "-------------------------------------" & vbCrLf &_
				vbTab & "strSignatureFilePath: " & strSignatureFilePath & vbCrLf)
	
	Dim strSignatureFileLines, arrSignatureFileLines
	
	If objFSO.FileExists(strSignatureFilePath) Then
	
		' https://stackoverflow.com/questions/13851473/read-utf-8-text-file-in-vbscript
		' strSignatureFileLines = objFSO.OpenTextFile(strSignatureFilePath, ForReading, TristateTrue).ReadAll()
		
		objStream.CharSet = "utf-8"
		objStream.Open
		objStream.LoadFromFile strSignatureFilePath
		strSignatureFileLines = objStream.ReadText()
		objStream.Close

		Dim arrSignatureUserVariables, strSignatureUserVariable
		arrSignatureUserVariables = Array("sAMAccountName", "DisplayName", "FirstName", "Initials", "LastName", _
												"TelephoneNumber", "OtherTelephoneNumber", "HomePhoneNumber", "OtherHomePhoneNumber", _
												"PagerNumber", "MobileNumber", "FaxNumber", "OtherFaxNumber", "Email", _
												"Street", "POBox", "City", "State", "ZipCode", "Country", "CountryAbbreviation", _
												"Title", "Department", "Manager", "Office", "Company")
		
		Dim i, arrRegExMatches, strRegExPattern
		Dim strExecuteCode
		
		' User Variables
		' -----------------------------------------------------------------------------------------------------------------------------------------
		Dim strObjUserValueName
		For Each strSignatureUserVariable In arrSignatureUserVariables
			ShowVerboseMessage(vbTab & vbTab & "strSignatureUserVariable: " & strSignatureUserVariable)
			strRegExPattern = "\%%(.*?" & strSignatureUserVariable & ".*?)\%%"
			ShowVerboseMessage(vbTab & vbTab & vbTab & "GetRegExMatches(strSignatureFileLines, " & strRegExPattern & ")")
			arrRegExMatches = RemoveDupsArray(GetRegExMatches(strSignatureFileLines, strRegExPattern))
			ShowVerboseMessage(vbTab & vbTab & vbTab & "UBound(arrRegExMatches): " & UBound(arrRegExMatches))
			For i = LBound(arrRegExMatches) To UBound(arrRegExMatches)
				ShowVerboseMessage(vbTab & vbTab & "arrRegExMatches(" & i & "): " & arrRegExMatches(i))
				
				'E.g.:
				'	strSignatureFileLines = <img width=515 height=215 src="%%Email%%.png">
				'	strRegExPattern = \%%(.*?Email.*?)\%%
				'	arrRegExMatches = Array(Email)
				'	For i = 0 To UBound(arrRegExMatches)
				'		0 =>
				'			RegExReplace(strSignatureFileLines, %%Email%%, strEmail)
				'	Next
				strExecuteCode = "strObjUserValueName = str" & strSignatureUserVariable
				ShowVerboseMessage(vbTab & vbTab & strExecuteCode)
				Execute(strExecuteCode)
				ShowVerboseMessage(vbTab & vbTab & "RegExReplace(" & vbCrLf & "" & strSignatureFileLines & """, %%" & arrRegExMatches(i) & "%%, " & strObjUserValueName & ")")
				strSignatureFileLines = RegExReplace(strSignatureFileLines, "%%" & arrRegExMatches(i) & "%%", strObjUserValueName)
				ShowVerboseMessage(vbTab & vbTab & "strSignatureFileLines: " & vbCrLf & "" & strSignatureFileLines & "")
			Next
		Next
		ShowVerboseMessage("")
		' -----------------------------------------------------------------------------------------------------------------------------------------
		
		' Date Variables
		' -----------------------------------------------------------------------------------------------------------------------------------------
		Dim arrSignatureDateVariables, strSignatureDateVariable
		arrSignatureDateVariables = Array("Month", "Day", "Year")
		
		Dim dtmDate
		For Each strSignatureDateVariable In arrSignatureDateVariables
			ShowVerboseMessage(vbTab & "Select Case " & strSignatureDateVariable)
			Select Case strSignatureDateVariable
				Case "Month"
					dtmDate = Month(Now)
				Case "Day"
					dtmDate = Day(Now)
				Case "Year"
					dtmDate = Year(Now)
			End Select
			ShowVerboseMessage(vbTab & vbTab & "dtmDate: " & dtmDate)
			
			strRegExPattern = "\%%((|\s|[0-9+-])+" & strSignatureDateVariable & "(|\s|[0-9+-])+)\%%"
			ShowVerboseMessage(vbTab & vbTab & "GetRegExMatches(strSignatureFileLines, " & strRegExPattern & ")")
			arrRegExMatches = RemoveDupsArray(GetRegExMatches(strSignatureFileLines, strRegExPattern))
			ShowVerboseMessage(vbTab & vbTab & "UBound(arrRegExMatches): " & UBound(arrRegExMatches))
			For i = LBound(arrRegExMatches) To UBound(arrRegExMatches)
				ShowVerboseMessage(vbTab & vbTab & "arrRegExMatches(" & i & "): " & arrRegExMatches(i))
							
				'E.g.:
				'	dtmDate = 2020
				'	strSignatureFileLines = #12/01/%%Year%%#;%%DayofWeek(#1/1/%{Year+1}#,2)%%;Christmas;@example.com
				'	strRegExPattern = \%%((|\s|[0-9+-])+year(|\s|[0-9+-])+)\%%
				'	arrRegExMatches = Array(Year, Year+1)
				'	For i = 0 To UBound(arrRegExMatches)
				'		0 =>
				'			RegExReplace(strSignatureFileLines, %%Year%%, RegExReplace(Year, Year, 2020))
				'				=> RegExReplace(strSignatureFileLines, %%Year%%, 2020)
				'					=> 2020
				'		1 =>
				'			RegExReplace(strSignatureFileLines, %%Year\+1%%, RegExReplace(Year\+1, Year, 2020))
				'				=> RegExReplace(strSignatureFileLines, %%Year\+1%%, 2021)
				'					=> 2021
				'	Next
				Execute("intSumNumsInStr = " & RegExReplace(arrRegExMatches(i), strSignatureDateVariable, dtmDate))
				ShowVerboseMessage(vbTab & vbTab & vbTab & "RegExReplace(strSignatureFileLines, %%" & RegExEscape(arrRegExMatches(i)) & "%%, " & intSumNumsInStr & ")")
				strSignatureFileLines = RegExReplace(strSignatureFileLines, "%%" & RegExEscape(arrRegExMatches(i)) & "%%", intSumNumsInStr)
				ShowVerboseMessage(vbTab & vbTab & vbTab & "strSignatureFileLines: " & vbCrLf & strSignatureFileLines)
			Next
			ShowVerboseMessage("")
		Next
		
		Dim dtmWeekday
		strRegExPattern = "%%Weekday\(([^)]+)\)%%"
		ShowVerboseMessage(vbTab & "GetRegExMatches(strSignatureFileLines, " & strRegExPattern & ")")
		arrRegExMatches = RemoveDupsArray(GetRegExMatches(strSignatureFileLines, strRegExPattern))
		ShowVerboseMessage(vbTab & "UBound(arrRegExMatches): " & UBound(arrRegExMatches))
		For i = LBound(arrRegExMatches) To UBound(arrRegExMatches)
			'ShowVerboseMessage(strSignatureFileLines)
			strExecuteCode = "GetWeekday " & arrRegExMatches(i) & ",dtmWeekday"
			ShowVerboseMessage(vbTab & "Execute(" & strExecuteCode & ")")
			Execute(strExecuteCode)
			ShowVerboseMessage(vbTab & "RegExReplace(strSignatureFileLines, " & "%%Weekday(" & arrRegExMatches(i) & ")%%" & ", " & dtmWeekday & ")" & vbCrLf)
			strSignatureFileLines = RegExReplace(strSignatureFileLines, RegExEscape("%%Weekday(" & arrRegExMatches(i) & ")%%"), dtmWeekday)
		Next
		' -----------------------------------------------------------------------------------------------------------------------------------------
		
		arrSignatureFileLines = RemoveDupsArray(Split(strSignatureFileLines, vbCrLf))
		
		ShowVerboseMessage(vbCrLf & vbTab & strSignatureFilePath & " contents (" & UBound(arrSignatureFileLines)+1 & " lines):" & vbCrLf &_
					vbTab & "-----------------------" & vbCrLf &_
					Join(arrSignatureFileLines, vbCrLf) & vbCrLf)
	Else
		ShowVerboseMessage(vbTab & "**ERROR** File " & strSignatureFilePath & " doesn't exists." & vbCrLf)
		ReDim arrSignatureFileLines(-1)
	End If
	
	ParseSignatureFileVariables = arrSignatureFileLines
End Function

Function IsOutlookRunning()
	ShowVerboseMessage("IsOutlookRunning function:" & vbCrLf &_
				"--------------------------")
	Dim boolOutlookIsRunning
	boolOutlookIsRunning = False
	Dim strWMIQuery, objWMIService, colProcesses, objProcess
	strWMIQuery = "Select * from Win32_Process " & _
	"Where Name = 'Outlook.exe'"
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _
	& strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery(strWMIQuery)
	For Each objProcess In colProcesses
		'If UCase(objProcess.Name) = "OUTLOOK.EXE" Then
			boolOutlookIsRunning = True
		'Else
		'	IsOutlookRunning = False
		'End If
	Next
	ShowVerboseMessage("boolIsOutlookRunning: " & boolOutlookIsRunning & vbCrLf)
	IsOutlookRunning = boolOutlookIsRunning
End Function

' https://stackoverflow.com/questions/7401168/convert-a-registry-binary-value-into-meaningful-string
Function RegBinaryToString(arrValue)
	Dim strInfo : strInfo = ""
	If IsArray(arrValue) Then
		For i=0 to UBound(arrValue)  
			If arrValue(i)<>0 Then strInfo = strInfo & Chr(arrValue(i))		
		Next
	End If
	RegBinaryToString = strInfo  
End Function  

Function StringToByteArray _
(Data, NeedNullTerminator)
	Dim strAll, intLen, arr, i
	strAll = StringToHex4(Data)
	If NeedNullTerminator Then
		strAll = strAll & "0000"
	End If
	intLen = Len(strAll) \ 2
	ReDim arr(intLen - 1)
	For i = 1 To Len(strAll) \ 2
		arr(i - 1) = CByte _
		("&H" & Mid(strAll, (2 * i) - 1, 2))
	Next
	StringToByteArray = arr
End Function

Function StringToHex4(Data)
	Dim strAll, i, strChar, strTemp
	For i = 1 To Len(Data)

		strChar = Mid(Data, i, 1)
		strTemp = Right("00" & Hex(AscW(strChar)), 4)
		strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
	Next
	StringToHex4 = strAll

End Function

Function GetOfficeVersion
	ShowDebugMessage("[debug] GetOfficeVersion function:" & vbCrLf &_
				"--------------------------")
				
	Dim strKeyOutlookAppPath, strOutlookPathValue, strOutlookVersionNumber

	'Determine path to outlook.exe
	strKeyOutlookAppPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
	objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyOutlookAppPath, "Path", strOutlookPathValue
	
	ShowDebugMessage("[debug] strKeyOutlookAppPath: HKEY_LOCAL_MACHINE\" & strKeyOutlookAppPath & vbCrLf &_
					"strOutlookPathValue: " & strOutlookPathValue)
	
	'Verify that the outlook.exe exist and get version information
	If (Not IsNull(strOutlookPathValue)) Then
		If objFSO.FileExists(objFSO.BuildPath(strOutlookPathValue, "OUTLOOK.exe")) Then
			strOutlookVersionNumber = objFSO.GetFileVersion(objFSO.BuildPath(strOutlookPathValue, "OUTLOOK.exe"))
			ShowDebugMessage("[debug] strOutlookVersionNumber: " & strOutlookVersionNumber & vbCrLf)
			GetOfficeVersion = Left(strOutlookVersionNumber, InStr(strOutlookVersionNumber, ".0")+1)
		End If
	End If
End Function

Function FirstVersionSupOrEqualToSecondVersion(strFirstVersion, strSecondVersion)
	
	Dim arrFirstVersion,  arrSecondVersion, i, iStop, iMax
	Dim iFirstArraySize, iSecondArraySize
	Dim blnArraySameSize : blnArraySameSize = False
	
	If strFirstVersion = strSecondVersion Then
		FirstVersionSupOrEqualToSecondVersion = True
		Exit Function
	End If
	
	If strFirstVersion = "" Then
		FirstVersionSupOrEqualToSecondVersion = False
		Exit Function
	End If
	If strSecondVersion = "" Then
		FirstVersionSupOrEqualToSecondVersion = True
		Exit Function
	End If

	arrFirstVersion = Split(strFirstVersion, "." )
	arrSecondVersion = Split(strSecondVersion, "." )
	iFirstArraySize = UBound(arrFirstVersion)
	iSecondArraySize = UBound(arrSecondVersion)
	
	If iFirstArraySize = iSecondArraySize Then
		blnArraySameSize = True
		iStop = iFirstArraySize
		For i=0 To iStop
			If CInt(arrFirstVersion(i)) < CInt(arrSecondVersion(i)) Then
				FirstVersionSupOrEqualToSecondVersion = False
				Exit Function
			End If
		Next
		FirstVersionSupOrEqualToSecondVersion = True
	Else
		If iFirstArraySize > iSecondArraySize Then
			iStop = iSecondArraySize
		Else
			iStop = iFirstArraySize
		End If
		For i=0 To iStop
			If CInt(arrFirstVersion(i)) < CInt(arrSecondVersion(i)) Then
				FirstVersionSupOrEqualToSecondVersion = False
				Exit Function
			End If
		Next
		If iFirstArraySize > iSecondArraySize Then
			FirstVersionSupOrEqualToSecondVersion = True
			Exit Function
		Else
			For i=iStop+1 To iSecondArraySize
				If CInt(arrSecondVersion(i)) > 0 Then
					FirstVersionSupOrEqualToSecondVersion = False
					Exit Function
				End If
			Next
			FirstVersionSupOrEqualToSecondVersion = True
		End If
	End If
End Function

Function GetFirstFreeDriveLetter
 
    Dim objFSO, strLetters, i,  blnError 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
     
    '* list of possible drive letters 
    '* A and B are reserved for floppy disc 
    '* you may limit the search using any subset of the alphabet 
    strLetters = "CDEFGHIJKLMNOPQRSTUVWXYZ"  
    GetFirstFreeDriveLetter = "" 
    blnError = True 
     
    '* walk through all possible drive letters 
    For i=1 to len(strLetters) 
    '* if the drive letter isn't in use the it's ours 
        If not objFSO.DriveExists(mid(strLetters, i, 1) & ":") Then 
            '* we have found a free drive letter, therefore blnError = False 
            blnError = False 
            '* assigning the return value 
            GetFirstFreeDriveLetter = mid(strLetters, i, 1) & ":" 
            '* we want to find the FIRST free drive letter 
            Exit For 
        End If 
    Next  
     
    '* error handling 
    If blnError then  
        ShowVerboseMessage("Error - no free drive letter found!")		
		WScript.Quit(1)
    End If 
     
    '* releasing file system object 
    Set objFSO = Nothing 
 
End Function 

Function FirstFileAreNewer(strF1, strF2)
	ShowVerboseMessage(vbTab & "FirstFileAreNewer function:" & vbCrLf &_
				vbTab & "---------------------------" & vbCrLf &_
				vbTab & "strF1: " & strF1 & vbCrLf &_
				vbTab & "strF2: " & strF2)
	
	' cmd = "%COMSPEC% /c fc /b " & qq(f1) & " " & qq(f2)
	' FirstFileAreNewer = CBool(CreateObject("WScript.Shell").Run(cmd, 0, True))
	
	Dim objWMIService, strWMIQuery
	Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
	
	strWMIQuery = "Select * from CIM_DataFile where Name = '" & Replace(strF1, "\", "\\") & "'"
	ShowDebugMessage(vbTab & "[debug] strWMIQuery: " & strWMIQuery)
	Dim colFilesF1, objFileF1
	Dim arrFileDateLastModifiedF1, dtmFileDateLastModifiedF1
	Set colFilesF1 = objWMIService.ExecQuery (strWMIQuery)
	For Each objFileF1 in colFilesF1
		' WMI queries use a date and time format called CIM_DATETIME (yyyymmddHHMMSS.mmmmmmsUUU).
		ShowDebugMessage(vbTab & "[debug] objFileF1.LastModified: " & objFileF1.LastModified)
		' dtmFileDateLastModifiedF1 = MM/DD/YYYY HH:MM:SS
		dtmFileDateLastModifiedF1 = RegExReplace(Left(objFileF1.LastModified, InStr(objFileF1.LastModified, ".")-1), _
													"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", "$2/$3/$1 $4:$5:$6")
		If Len(dtmFileDateLastModifiedF1) > 0 Then
			dtmFileDateLastModifiedF1 = CDate(dtmFileDateLastModifiedF1)
		End If
	Next
	ShowVerboseMessage(vbTab & "dtmFileDateLastModifiedF1: " & dtmFileDateLastModifiedF1)
	
	strWMIQuery = "Select * from CIM_DataFile where Name = '" & Replace(strF2, "\", "\\") & "'"
	ShowDebugMessage(vbTab & "[debug] strWMIQuery: " & strWMIQuery)
	Dim colFilesF2, objFileF2
	Dim arrFileDateLastModifiedF2, dtmFileDateLastModifiedF2
	Set colFilesF2 = objWMIService.ExecQuery (strWMIQuery)
	For Each objFileF2 in colFilesF2
		' WMI queries use a date and time format called CIM_DATETIME (yyyymmddHHMMSS.mmmmmmsUUU).
		ShowDebugMessage(vbTab & "[debug] objFileF2.LastModified: " & objFileF2.LastModified)
		' dtmFileDateLastModifiedF2 = MM/DD/YYYY HH:MM:SS
		dtmFileDateLastModifiedF2 = RegExReplace(Left(objFileF2.LastModified, InStr(objFileF2.LastModified, ".")-1), _
													"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", "$2/$3/$1 $4:$5:$6")
		If Len(dtmFileDateLastModifiedF2) > 0 Then
			dtmFileDateLastModifiedF2 = CDate(dtmFileDateLastModifiedF2)
		End If
	Next
	ShowVerboseMessage(vbTab & "dtmFileDateLastModifiedF2: " & dtmFileDateLastModifiedF2)
	
	If dtmFileDateLastModifiedF1 > dtmFileDateLastModifiedF2 Then
		FirstFileAreNewer = True
		ShowVerboseMessage(vbTab & "FirstFileAreNewer: " & True & vbCrLf)
	Else
		FirstFileAreNewer = False
		ShowVerboseMessage(vbTab & "FirstFileAreNewer: " & False & vbCrLf)
	End If
End Function

Function GetRegExMatches(inputString, strPattern)
	ShowDebugMessage("[debug] GetRegExMatches function:" & vbCrLf &_
					"-------------------------" & vbCrLf &_
					"inputString: " & inputString & vbCrLf &_
					"strPattern: " & strPattern)
	
    Dim regEx_, colMatches_, arrMatches_, objMatch_
    Set regEx_ = New RegExp
	
	With regEx_
		.Global = True
		.MultiLine = True
		.IgnoreCase = True
	End With
	
	If Len(inputString) > 0 Then
		regEx_.Pattern = strPattern
		ShowDebugMessage("[debug] regEx_.Test(" & inputString & "): " & regEx_.Test(inputString))
		
		Dim i
		i = 0
		If regEx_.Test(inputString) Then
			Set colMatches_ = regEx_.Execute(inputString)
			ReDim arrMatches_(colMatches_.Count-1)
			ShowDebugMessage("[debug] colMatches_.Count: " & colMatches_.Count & vbCrLf &_
						  "UBound(arrMatches_): " & UBound(arrMatches_))
			For Each objMatch_ In colMatches_
				ShowDebugMessage("[debug] objMatch_.SubMatches(0): " & objMatch_.SubMatches(0))
				arrMatches_(i) = objMatch_.SubMatches(0)
				i = i + 1
			Next
		Else
			ReDim arrMatches_(-1)
		End If
	Else
		ReDim arrMatches_(-1)
	End If
	GetRegExMatches = arrMatches_
End Function

Function RegExReplace(inputString, strRegExPattern, strReplacePattern)
	ShowDebugMessage(vbCrLf & "[debug] RegExReplace function:" & vbCrLf &_
				"----------------------" & vbCrLf &_
				"inputString: " & inputString & vbCrLf &_
				"strRegExPattern: " & strRegExPattern & vbCrLf &_
				"strReplacePattern: " & strReplacePattern)
				
	Dim regEx_
	Set regEx_ = New RegExp
	With regEx_
		.Global = True
		.MultiLine = True
		.IgnoreCase = True
		.Pattern = strRegExPattern
	End With
	ShowDebugMessage("[debug] regEx_.Test(" & inputString & "): " & regEx_.Test(inputString) & vbCrLf)
	RegExReplace = regEx_.Replace(inputString, strReplacePattern)
End Function

Function RegExEscape(inputString)
	' https://stackoverflow.com/questions/50440808/excel-vba-regex-escape-method
	Dim regEx_
	Set regEx_ = New RegExp
	With regEx_
		.Global = True
		.MultiLine = True
		.IgnoreCase = True
		.Pattern = "[-/\\^$*+?.()|[\]{}]"
	End With
	RegExEscape = regEx_.Replace(inputString, "\$&")
End Function

Function RemoveTrailingBackslash(strString)
	If Right(strString, 1) = "\" Then
		RemoveTrailingBackslash = Left(strString, Len(strString)-1)
	Else
		RemoveTrailingBackslash = strString
	End If
End Function

Function RemoveDupsArray(arrItems)
	' https://stackoverflow.com/questions/6591655/remove-double-values-from-array-classic-asp
	Dim objDictionary, strItem, intItems, strKey
	Set objDictionary = CreateObject("Scripting.Dictionary")

	For Each strItem in arrItems
	  If Not objDictionary.Exists(strItem) Then
		  objDictionary.Add strItem, strItem   
	  End If
	Next

	intItems = objDictionary.Count - 1

	ReDim arrItems(intItems)

	i = 0

	For Each strKey in objDictionary.Keys
	   arrItems(i) = strKey
	   i = i + 1
	Next

	RemoveDupsArray = arrItems
End Function

Function GetWeekday(dtmDate, intDayOfWeek, strFormatDateTime, ByRef outputString)
	' https://devblogs.microsoft.com/scripting/how-can-i-determine-the-day-of-the-week/
	' https://www.w3schools.com/asp/func_weekday.asp
	ShowVerboseMessage(vbTab & vbTab & vbTab & "GetWeekday function:" & vbCrLf &_
				 vbTab & vbTab & vbTab & "----------------------" & vbCrLf &_
				 vbTab & vbTab & vbTab & "dtmDate: " & dtmDate & vbCrLf &_
				 vbTab & vbTab & vbTab & "intDayOfWeek: " & WeekdayName(intDayOfWeek) & " (" & intDayOfWeek & ")" & vbCrLf &_
				 vbTab & vbTab & vbTab & "strFormatDateTime: " & strFormatDateTime)
	Dim x
	Do Until x = 1
		If Weekday(dtmDate) = intDayOfWeek Then
			ShowVerboseMessage(vbTab & vbTab & vbTab & "The first " & WeekdayName(intDayOfWeek) & " of the week is """ & dtmDate & """.")
			dtmDate = RegExReplace(Month(dtmDate) & "/" & Day(dtmDate) & "/" & Year(dtmDate), "^(\d{1,2})\/(\d{1,2})\/(\d{4})$", strFormatDateTime)
			' https://stackoverflow.com/questions/19829002/how-to-add-zero-in-front-of-single-digit-values-using-regex-in-pentaho
			dtmDate = RegExReplace(dtmDate, "\b([0-9])\b", "0$1")
			Exit Do
		Else
			dtmDate = dtmDate + 1
		End If
	Loop
	outputString = dtmDate
	GetWeekday = dtmDate
End Function

'------------------------------------------------------------------
' Used to get the time for the trigger 
' startBoundary and endBoundary.
' Return the time in the correct format: 
' YYYY-MM-DDTHH:MM:SS. 
' https://docs.microsoft.com/en-us/windows/win32/taskschd/time-trigger-example--scripting-
'------------------------------------------------------------------
Function XmlTime(t)
    Dim cSecond, cMinute, CHour, cDay, cMonth, cYear
    Dim tTime, tDate

    cSecond = "0" & Second(t)
    cMinute = "0" & Minute(t)
    cHour = "0" & Hour(t)
    cDay = "0" & Day(t)
    cMonth = "0" & Month(t)
    cYear = Year(t)

    tTime = Right(cHour, 2) & ":" & Right(cMinute, 2) & _
        ":" & Right(cSecond, 2)
    tDate = cYear & "-" & Right(cMonth, 2) & "-" & Right(cDay, 2)
    XmlTime = tDate & "T" & tTime 
End Function

'*************************************************************************
' Name: ConvertToUTCTime ( sTime as Date )
' Parameters: Takes the Local Time. eg., 11/07/2006 10:00:00 AM
' Returns the converted UTC Time as 2006-11-07T18:00:000Z
'
' Written by : Anand Venkatachalapathy
' URL        : https://anandthearchitect.com/2006/11/08/convert-utc-to-local-time-and-vice-versa-using-vbscript/
'*************************************************************************
Function ConvertToUTCTime(sTime)

	Dim od, ad, oShell, atb, offsetMin
	Dim sHour, sMinute, sMonth, sDay

	' od = sTime
	' If you passed sTime as string, comment the above line and
	' uncomment the below line.
	od = CDate(sTime)

	' Create Shell object to read registry
	Set oShell = CreateObject("WScript.Shell")

	atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\" & _
	"Control\TimeZoneInformation\ActiveTimeBias"
	offsetMin = oShell.RegRead(atb) ' Reading the registry

	' Convert the local time to UTC time
	ad = DateAdd("n", offsetMin, od)

	' If Month is single digit value, add zero
	sMonth = Month(CDate(ad))
	If Len(sMonth) = 1 Then
				sMonth = "0" & sMonth
	End If

	' If Day is single digit, add zero
	sDay = Day(CDate(ad))
	If Len(sDay) = 1 Then
			  sDay = "0" & sDay
	End If

	' If Hour is single digit, add zero
	sHour = Hour(CDate(ad))
	If Len(sHour) = 1 Then
			  sHour = "0" & sHour
	End If

	' If Minute is single digit, add zero
	sMinute = Minute(CDate(ad))
	If Len(sMinute) = 1 Then
			 sMinute = "0" & sMinute
	End If

	' Assign the reutrn value in UTC format as 2006-11-07T18:00:000Z
	' ConvertToUTCTime = Year(CDate(ad)) & "-" & _
	' sMonth & "-" & _
	' sDay & "T" & _
	' sHour & ":" & _
	' sMinute & ":00Z"

	ConvertToUTCTime = ad

End Function
' End of Function ConvertToUTCTime

Function ShowVerboseMessage(strMessage)
   If boolVerbose Then
      WScript.Echo strMessage
   End If
End Function

Function ShowDebugMessage(strMessage)
   If boolDebug Then
      WScript.Echo strMessage
   End If
End Function
