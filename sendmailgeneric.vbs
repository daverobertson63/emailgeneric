
'*********************************************************************
'Created by : Shuja
'Description : Reads and Writes to the INI file using the API calls
'For : A dude on Codeguru
'Creation Date : 24-03-2005
'*********************************************************************


'Option Explicit

Const INIFileLocation = "C:\Documentum\dba\config\EPAProd\edms.ini"

'API Function to read information from INI File
Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

'API Function to write information to the INI File
Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, sINIFileName) As String
    Const cparmLen = 1024
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, sKey _
            , sDefault, sReturn, cparmLen, sINIFileName)
    GetINISetting = Mid(sReturn, 1, lLength)
End Function

'Save INI Setting in the File
Public Function PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, sINIFileName) As Boolean
    On Error GoTo HandleError
    Const cparmLen = 1024
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim aLength As Long
    aLength = WritePrivateProfileString(sHeading, sKey _
            , sSetting, sINIFileName)
    PutINISetting = True
    Exit Function
    
HandleError:
    
    Print Err.Number & " " & Err.Description
End Function



' Return the recipient part
Function EmailRecipient(emailAddress as String) as String

	Dim p as integer
	
	p = InStr(emailAddress,"@")
	
	If p > 0 Then
		EmailRecipient = Mid(emailAddress,1,p-1)
	Else
		EmailRecipient = ""
	End If
	

End Function


' Return the domain adress part
Function EmailDomain(emailAddress as String) as String

	Dim p as integer
	
	p = InStr(emailAddress,"@")
	
	If p > 0 Then
		EmailDomain = Mid(emailAddress,p+1)
	Else
		EmailDomain = ""
	End If
	

End Function

Function getGenericAEREmailAddress(emailAddress as String)

	Dim sRecipient as String
	Dim sDomain as String
	Dim sEmailGeneric as String
	
	sDomain = EmailDomain(emailAddress)
	sRecipient = EmailRecipient(emailAddress)
	getGenericAEREmailAddress =""
	
	print sDomain
	print sRecipient
	
	' Search the list of avaialble generics and return if the domain matches
	
	sEmailGeneric = GetINISetting("EMAIL", "AERGeneric", INIFileLocation)
	
	print itemcount(sEmailGeneric)
	print sEmailGeneric
	
	print item$(sEmailGeneric,1,1)
	print item$(sEmailGeneric,2,2)
	print item$(sEmailGeneric,3)
	
	For i = 1 To itemcount(sEmailGeneric)
		print item$(sEmailGeneric,i,i,",")
			If sDomain = EmailDomain(emailAddress) then
				getGenericAEREmailAddress = item$(sEmailGeneric,i,i,",")
				print "Found a match of generic email address"
				Exit Function
			End If
	Next i

	
	
	
	


End Function

Sub Main()
    
    Dim sSetting As String
    Dim genericAddress As String
    
    genericAddress = getGenericAEREmailAddress("dave.robertson@water.ie")
    
    print genericAddress
    
    'Reads a INI File (SETTINGS.INI) which has SECTION (SQLSERVER) and HEADING (SERVER) in It
    'sSetting = GetINISetting("EMAIL", "AERGeneric", INIFileLocation)
    
    print sSetting
    
    'Change the above setting to this one
    'PutINISetting "SQLSERVER", "SERVER", "MyNewSQLServer", App.Path & "\SETTINGS.INI"
End Sub
