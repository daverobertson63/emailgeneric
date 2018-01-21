Private bSectionExists         As Boolean
Private bKeyExists             As Boolean
 
'---------------------------------------------------------------------------------------
' Procedure : Ini_ReadKeyVal
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Read an Ini file's Key
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'             No APIs either! 100% VBA
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile  : Full path and filename of the ini file to read
' sSection  : Ini Section to search for the Key to read the Key from
' sKey      : Name of the Key to read the value of
'
' Usage:
' ~~~~~~
' ? Ini_Read(Application.CurrentProject.Path & "\MyIniFile.ini", "LINKED TABLES", "Path")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-08-09              Initial Release
'---------------------------------------------------------------------------------------
Function Ini_ReadKeyVal(ByVal sIniFile As String, _
                        ByVal sSection As String, _
                        ByVal sKey As String) As String
    On Error GoTo Error_Handler
    Dim sIniFileContent       As String
    Dim aIniLines()           As String
    Dim sLine                 As String
    Dim i                     As Long
 
    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False
 
    'Validate that the file actually exists
    If FileExist(sIniFile) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFile & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If
 
    sIniFileContent = ReadFile(sIniFile)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)
    For i = 0 To UBound(aIniLines)
        sLine = Trim(aIniLines(i))
        If bSectionExists = True And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
            Exit For    'Start of a new section
        End If
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
        End If
        If bSectionExists = True Then
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    bKeyExists = True
                    Ini_ReadKeyVal = Mid(sLine, InStr(sLine, "=") + 1)
                End If
            End If
        End If
    Next i
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    'Err.Number = 75 'File does not exist, Permission issues to write is denied,
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Ini_ReadKeyVal" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function
 
'---------------------------------------------------------------------------------------
' Procedure : Ini_WriteKeyVal
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Writes a Key value to the specified Ini file's Section
'               If the file does not exist, it will be created
'               If the Section does not exist, it will be appended to the existing content
'               If the Key does not exist, it will be appended to the existing Section content
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'             No APIs either! 100% VBA
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile  : Full path and filename of the ini file to edit
' sSection  : Ini Section to search for the Key to edit
' sKey      : Name of the Key to edit
' sValue    : Value to associate to the Key
'
' Usage:
' ~~~~~~
' Call Ini_WriteKeyVal(Application.CurrentProject.Path & "\MyIniFile.ini", "LINKED TABLES", "Paths", "D:\")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-08-09              Initial Release
'---------------------------------------------------------------------------------------
Function Ini_WriteKeyVal(ByVal sIniFile As String, _
                         ByVal sSection As String, _
                         ByVal sKey As String, _
                         ByVal sValue As String) As Boolean
    On Error GoTo Error_Handler
    Dim sIniFileContent       As String
    Dim aIniLines()           As String
    Dim sLine                 As String
    Dim sNewLine              As String
    Dim i                     As Long
    Dim bFileExist            As Boolean
    Dim bInSection            As Boolean
    Dim bKeyAdded             As Boolean
 
    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False
 
    'Validate that the file actually exists
    If FileExist(sIniFile) = False Then
        GoTo SectionDoesNotExist
    End If
    bFileExist = True
 
    sIniFileContent = ReadFile(sIniFile)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)    'Break the content into individual lines
    sIniFileContent = ""    'Reset it
    For i = 0 To UBound(aIniLines)    'Loop through each line
        sNewLine = ""
        sLine = Trim(aIniLines(i))
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
            bInSection = True
        End If
        If bInSection = True Then
            If sLine <> "[" & sSection & "]" _
               And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                'Our section exists, but the key wasn't found, so append it
                bInSection = False    ' we're switching section
            End If
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    sNewLine = sKey & "=" & sValue
                    bKeyExists = True
                    bKeyAdded = True
                End If
            End If
        End If
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        If sNewLine = "" Then
            sIniFileContent = sIniFileContent & sLine
        Else
            sIniFileContent = sIniFileContent & sNewLine
        End If
    Next i
 
SectionDoesNotExist:
    'if not found, add it to the end
    If bSectionExists = False Then
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        sIniFileContent = sIniFileContent & "[" & sSection & "]"
    End If
    If bKeyAdded = False Then
        sIniFileContent = sIniFileContent & vbCrLf & sKey & "=" & sValue
    End If
 
    'Write to the ini file the new content
    Call OverwriteTxt(sIniFile, sIniFileContent)
    Ini_WriteKeyVal = True
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Ini_WriteKeyVal" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileExist
' DateTime  : 2007-Mar-06 13:51
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Test for the existance of a file; Returns True/False
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - name of the file to be tested for including full path
'---------------------------------------------------------------------------------------
Function FileExist(strFile As String) As Boolean
    On Error GoTo Err_Handler
 
    FileExist = False
    If Len(Dir(strFile)) > 0 Then
        FileExist = True
    End If
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FileExist" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler
End Function
 
'---------------------------------------------------------------------------------------
' Procedure : OverwriteTxt
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Output Data to an external file (*.txt or other format)
'             ***Do not forget about access' DoCmd.OutputTo Method for
'             exporting objects (queries, report,...)***
'             Will overwirte any data if the file already exists
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - name of the file that the text is to be output to including the full path
' sText - text to be output to the file
'
' Usage:
' ~~~~~~
' Call OverwriteTxt("C:\Users\Vance\Documents\EmailExp2.txt", "Text2Export")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-Jul-06                 Initial Release
'---------------------------------------------------------------------------------------
Function OverwriteTxt(sFile As String, sText As String)
On Error GoTo Err_Handler
    Dim FileNumber As Integer
 
    FileNumber = FreeFile                   ' Get unused file number
    Open sFile For Output As #FileNumber    ' Connect to the file
    Print #FileNumber, sText;                ' Append our string
    Close #FileNumber                       ' Close the file
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: OverwriteTxt" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler
End Function
 
'---------------------------------------------------------------------------------------
' Procedure : ReadFile
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Faster way to read text file all in RAM rather than line by line
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - name of the file that is to be read
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' MyTxt = ReadText("c:\tmp\test.txt")
' MyTxt = ReadText("c:\tmp\test.sql")
' MyTxt = ReadText("c:\tmp\test.csv")
'---------------------------------------------------------------------------------------
Function ReadFile(ByVal strFile As String) As String
On Error GoTo Error_Handler
    Dim FileNumber  As Integer
    Dim sFile       As String 'Variable contain file content
 
    FileNumber = FreeFile
    Open strFile For Binary Access Read As FileNumber
    sFile = Space(LOF(FileNumber))
    Get #FileNumber, , sFile
    Close FileNumber
 
    ReadFile = sFile
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: ReadFile" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Sub TestReadKey()
    MsgBox "INI File: " & Application.CurrentProject.Path & "\MyIniFile.ini" & vbCrLf & _
           "Section: SETTINGS" & vbCrLf & _
           "Section Exist: " & bSectionExists & vbCrLf & _
           "Key: License" & vbCrLf & _
           "Key Exist: " & bKeyExists & vbCrLf & _
           "Key Value: " & Ini_ReadKeyVal(Application.CurrentProject.Path & "\MyIniFile.ini", "SETTINGS", "License")
    'You can validate the value by checking the bSectionExists and bKeyExists variable to ensure they were actually found in the ini file
End Sub
 
Sub TestWriteKey()
    If Ini_WriteKeyVal(Application.CurrentProject.Path & "\MyIniFile.ini", "SETTINGS", "License", "JBXR-HHTY-LKIP-HJNB-GGGT") = True Then
        MsgBox "The key was written"
    Else
        MsgBox "An error occured!"
    End If
End Sub

Sub Main

    print Ini_ReadKeyVal("C:\Documentum\dba\config\EPAProd\edms.ini", "EMAIL", "AERGeneric")

End Sub