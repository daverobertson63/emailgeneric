Option Explicit

Dim goFS : Set goFS = CreateObject("Scripting.FileSystemObject")

'WScript.Quit demoReadFile()
WScript.Quit demoReadIniFile()

Function demoReadFile()
  demoReadFile = 0
  Dim tsIn : Set tsIn = goFS.OpenTextFile("C:\Documentum\dba\config\EPAProd\edms.ini")
  Do Until tsIn.AtEndOfStream
     Dim sLine : sLine = tsIn.ReadLine()
     WScript.Echo tsIn.Line - 1, sLine
  Loop
  tsIn.Close
End Function

Function demoReadIniFile()
  demoReadIniFile = 0
  Dim dicIni : Set dicIni = ReadIniFile("C:\Documentum\dba\config\EPAProd\edms.ini")
  Dim sSec, sKV
  For Each sSec In dicIni.Keys()
      WScript.Echo "---", sSec
      For Each sKV In dicIni(sSec).Keys()
          WScript.Echo " ", sKV, "=>", dicIni(sSec)(sKV)
      Next
  Next
  WScript.Echo dicIni("owner")("name")
End Function

Function ReadIniFile(sFSpec)
  Dim dicTmp : Set dicTmp = CreateObject("Scripting.Dictionary")
  Dim tsIn   : Set tsIn   = goFS.OpenTextFile(sFSpec)
  Dim sLine, sSec, aKV
  Do Until tsIn.AtEndOfStream
     sLine = Trim(tsIn.ReadLine())
     If "[" = Left(sLine, 1) Then
        sSec = Mid(sLine, 2, Len(sLine) - 2)
        Set dicTmp(sSEc) = CreateObject("Scripting.Dictionary")
     Else
        If "" <> sLine Then
           aKV = Split(sLine, "=")
           If 1 = UBound(aKV) Then
              dicTmp(sSec)(Trim(aKV(0))) = Trim(aKV(1))
           End If
        End If
     End If
  Loop
  tsIn.Close
  Set ReadIniFile = dicTmp
End Function