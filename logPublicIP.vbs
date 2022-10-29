'2021-02-03 Sofia Engvall
'Adds one line containing date, time and external IP to the textfile
' strLogFileName when the script is run
'Script can be run at set times by starting it with Task Sceduler

Option Explicit

Const ForReading = 1, ForAppending = 8

Dim strURLIP : strURLIP = "http://checkip.amazonaws.com/"
Dim strLogFileName : strLogFileName = "publicIP.txt"

'Connect to Internet URL
Dim objXMLHTTP : Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.open "GET", strURLIP, false
objXMLHTTP.send()

'Exit if no result
If objXMLHTTP.Status <> 200 Then WScript.Quit(False)

'Append date and IP to text file
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim logFile : Set logFile = objFSO.OpenTextFile(strLogFileName, ForAppending, True)
logFile.Write CStr(Now) + " " + SimpleBinaryToString(objXMLHTTP.responseBody)
logFile.Close

'------------------------------------------------------------------------
'Finish up
Set objXMLHTTP = Nothing
Set objFSO = Nothing

'------------------------------------------------------------------------
'Conversion function for .responceBody conversion to string - from motobit.com
Function SimpleBinaryToString(Binary)
  'SimpleBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using MultiByte VBS functions
  Dim I, S 
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  SimpleBinaryToString = S
End Function
