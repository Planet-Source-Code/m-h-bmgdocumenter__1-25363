Attribute VB_Name = "modUtility"
Option Explicit

Public Function GetProjectPath(ProjectFile As String) As String
   Dim strProjectPath As String
   Dim lngEndOfPath As String
   
   If FileExists(ProjectFile) Then
      'Want to find last '\' -> reverse string and look for first!
      lngEndOfPath = Len(ProjectFile) - InStr(1, StrReverse(ProjectFile), "\")
      
      'The above math trims the '\', so we add it back in (always consistent now!)
      GetProjectPath = Left$(ProjectFile, lngEndOfPath) & "\"
   Else
      GetProjectPath = ""
   End If
   
   GetProjectPath = UCase$(GetProjectPath)
End Function

Public Function FileExists(FileName As String) As Boolean
   Dim lngFileHandle As Long
   
   lngFileHandle = FreeFile
   
   'Check to ensure project is still there, and available...
   On Error GoTo FileError
   Open FileName For Input As #lngFileHandle
   Close #lngFileHandle
   On Error GoTo 0
   
   FileExists = True
   
   Exit Function
   
FileError:
   FileExists = False
End Function

Public Function GetCompleteLine(FileNo As Long) As String
   Dim strTemp As String
   Dim strBuffer As String
   
   Do
      Line Input #FileNo, strBuffer
      
      strBuffer = Trim$(strBuffer)
      
      'Strip
      If Right$(strBuffer, 1) = "_" Then
         strTemp = strTemp & Left$(strBuffer, Len(strBuffer) - 1)
      Else
         strTemp = strTemp & strBuffer
      End If
   Loop While Right$(strBuffer, 1) = "_"
   
   'Return rejoined line...
   GetCompleteLine = strTemp
End Function
