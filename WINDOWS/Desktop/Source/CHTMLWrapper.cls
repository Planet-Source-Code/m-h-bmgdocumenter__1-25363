VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************
' CHTMLWrapper
' by Matthew Hickson (BMG)
' written 03/16/2001
' updated --/--/----
'
' Purpose:
' This module will house wrapper functions for HTML output.  This
' will allow us to create HTML documents (for reporting, etc.)
' more easily.
'******************************************************************

Private strFileName As String
Private lngFileHandle As Long
Private boolDocOpen As Boolean

'******************************************************************
' Public Subs / Functions
'******************************************************************
Public Function OpenHTMLDoc(Optional pFileName As String = "") As Boolean
   OpenHTMLDoc = False
   
   'Check to see if the filename was already set (via property)
   If pFileName = "" Then pFileName = strFileName
   
   'If filename still doesn't exist, then exit (no error yet)
   If pFileName = "" Then Exit Function
   
   'Physically open file...
   lngFileHandle = FreeFile
   Open pFileName For Output As #lngFileHandle
   
   Print #lngFileHandle, "<HTML>"

   OpenHTMLDoc = True
   boolDocOpen = True
End Function

Public Function OpenHTMLHeader(pPageTitle As String) As Boolean
   OpenHTMLHeader = False

   If boolDocOpen Then
      Print #lngFileHandle, "<HEAD>"
      Print #lngFileHandle, "<TITLE>" & pPageTitle & "</TITLE>"
      
      OpenHTMLHeader = True
   End If
End Function

Public Function CloseHTMLHeader() As Boolean
   CloseHTMLHeader = False

   If boolDocOpen Then
      Print #lngFileHandle, "</HEAD>"
      
      CloseHTMLHeader = True
   End If
End Function

Public Function OpenHTMLBody() As Boolean
   OpenHTMLBody = False

   If boolDocOpen Then
      Print #lngFileHandle, "<BODY>"
      
      OpenHTMLBody = True
   End If
End Function

Public Function CloseHTMLBody() As Boolean
   CloseHTMLBody = False

   If boolDocOpen Then
      Print #lngFileHandle, "</BODY>"
      
      CloseHTMLBody = True
   End If
End Function

Public Function CloseHTMLDoc() As Boolean
   CloseHTMLDoc = False

   If boolDocOpen Then
      Print #lngFileHandle, "</HTML>"
      Close #lngFileHandle
      
      CloseHTMLDoc = True
      boolDocOpen = False
   End If
End Function

Public Function StartList(pIsOrdered As Boolean) As Boolean
   StartList = False
   
   If boolDocOpen Then
      If pIsOrdered Then
         Print #lngFileHandle, "<OL>"
      Else
         Print #lngFileHandle, "<UL>"
      End If
   
      StartList = True
   End If
End Function

Public Function EndList(pIsOrdered As Boolean) As Boolean
   EndList = False
   
   If boolDocOpen Then
      If pIsOrdered Then
         Print #lngFileHandle, "</OL>"
      Else
         Print #lngFileHandle, "</UL>"
      End If
   
      EndList = True
   End If
End Function

Public Function WriteListItem(pText As String) As Boolean
   WriteListItem = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<LI>" & pText & "</LI>"
   
      WriteListItem = True
   End If
End Function

Public Function StartTable(pBorder As Integer) As Boolean
   StartTable = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<TABLE BORDER=" & pBorder & ">"
   
      StartTable = True
   End If
End Function

Public Function EndTable() As Boolean
   EndTable = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "</TABLE>"
   
      EndTable = True
   End If
End Function

Public Function StartTableRow() As Boolean
   StartTableRow = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<TR>"
   
      StartTableRow = True
   End If
End Function

Public Function EndTableRow() As Boolean
   EndTableRow = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "</TR>"
   
      EndTableRow = True
   End If
End Function

Public Function WriteTableHeader(pText As String) As Boolean
   WriteTableHeader = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<TH>" & pText & "</TH>"
   
      WriteTableHeader = True
   End If
End Function

Public Function WriteTableData(pText As String) As Boolean
   WriteTableData = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<TD>" & pText & "</TD>"
   
      WriteTableData = True
   End If
End Function

Public Function WriteParagraph(pText As String) As Boolean
   WriteParagraph = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<P>" & pText & "</P>"
   
      WriteParagraph = True
   End If
End Function

Public Function WriteBoldText(pText As String) As Boolean
   WriteBoldText = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<B>" & pText & "</B>"
   
      WriteBoldText = True
   End If
End Function

Public Function WriteItalicText(pText As String) As Boolean
   WriteItalicText = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<I>" & pText & "</I>"
   
      WriteItalicText = True
   End If
End Function

Public Function WriteNormalText(pText As String) As Boolean
   WriteNormalText = False
   
   If boolDocOpen Then
      Print #lngFileHandle, pText
   
      WriteNormalText = True
   End If
End Function

Public Function WriteHeader(pLevel As Integer, pText As String) As Boolean
   WriteHeader = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<H" & pLevel & ">" & pText & "</H" & pLevel & ">"
   
      WriteHeader = True
   End If
End Function

Public Function InsertHorizontalRule() As Boolean
   InsertHorizontalRule = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<HR>"
   
      InsertHorizontalRule = True
   End If
End Function

Public Function InsertBreak() As Boolean
   InsertBreak = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<BR>"
   
      InsertBreak = True
   End If
End Function

Public Function InsertLink(pLink As String, pDescription As String) As Boolean
   InsertLink = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<A HREF=" & pLink & ">" & pDescription & "</A>"
   
      InsertLink = True
   End If
End Function

Public Function InsertPicture(pSource As String, pDescription As String) As Boolean
   InsertPicture = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<IMG SRC=" & pSource & " ALT=" & pDescription & ">"
   
      InsertPicture = True
   End If
End Function

Public Function InsertJavaScript(pScript As String, pNoScriptMsg As String) As Boolean
   InsertJavaScript = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<SCRIPT><!--"
      Print #lngFileHandle, pScript
      Print #lngFileHandle, "//--></SCRIPT>"
      Print #lngFileHandle, "<NOSCRIPT>"
      Print #lngFileHandle, pNoScriptMsg
      Print #lngFileHandle, "</NOSCRIPT>"
   
      InsertJavaScript = True
   End If
End Function

Public Function InsertComment(pText As String) As Boolean
   InsertComment = False
   
   If boolDocOpen Then
      Print #lngFileHandle, "<!-- " & pText & " -->"
   
      InsertComment = True
   End If
End Function

'******************************************************************
' Properties
'******************************************************************
Public Property Let FileName(pFileName As String)
   strFileName = pFileName
End Property

Public Property Get FileName() As String
   FileName = strFileName
End Property

'******************************************************************
' Class Constructor / Destructor
'******************************************************************
Private Sub Class_Initialize()
   'Nothing
End Sub

Private Sub Class_Terminate()
   If boolDocOpen Then
      Me.CloseHTMLDoc
   End If
End Sub
