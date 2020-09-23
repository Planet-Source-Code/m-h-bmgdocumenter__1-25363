Attribute VB_Name = "modReports"
Option Explicit

Public Sub ProduceProjectIndex(ProjectName As String, ProjectFile As String)
   Dim lngProjectFileHandle As Long
   Dim objHTMLIndex As CHTMLWrapper
   Dim strBuffer As String
   Dim strJunk As String
   Dim strFileName As String
   Dim colClasses As Collection
   Dim colForms As Collection
   Dim colModules As Collection
   Dim colDesigners As Collection
   Dim lngItemCtr As Long
   
   Set colClasses = New Collection
   Set colForms = New Collection
   Set colModules = New Collection
   Set colDesigners = New Collection
   
   'Open Project File
   lngProjectFileHandle = FreeFile
   Open ProjectFile For Input As #lngProjectFileHandle
   
   'Process Project File
   Do While Not EOF(lngProjectFileHandle)
      Line Input #lngProjectFileHandle, strBuffer
      
      If InStr(1, strBuffer, "Class=") > 0 Then
         colClasses.Add strBuffer
      ElseIf InStr(1, strBuffer, "IconForm=") > 0 Then
         'Do nothing (screws up Form finding!)
      ElseIf InStr(1, strBuffer, "Form=") > 0 Then
         colForms.Add strBuffer
      ElseIf InStr(1, strBuffer, "Module=") > 0 Then
         colModules.Add strBuffer
      ElseIf InStr(1, strBuffer, "Designer=") > 0 Then
         colDesigners.Add strBuffer
      Else
         'Nothing
      End If
   Loop
   
   'Close Project File
   Close lngProjectFileHandle
   
   'Open HTML Index
   Set objHTMLIndex = New CHTMLWrapper
   With objHTMLIndex
      'Start Document
      .OpenHTMLDoc modUtility.GetProjectPath(ProjectFile) & "DOCS\index.html"
      .OpenHTMLHeader ProjectName
      .CloseHTMLHeader
      
      'Write Visible Body
      .OpenHTMLBody
      .WriteHeader 1, "Project: " & ProjectName
      .WriteItalicText "Location: " & ProjectFile
      .InsertBreak
      .WriteItalicText "Produced: " & Format(Now, "mm/dd/yyyy")
      .InsertBreak
      .InsertHorizontalRule
      .WriteHeader 2, "Table of Contents"
   End With
   
   'Write out Classes
   objHTMLIndex.WriteBoldText "Classes (" & colClasses.Count & ")"
   objHTMLIndex.InsertBreak
   For lngItemCtr = 1 To colClasses.Count
      'Parse String
      strJunk = Left$(colClasses.Item(lngItemCtr), InStr(1, colClasses.Item(lngItemCtr), ";"))
      strFileName = Trim$(Mid$(colClasses.Item(lngItemCtr), Len(strJunk) + 1))
      
      'Product Class Documentation
      ProduceClassDocumentation ProjectName, modUtility.GetProjectPath(ProjectFile), strFileName
      
      'Write Index Entry
      objHTMLIndex.InsertLink "CLASSES\" & Replace$(Replace$(strFileName, ".", "_"), "\", "-") & ".html", strFileName
      objHTMLIndex.InsertBreak
   Next lngItemCtr
   
   'Write out Forms
   objHTMLIndex.WriteBoldText "Forms (" & colForms.Count & ")"
   objHTMLIndex.InsertBreak
   For lngItemCtr = 1 To colForms.Count
      'Parse String
      strFileName = Mid$(colForms.Item(lngItemCtr), Len("Forms="))
      
      'Product Form Documentation
      ProduceFormDocumentation ProjectName, modUtility.GetProjectPath(ProjectFile), strFileName
      
      'Write Index Entry
      objHTMLIndex.InsertLink "FORMS\" & Replace$(Replace$(strFileName, ".", "_"), "\", "-") & ".html", strFileName
      objHTMLIndex.InsertBreak
   Next lngItemCtr
   
   'Write out Modules
   objHTMLIndex.WriteBoldText "Modules (" & colModules.Count & ")"
   objHTMLIndex.InsertBreak
   For lngItemCtr = 1 To colModules.Count
      'Parse String
      strJunk = Left$(colModules.Item(lngItemCtr), InStr(1, colModules.Item(lngItemCtr), ";"))
      strFileName = Trim$(Mid$(colModules.Item(lngItemCtr), Len(strJunk) + 1))
      
      'Product Module Documentation
      ProduceModuleDocumentation ProjectName, modUtility.GetProjectPath(ProjectFile), strFileName
      
      'Write Index Entry
      objHTMLIndex.InsertLink "MODULES\" & Replace$(Replace$(strFileName, ".", "_"), "\", "-") & ".html", strFileName
      objHTMLIndex.InsertBreak
   Next lngItemCtr
   
   'Write out Designers
   objHTMLIndex.WriteBoldText "Data Environment/Reports (" & colDesigners.Count & ")"
   objHTMLIndex.InsertBreak
   For lngItemCtr = 1 To colDesigners.Count
      'Parse String
      strFileName = Mid$(colDesigners.Item(lngItemCtr), Len("Designer=") + 1)
      
      'Don't bother to parse Designers!
      
      'Write Index Entry
      objHTMLIndex.WriteNormalText strFileName
      objHTMLIndex.InsertBreak
   Next lngItemCtr
   
   'Close HTML Document
   With objHTMLIndex
      .InsertHorizontalRule
      .WriteItalicText "[End of File]"
      .CloseHTMLBody
      .CloseHTMLDoc
   End With
   
   'Clean up
   Set objHTMLIndex = Nothing
   Set colClasses = Nothing
   Set colForms = Nothing
   Set colModules = Nothing
   Set colDesigners = Nothing
End Sub

Private Sub ProduceClassDocumentation(ProjectName As String, ProjectLocation As String, FileName As String)
   Dim lngInputFileHandle As Long
   Dim objHTMLFile As CHTMLWrapper
   Dim strVBName As String
   Dim strBuffer As String
   Dim colSubs As Collection
   Dim colDeclares As Collection
   Dim colFunctions As Collection
   Dim colProperties As Collection
   Dim colGlobalVars As Collection
   Dim colTypeOrEnums As Collection
   Dim boolInProgramBody As Boolean
   Dim lngItemCtr As Long
   Dim lngLineCtr As Long
   Dim lngFunctionCtr As Long 'Count Subs Too...
   
   Set colGlobalVars = New Collection
   Set colSubs = New Collection
   Set colDeclares = New Collection
   Set colFunctions = New Collection
   Set colProperties = New Collection
   Set colTypeOrEnums = New Collection
   
   'Parse Class File...
   lngInputFileHandle = FreeFile
   Open ProjectLocation & FileName For Input As #lngInputFileHandle
   
   boolInProgramBody = False
   lngLineCtr = 0
   lngFunctionCtr = 0
   Do While Not EOF(lngInputFileHandle)
      strBuffer = GetCompleteLine(lngInputFileHandle)
      
      'Perform Line Count
      If Trim$(strBuffer) <> "" Then
         lngLineCtr = lngLineCtr + 1
      End If
      
      'Weed out comments
      If Left$(Trim$(strBuffer), 1) <> "'" Then
         If InStr(1, strBuffer, "Attribute VB_Name = ") > 0 Then
            strVBName = Mid$(strBuffer, Len("Attribute VB_Name = "))
            strVBName = Replace(strVBName, """", "")
         ElseIf Left$(strBuffer, 4) = "Sub " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colSubs.Add strBuffer
         ElseIf InStr(1, strBuffer, " Sub ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Sub")) = "Sub" _
            Or Left$(Trim$(strBuffer), Len("Public Sub")) = "Public Sub" _
            Or Left$(Trim$(strBuffer), Len("Private Sub")) = "Private Sub" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colSubs.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Declare ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Declare")) = "Declare" _
            Or Left$(Trim$(strBuffer), Len("Public Declare")) = "Public Declare" _
            Or Left$(Trim$(strBuffer), Len("Private Declare")) = "Private Declare" _
            Then
               colDeclares.Add strBuffer
            End If
         ElseIf Left$(strBuffer, 9) = "Function " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colFunctions.Add strBuffer
         ElseIf InStr(1, strBuffer, " Function ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Function")) = "Function" _
            Or Left$(Trim$(strBuffer), Len("Public Function")) = "Public Function" _
            Or Left$(Trim$(strBuffer), Len("Private Function")) = "Private Function" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colFunctions.Add strBuffer
            End If
         ElseIf Left$(strBuffer, 9) = "Property " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colProperties.Add strBuffer
         ElseIf InStr(1, strBuffer, " Property ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Property")) = "Property" _
            Or Left$(Trim$(strBuffer), Len("Public Property")) = "Public Property" _
            Or Left$(Trim$(strBuffer), Len("Private Property")) = "Private Property" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colProperties.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Type ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Type")) = "Type" _
            Or Left$(Trim$(strBuffer), Len("Public Type")) = "Public Type" _
            Or Left$(Trim$(strBuffer), Len("Private Type")) = "Private Type" _
            Then
               colTypeOrEnums.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Enum ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Enum")) = "Enum" _
            Or Left$(Trim$(strBuffer), Len("Public Enum")) = "Public Enum" _
            Or Left$(Trim$(strBuffer), Len("Private Enum")) = "Private Enum" _
            Then
               colTypeOrEnums.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Dim ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 3) = "Dim" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         ElseIf InStr(1, strBuffer, "Private ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 7) = "Private" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         ElseIf InStr(1, strBuffer, "Public ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 6) = "Public" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         Else
            'Nothing - unhandled right now...
         End If
      End If
   Loop
   
   Close #lngInputFileHandle
   
   'Write out documentation
   Set objHTMLFile = New CHTMLWrapper
   
   With objHTMLFile
      .OpenHTMLDoc ProjectLocation & "DOCS\CLASSES\" & Replace$(Replace$(FileName, ".", "_"), "\", "-") & ".html"
      
      .OpenHTMLHeader FileName
      .CloseHTMLHeader
      
      .OpenHTMLBody
      .WriteHeader 1, "Project: " & ProjectName
      .WriteHeader 2, "Class Name: " & strVBName
      .WriteItalicText "Location: " & FileName
      .InsertBreak
      .WriteItalicText "Produced: " & Format$(Now, "mm/dd/yyyy")
      .InsertBreak
      .InsertHorizontalRule
      
      'Write out Global Variables
      .WriteBoldText "Global Variables (" & colGlobalVars.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colGlobalVars.Count
         .WriteListItem colGlobalVars.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Types and Enums
      .WriteBoldText "Types and Enums (" & colTypeOrEnums.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colTypeOrEnums.Count
         .WriteListItem colTypeOrEnums.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Declarations
      .WriteBoldText "Declarations (" & colDeclares.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colDeclares.Count
         .WriteListItem colDeclares.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Subs
      .WriteBoldText "Subprocedures (" & colSubs.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colSubs.Count
         .WriteListItem colSubs.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Functions
      .WriteBoldText "Functions (" & colFunctions.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colFunctions.Count
         .WriteListItem colFunctions.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Properties
      .WriteBoldText "Properties (" & colProperties.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colProperties.Count
         .WriteListItem colProperties.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Summary
      .WriteBoldText "Summary"
      .InsertBreak
      .WriteNormalText "Lines of source: " & lngLineCtr & " (counting comments, but not blank lines)"
      .InsertBreak
      .WriteNormalText "Total function count: " & lngFunctionCtr & " (Subs, Functions, and Properties)"
      .InsertBreak
      
      If lngFunctionCtr > 0 Then
         .WriteNormalText "Average Lines/Function: " & Format$(lngLineCtr / lngFunctionCtr, "0.00")
      End If
      
      .InsertHorizontalRule
      .WriteItalicText "[End of File]"
      .CloseHTMLBody
      .CloseHTMLDoc
   End With
   
   Set objHTMLFile = Nothing
   
   Set colGlobalVars = Nothing
   Set colSubs = Nothing
   Set colDeclares = Nothing
   Set colFunctions = Nothing
   Set colProperties = Nothing
   Set colTypeOrEnums = Nothing
End Sub

Private Sub ProduceFormDocumentation(ProjectName As String, ProjectLocation As String, FileName As String)
   Dim lngInputFileHandle As Long
   Dim strVBName As String
   Dim strBuffer As String
   Dim objHTMLFile As CHTMLWrapper
   Dim colElements As Collection
   Dim colElementNames As Collection
   Dim colSubs As Collection
   Dim colDeclares As Collection
   Dim colFunctions As Collection
   Dim colProperties As Collection
   Dim colGlobalVars As Collection
   Dim colTypeOrEnums As Collection
   Dim boolInProgramBody As Boolean
   Dim lngItemCtr As Long
   Dim lngLineCtr As Long
   Dim lngFunctionCtr As Long

   Set colGlobalVars = New Collection
   Set colElements = New Collection
   Set colElementNames = New Collection
   Set colSubs = New Collection
   Set colDeclares = New Collection
   Set colFunctions = New Collection
   Set colProperties = New Collection
   Set colTypeOrEnums = New Collection
   
   'Parse Form File...
   lngInputFileHandle = FreeFile
   Open ProjectLocation & FileName For Input As #lngInputFileHandle
   
   boolInProgramBody = False
   lngLineCtr = 0
   lngFunctionCtr = 0
   Do While Not EOF(lngInputFileHandle)
      strBuffer = GetCompleteLine(lngInputFileHandle)
      
      'Perform Line Count
      If Trim$(strBuffer) <> "" Then
         lngLineCtr = lngLineCtr + 1
      End If
      
      'Weed out comments
      If Left$(Trim$(strBuffer), 1) <> "'" Then
         If InStr(1, strBuffer, "Attribute VB_Name = ") > 0 Then
            strVBName = Mid$(strBuffer, Len("Attribute VB_Name = "))
            strVBName = Replace(strVBName, """", "")
         ElseIf InStr(1, strBuffer, "Begin VB.") > 0 Then
            'Handle form elements...
            strBuffer = Trim$(strBuffer)                      'Strip spaces
            strBuffer = Mid$(strBuffer, Len("Begin VB.") + 1) 'Strip Begin.VB.
            
            'These two collections should always be synchronized!
            colElements.Add Trim$(Left$(strBuffer, InStr(1, strBuffer, " ")))
            colElementNames.Add Mid$(strBuffer, InStr(1, strBuffer, " "))
         ElseIf Left$(strBuffer, 4) = "Sub " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colSubs.Add strBuffer
         ElseIf InStr(1, strBuffer, " Sub ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Sub")) = "Sub" _
            Or Left$(Trim$(strBuffer), Len("Public Sub")) = "Public Sub" _
            Or Left$(Trim$(strBuffer), Len("Private Sub")) = "Private Sub" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colSubs.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Declare ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Declare")) = "Declare" _
            Or Left$(Trim$(strBuffer), Len("Public Declare")) = "Public Declare" _
            Or Left$(Trim$(strBuffer), Len("Private Declare")) = "Private Declare" _
            Then
               colDeclares.Add strBuffer
            End If
         ElseIf Left$(strBuffer, 9) = "Function " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colFunctions.Add strBuffer
         ElseIf InStr(1, strBuffer, " Function ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Function")) = "Function" _
            Or Left$(Trim$(strBuffer), Len("Public Function")) = "Public Function" _
            Or Left$(Trim$(strBuffer), Len("Private Function")) = "Private Function" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colFunctions.Add strBuffer
            End If
         ElseIf Left$(strBuffer, 9) = "Property " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colProperties.Add strBuffer
         ElseIf InStr(1, strBuffer, " Property ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Property")) = "Property" _
            Or Left$(Trim$(strBuffer), Len("Public Property")) = "Public Property" _
            Or Left$(Trim$(strBuffer), Len("Private Property")) = "Private Property" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colProperties.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Type ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Type")) = "Type" _
            Or Left$(Trim$(strBuffer), Len("Public Type")) = "Public Type" _
            Or Left$(Trim$(strBuffer), Len("Private Type")) = "Private Type" _
            Then
               colTypeOrEnums.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Enum ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Enum")) = "Enum" _
            Or Left$(Trim$(strBuffer), Len("Public Enum")) = "Public Enum" _
            Or Left$(Trim$(strBuffer), Len("Private Enum")) = "Private Enum" _
            Then
               colTypeOrEnums.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Dim ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 3) = "Dim" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         ElseIf InStr(1, strBuffer, "Private ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 7) = "Private" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         ElseIf InStr(1, strBuffer, "Public ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 6) = "Public" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         Else
            'Nothing - unhandled right now...
         End If
      End If
   Loop
   
   Close #lngInputFileHandle
   
   'Write out documentation
   Set objHTMLFile = New CHTMLWrapper
   
   With objHTMLFile
      .OpenHTMLDoc ProjectLocation & "DOCS\FORMS\" & Replace$(Replace$(FileName, ".", "_"), "\", "-") & ".html"
      
      .OpenHTMLHeader FileName
      .CloseHTMLHeader
      
      .OpenHTMLBody
      .WriteHeader 1, "Project: " & ProjectName
      .WriteHeader 2, "Form Name: " & strVBName
      .WriteItalicText "Location: " & FileName
      .InsertBreak
      .WriteItalicText "Produced: " & Format$(Now, "mm/dd/yyyy")
      .InsertBreak
      .InsertHorizontalRule
      
      'Write out Global Variables
      .WriteBoldText "Global Variables (" & colGlobalVars.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colGlobalVars.Count
         .WriteListItem colGlobalVars.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Types and Enums
      .WriteBoldText "Types and Enums (" & colTypeOrEnums.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colTypeOrEnums.Count
         .WriteListItem colTypeOrEnums.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Declarations
      .WriteBoldText "Declarations (" & colDeclares.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colDeclares.Count
         .WriteListItem colDeclares.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Form Elements
      .WriteBoldText "Controls (" & colElements.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colElements.Count
         .WriteListItem colElementNames.Item(lngItemCtr) & _
                        " {<I>" & colElements.Item(lngItemCtr) & "</I>}"
      Next lngItemCtr
      .EndList False
      
      'Write out Subs
      .WriteBoldText "Subprocedures (" & colSubs.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colSubs.Count
         .WriteListItem colSubs.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Functions
      .WriteBoldText "Functions (" & colFunctions.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colFunctions.Count
         .WriteListItem colFunctions.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Properties
      .WriteBoldText "Properties (" & colProperties.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colProperties.Count
         .WriteListItem colProperties.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
           
      'Write out Summary
      .WriteBoldText "Summary"
      .InsertBreak
      .WriteNormalText "Lines of source: " & lngLineCtr & " (counting comments, but not blank lines)"
      .InsertBreak
      .WriteNormalText "Total function count: " & lngFunctionCtr & " (Subs, Functions, and Properties)"
      .InsertBreak
      
      If lngFunctionCtr > 0 Then
         .WriteNormalText "Average Lines/Function: " & Format$(lngLineCtr / lngFunctionCtr, "0.00")
      End If
      
      .InsertHorizontalRule
      .WriteItalicText "[End of File]"
      .CloseHTMLBody
      .CloseHTMLDoc
   End With
   
   Set objHTMLFile = Nothing

   Set colGlobalVars = Nothing
   Set colElements = Nothing
   Set colElementNames = Nothing
   Set colSubs = Nothing
   Set colDeclares = Nothing
   Set colFunctions = Nothing
   Set colProperties = Nothing
   Set colTypeOrEnums = Nothing
End Sub

Private Sub ProduceModuleDocumentation(ProjectName As String, ProjectLocation As String, FileName As String)
   Dim lngInputFileHandle As Long
   Dim strVBName As String
   Dim strBuffer As String
   Dim objHTMLFile As CHTMLWrapper
   Dim colSubs As Collection
   Dim colDeclares As Collection
   Dim colFunctions As Collection
   Dim colProperties As Collection
   Dim colGlobalVars As Collection
   Dim colTypeOrEnums As Collection
   Dim boolInProgramBody As Boolean
   Dim lngItemCtr As Long
   Dim lngLineCtr As Long
   Dim lngFunctionCtr As Long

   Set colGlobalVars = New Collection
   Set colSubs = New Collection
   Set colDeclares = New Collection
   Set colFunctions = New Collection
   Set colProperties = New Collection
   Set colTypeOrEnums = New Collection

   'Parse Module File...
   lngInputFileHandle = FreeFile
   Open ProjectLocation & FileName For Input As #lngInputFileHandle
   
   boolInProgramBody = False
   lngLineCtr = 0
   lngFunctionCtr = 0
   Do While Not EOF(lngInputFileHandle)
      strBuffer = GetCompleteLine(lngInputFileHandle)
      
      'Perform Line Count
      If Trim$(strBuffer) <> "" Then
         lngLineCtr = lngLineCtr + 1
      End If
      
      'Weed out comments
      If Left$(Trim$(strBuffer), 1) <> "'" Then
         If InStr(1, strBuffer, "Attribute VB_Name = ") > 0 Then
            strVBName = Mid$(strBuffer, Len("Attribute VB_Name = "))
            strVBName = Replace(strVBName, """", "")
         ElseIf Left$(strBuffer, 4) = "Sub " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colSubs.Add strBuffer
         ElseIf InStr(1, strBuffer, " Sub ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Sub")) = "Sub" _
            Or Left$(Trim$(strBuffer), Len("Public Sub")) = "Public Sub" _
            Or Left$(Trim$(strBuffer), Len("Private Sub")) = "Private Sub" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colSubs.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Declare ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Declare")) = "Declare" _
            Or Left$(Trim$(strBuffer), Len("Public Declare")) = "Public Declare" _
            Or Left$(Trim$(strBuffer), Len("Private Declare")) = "Private Declare" _
            Then
               colDeclares.Add strBuffer
            End If
         ElseIf Left$(strBuffer, 9) = "Function " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colFunctions.Add strBuffer
         ElseIf InStr(1, strBuffer, " Function ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Function")) = "Function" _
            Or Left$(Trim$(strBuffer), Len("Public Function")) = "Public Function" _
            Or Left$(Trim$(strBuffer), Len("Private Function")) = "Private Function" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colFunctions.Add strBuffer
            End If
         ElseIf Left$(strBuffer, 9) = "Property " Then
            boolInProgramBody = True
            lngFunctionCtr = lngFunctionCtr + 1
            colProperties.Add strBuffer
         ElseIf InStr(1, strBuffer, " Property ") > 0 Then
            boolInProgramBody = True
            If Left$(Trim$(strBuffer), Len("Property")) = "Property" _
            Or Left$(Trim$(strBuffer), Len("Public Property")) = "Public Property" _
            Or Left$(Trim$(strBuffer), Len("Private Property")) = "Private Property" _
            Then
               lngFunctionCtr = lngFunctionCtr + 1
               colProperties.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Type ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Type")) = "Type" _
            Or Left$(Trim$(strBuffer), Len("Public Type")) = "Public Type" _
            Or Left$(Trim$(strBuffer), Len("Private Type")) = "Private Type" _
            Then
               colTypeOrEnums.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Enum ") > 0 Then
            If Left$(Trim$(strBuffer), Len("Enum")) = "Enum" _
            Or Left$(Trim$(strBuffer), Len("Public Enum")) = "Public Enum" _
            Or Left$(Trim$(strBuffer), Len("Private Enum")) = "Private Enum" _
            Then
               colTypeOrEnums.Add strBuffer
            End If
         ElseIf InStr(1, strBuffer, "Dim ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 3) = "Dim" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         ElseIf InStr(1, strBuffer, "Private ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 7) = "Private" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         ElseIf InStr(1, strBuffer, "Public ") > 0 Then 'A Global Var
            If Not boolInProgramBody Then
               If Left$(Trim$(strBuffer), 6) = "Public" Then
                  colGlobalVars.Add strBuffer
               End If
            End If
         Else
            'Nothing - unhandled right now...
         End If
      End If
   Loop
   
   Close #lngInputFileHandle
   
   'Write out documentation
   Set objHTMLFile = New CHTMLWrapper
   
   With objHTMLFile
      .OpenHTMLDoc ProjectLocation & "DOCS\MODULES\" & Replace$(Replace$(FileName, ".", "_"), "\", "-") & ".html"
      
      .OpenHTMLHeader FileName
      .CloseHTMLHeader
      
      .OpenHTMLBody
      .WriteHeader 1, "Project: " & ProjectName
      .WriteHeader 2, "Module Name: " & strVBName
      .WriteItalicText "Location: " & FileName
      .InsertBreak
      .WriteItalicText "Produced: " & Format$(Now, "mm/dd/yyyy")
      .InsertBreak
      .InsertHorizontalRule
      
      'Write out Global Variables
      .WriteBoldText "Global Variables (" & colGlobalVars.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colGlobalVars.Count
         .WriteListItem colGlobalVars.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Types and Enums
      .WriteBoldText "Types and Enums (" & colTypeOrEnums.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colTypeOrEnums.Count
         .WriteListItem colTypeOrEnums.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Declarations
      .WriteBoldText "Declarations (" & colDeclares.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colDeclares.Count
         .WriteListItem colDeclares.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Subs
      .WriteBoldText "Subprocedures (" & colSubs.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colSubs.Count
         .WriteListItem colSubs.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Functions
      .WriteBoldText "Functions (" & colFunctions.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colFunctions.Count
         .WriteListItem colFunctions.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Properties
      .WriteBoldText "Properties (" & colProperties.Count & ")"
      .StartList False
      For lngItemCtr = 1 To colProperties.Count
         .WriteListItem colProperties.Item(lngItemCtr)
      Next lngItemCtr
      .EndList False
      
      'Write out Summary
      .WriteBoldText "Summary"
      .InsertBreak
      .WriteNormalText "Lines of source: " & lngLineCtr & " (counting comments, but not blank lines)"
      .InsertBreak
      .WriteNormalText "Total function count: " & lngFunctionCtr & " (Subs, Functions, and Properties)"
      .InsertBreak
      
      If lngFunctionCtr > 0 Then
         .WriteNormalText "Average Lines/Function: " & Format$(lngLineCtr / lngFunctionCtr, "0.00")
      End If
      
      .InsertHorizontalRule
      .WriteItalicText "[End of File]"
      .CloseHTMLBody
      .CloseHTMLDoc
   End With
   
   Set objHTMLFile = Nothing

   Set colGlobalVars = Nothing
   Set colSubs = Nothing
   Set colDeclares = New Collection
   Set colFunctions = Nothing
   Set colProperties = Nothing
   Set colTypeOrEnums = Nothing
End Sub

