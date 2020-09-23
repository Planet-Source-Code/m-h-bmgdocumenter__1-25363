VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMG Source Code Documenter"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDocument 
      Caption         =   "&Document!"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtProjectTitle 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   825
      Width           =   3015
   End
   Begin VB.CommandButton cmdBrowseForProject 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtProjectLocation 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   225
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Project Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "VB Project:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseForProject_Click()
   frmProjectBrowser.Show vbModal, Me
End Sub

Private Sub cmdDocument_Click()
   Dim strProjectPath As String
   Dim intResponse As Integer
   
   If Not FormValidated Then
      MsgBox "You must specify valid values!", , "Error"
      Exit Sub
   End If
   
   strProjectPath = modUtility.GetProjectPath(txtProjectLocation.Text)

   If strProjectPath <> "" Then
      'Make Directories for output
      On Error Resume Next 'overwrite by default...
      MkDir strProjectPath & "DOCS"
      MkDir strProjectPath & "DOCS\CLASSES"
      MkDir strProjectPath & "DOCS\FORMS"
      MkDir strProjectPath & "DOCS\MODULES"
      On Error GoTo 0
      
      'Load and Parse Project File (which will produce lower level documentation)
      modReports.ProduceProjectIndex txtProjectTitle.Text, txtProjectLocation.Text
      
      'Finish
      intResponse = MsgBox("Would you like to see the documentation now?", vbYesNo, "Confirm")
      If intResponse = vbYes Then
         Shell "start " & """" & strProjectPath & "DOCS\index.html" & """", vbHide
      End If
   Else
      MsgBox "Your project is missing or damaged." & vbNewLine & _
         "Please try locating it again.", , "Error"
   End If
End Sub

Private Function FormValidated() As Boolean
   FormValidated = False
   
   If Trim$(txtProjectLocation.Text) = "" Then Exit Function
   If Not modUtility.FileExists(txtProjectLocation.Text) Then Exit Function
   If Trim$(txtProjectTitle.Text) = "" Then Exit Function
   
   FormValidated = True
End Function
