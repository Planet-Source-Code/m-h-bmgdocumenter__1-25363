VERSION 5.00
Begin VB.Form frmProjectBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse For VB Project..."
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.FileListBox filProjectFile 
      Height          =   2820
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox dirDirToProject 
      Height          =   1665
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.DriveListBox drvDriveToProject 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "frmProjectBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   If filProjectFile.FileName = "" Then
      MsgBox "You must select a filename!", , "Error"
      Exit Sub
   Else
      frmMain.txtProjectLocation.Text = GetFullFileName
   End If
   
   Unload Me
End Sub

Private Sub dirDirToProject_Change()
   filProjectFile.Path = dirDirToProject.Path
   filProjectFile.Refresh
End Sub

Private Sub drvDriveToProject_Change()
   dirDirToProject.Path = Left$(drvDriveToProject.Drive, 2) & "\"
   dirDirToProject.Refresh
End Sub

Private Sub filProjectFile_DblClick()
   frmMain.txtProjectLocation.Text = GetFullFileName
   
   Unload Me
End Sub

Private Sub Form_Load()
   filProjectFile.Pattern = "*.vbp" 'VB Project files
End Sub

Private Function GetFullFileName() As String
   Dim strFullPath As String
   
   strFullPath = dirDirToProject.Path
   If Right$(strFullPath, 1) <> "\" Then
      strFullPath = strFullPath & "\"
   End If
   
   GetFullFileName = strFullPath & filProjectFile.FileName
   
   GetFullFileName = UCase$(GetFullFileName)
End Function
