VERSION 5.00
Begin VB.Form frmAddFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add File"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   5295
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.DirListBox Dir 
      Appearance      =   0  'Flat
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.FileListBox File 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
FileToAdd = txtFileName.Text

Remember 'set the frmAddFile back to is last opened location
frmWhere.Show
End Sub

Private Sub Dir_Change()
File.FileName = Dir.Path

End Sub

Private Sub Drive1_Change()
On Error GoTo Errr:
Dir.Path = Drive1.Drive
Errr:
Select Case Err
Case 68
MsgBox "The drive doesn't appear to be available."
Drive1.Drive = App.Path
Drive1.SetFocus
End Select
End Sub

Private Sub File_Click()
txtFileName.Text = File.FileName
txt1.Text = Dir.Path & "\" & File.FileName
End Sub

Private Sub Form_Load()
If RemDrive <> "" Then Me.Drive1.Drive = RemDrive
If RemDir <> "" Then Me.Dir.Path = RemDir

End Sub
