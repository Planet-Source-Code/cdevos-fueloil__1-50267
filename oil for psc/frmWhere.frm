VERSION 5.00
Begin VB.Form frmWhere 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Install Where"
   ClientHeight    =   1110
   ClientLeft      =   11985
   ClientTop       =   1065
   ClientWidth     =   2070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSystem 
      Caption         =   "System"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAppPath 
      Caption         =   "AppPath"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Folder"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheFolder As String
Option Explicit

Private Sub cmdAppPath_Click()
    Dim itmX As ListItem
'add to the the listview
     Set itmX = frmMain.listFiles.ListItems.Add()
        itmX.Text = FileToAdd
        itmX.SubItems(1) = frmAddFile.txt1.Text
        itmX.SubItems(2) = "/AppPath"
        itmX.Checked = True
'FileToAdd = FileToAdd & "/AppPath"
 Me.Hide
Unload frmAddFile
frmMain.cmdNext.Enabled = True 'enable the next button
End Sub


Private Sub cmdNew_Click(Index As Integer)
TheFolder = cmdNew(Index).Caption

'add to the the listview
Dim itmX As ListItem
     Set itmX = frmMain.listFiles.ListItems.Add()
        itmX.Text = FileToAdd
        itmX.SubItems(1) = frmAddFile.txt1.Text
        itmX.SubItems(2) = "/" & TheFolder
        itmX.Checked = True
'FileToAdd = FileToAdd & "/" & TheFolder
Me.Hide
Unload frmAddFile
frmMain.cmdNext.Enabled = True 'enable the next button

End Sub

Private Sub cmdSystem_Click()
 'add to the the listview
    Dim itmX As ListItem
     Set itmX = frmMain.listFiles.ListItems.Add()
        itmX.Text = FileToAdd
        itmX.SubItems(1) = frmAddFile.txt1.Text
        itmX.SubItems(2) = "/System"
        itmX.Checked = True
'take the FileToAdd from the frmWhere and add "/System"
'so that the Fuel installer knows where to install the file
'FileToAdd = FileToAdd & "/System"
 Me.Hide
Unload frmAddFile
frmMain.cmdNext.Enabled = True 'enable the next button

End Sub


