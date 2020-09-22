VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oil"
   ClientHeight    =   5505
   ClientLeft      =   2895
   ClientTop       =   2700
   ClientWidth     =   8430
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8430
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   6840
      Width           =   5775
   End
   Begin RichTextLib.RichTextBox txtCombo 
      Height          =   735
      Left            =   2160
      TabIndex        =   50
      Top             =   6000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":030A
   End
   Begin VB.CheckBox ckView 
      Caption         =   "View Oil List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   46
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next  >"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   0
      Picture         =   "frmMain.frx":038C
      ScaleHeight     =   5505
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   240
         MaskColor       =   &H8000000F&
         TabIndex        =   29
         Top             =   4920
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtIcon 
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7800
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cmImage 
      Left            =   6000
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   -2147483633
      Filter          =   ".bmp"
   End
   Begin MSComDlg.CommonDialog cmOpen 
      Left            =   8280
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   3
      Left            =   1920
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   6100
      Begin VB.CommandButton cmdNewFolder 
         Caption         =   "Create New Folder"
         Height          =   495
         Left            =   4560
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtFolderName 
         Height          =   285
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Create any new Folders in the Applicaton's Directory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Folder Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Click ""Next""  if you need no new Folders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   20
         Top             =   3000
         Width           =   3510
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4365
      Index           =   4
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   6100
      Begin VB.CommandButton cmdFiles 
         Caption         =   "Add File"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   3840
         Width           =   1215
      End
      Begin MSComctlLib.ListView listFiles 
         Height          =   2895
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "source"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Install to"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Add any New files such as OCX""s or DLL's that were not part of the original Applications package. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "Select the file and Uncheck  to remove it from the list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   33
         Top             =   3720
         Width           =   4065
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   7
      Left            =   2040
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdView 
         Caption         =   "View the Files"
         Height          =   375
         Left            =   3120
         TabIndex        =   48
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblWhere 
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label12 
         Caption         =   "Oil Package has been succesfully created at:"
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
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   4935
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   6
      Left            =   1920
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox ckZip 
         Caption         =   "Compress ""Files"" folder       in Zip format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         TabIndex        =   42
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtSave 
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   480
         Width           =   5775
      End
      Begin VB.CommandButton cmdFolder 
         Caption         =   "New folder"
         Height          =   375
         Left            =   4560
         TabIndex        =   36
         Top             =   3840
         Width           =   1455
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   2775
      End
      Begin VB.DirListBox Dir 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label13 
         Caption         =   "Select a Directory where the Oil package will be saved"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label lblSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3240
         TabIndex        =   38
         Top             =   1200
         Width           =   2655
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   6100
      Begin VB.OptionButton optWhere 
         Caption         =   "Install from a File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   44
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optWhere 
         Caption         =   "Install from the Web"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   43
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtUrl 
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblInstallFromHelp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   360
         TabIndex        =   45
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label lblUrl 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Full Url of the files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select where the Update files are to be installed from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   120
         Width           =   4665
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   5
      Left            =   1920
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   6100
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3360
         ScaleHeight     =   585
         ScaleWidth      =   705
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "Add Image"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Click ""Next""  if you do not wish to place the Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   28
         Top             =   3720
         Width           =   4320
      End
      Begin VB.Label Label7 
         Caption         =   "The Image must be a GIF image and cannot be larger then 55 x 55 pixels."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "Choose an Image to be placed on the Fuel Installer that represents your Application."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4125
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   6100
      Begin VB.CheckBox ckNoLog 
         Caption         =   $"frmMain.frx":20680
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   49
         Top             =   2640
         Width           =   5295
      End
      Begin VB.ComboBox cmbApp 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   5415
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Select an EXE for Oil Package"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Make sure that you have compliled the most current version of your Application. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         TabIndex        =   32
         Top             =   1920
         Width           =   5295
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1800
      X2              =   8160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0059C0EE&
      BorderWidth     =   2
      X1              =   8160
      X2              =   1800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblVersion 
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblAppName 
      Height          =   135
      Left            =   3600
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Dim ThePath As String
Dim FrameCount As Integer 'track which frame is open

Private Sub ckNoLog_Click()
Select Case ckNoLog.Value
Case 1 'update the install log
NoLog = False
Case 0 'don't update the install log
NoLog = True
End Select
End Sub

Private Sub ckView_Click()

Select Case ckView.Value
Case 1
frmOilIni.Visible = True
Case 0
frmOilIni.Visible = False
End Select
End Sub

Private Sub ckZip_Click()
'make a zip or not
Select Case ckZip.Value
Case 1
DoCompress = True
Case 0
DoCompress = False
End Select
End Sub



Private Sub cmbApp_Click()
On Error Resume Next
 addList = frmMain.txtCombo.Find(frmMain.cmbApp.Text, 0)
 txtCombo.SetFocus
 Text1.Text = addList
lblVersion.Caption = GetProductVersion(cmbApp.Text)
lblAppName.Caption = GetProductName(cmbApp.Text)
AppName = GetOriginalFilename(cmbApp.Text)
AppVersion = GetProductVersion(cmbApp.Text)
If AppName = "" Then
Exit Sub
Else
cmdNext.Enabled = True
End If
End Sub

Private Sub cmdBrowse_Click()
Dim tempLog As String
On Error GoTo done:
cmOpen.ShowOpen
cmbApp.Text = cmOpen.FileName
'get the info from the registry about the EXE we are going to package
lblVersion.Caption = GetProductVersion(cmbApp.Text)
lblAppName.Caption = GetProductName(cmbApp.Text)
AppName = GetOriginalFilename(cmbApp.Text)
AppVersion = GetProductVersion(cmbApp.Text)
cmdNext.Enabled = True
done:
 addList = frmMain.txtCombo.Find(frmMain.cmbApp.Text, 0)
 txtCombo.SetFocus
 Text1.Text = addList

End Sub

Private Sub cmdCancel_Click()
Dim leave
leave = MsgBox("Are you sure you want to quite?", vbOKCancel, "Quit")
If leave = vbOK Then End

End Sub

Private Sub cmdFiles_Click()
frmAddFile.Show
'enable the next button if a file as been added
If listFiles.ListItems.Count > 0 Then
cmdNext.Enabled = True
Else
cmdNext.Enabled = False
End If

End Sub

Private Sub cmdFolder_Click()
On Error GoTo done:
Dim MakeNew As String
MakeNew = InputBox("The New Folder will be created in the Directory:" & vbNewLine & Dir.Path & "Enter new name", "NewFolder")
MkDir Dir.Path & "\" & MakeNew
Dir.Refresh
done:
End Sub



Private Sub cmdHelp_Click()
'open the help files
 ShellExecute hwnd, "open", App.Path & "\help\index.htm", vbNullString, vbNullString, conSwNormal

End Sub

Private Sub cmdIcon_Click()
On Error Resume Next
Dim intFile As Integer
 cmImage.ShowOpen
 intFile = FreeFile
 Open cmImage.FileName For Input As intFile
 txtIcon.Text = cmImage.FileName
 Pic.Picture = LoadPicture(txtIcon.Text)

 
 Close #intFile
   
 InstallerIcon = txtIcon.Text
Pic.Refresh

End Sub

Private Sub cmdNewFolder_Click()
Dim CName As String 'label control name
'frmWhere.Show
If txtFolderName = vbNullString Then Exit Sub
FileToAdd = txtFolderName & "/NewFolder"
WriteTheFile
'add a new button to frmWhere to show the new folder name
 Load frmWhere.cmdNew(NoNewFolders)
With frmWhere.cmdNew(NoNewFolders)
'set the properties for the new button
.Caption = txtFolderName
.Visible = True
.Height = 375
If NoNewFolders = 1 Then
.Top = 1095
Else
.Top = (NoNewFolders * 495) + 600
End If
.Left = 240
.Width = 1695
.FontSize = 10
End With
frmWhere.Height = frmWhere.Height + 495
NoNewFolders = NoNewFolders + 1
txtFolderName = ""
End Sub

Private Sub cmdNext_Click()
Select Case FrameCount
Case 1
If AppName = "" Or AppVersion = "" Then
MsgBox "Oil was unable to obtain the correct information "
Exit Sub
End If
'add the app name to the oil list
frmOilIni.rtfOil.SelText = "[AppName]" & vbNewLine & "Name=" & Replace(AppName, ".exe", vbNullString) & vbNewLine

'add the version number to the oil list
frmOilIni.rtfOil.SelText = "[Update]" & vbNewLine & "Update=" & AppVersion & vbNewLine
'add the EXE to file list
 Set itmX = frmMain.listFiles.ListItems.Add()
        itmX.Text = AppName '& ".exe"
        itmX.SubItems(1) = cmOpen.FileName
        itmX.SubItems(2) = "/AppPath"
        itmX.Checked = True
       
cmdNext.Enabled = True
frmMain.txtCombo.Find (cmbApp.Text)

Case 2
If CanCompress = False Then
If txtUrl.Text = "" Then 'require a url
MsgBox "Please enter a URL"
txtUrl.SetFocus
Exit Sub
End If
End If
InstallFrom = txtUrl.Text
WriteInstallType

Case 4
'disable the next button until a file is added
If listFiles.ListItems.Count > 0 Then
cmdNext.Enabled = True
Else
cmdNext.Enabled = False
End If

Dim FilesToWrite As Integer
'loop through the list items and write each one to the oil list
For FilesToWrite = 1 To listFiles.ListItems.Count
With listFiles.ListItems.Item(FilesToWrite)
FileToAdd = .Text & .ListSubItems(2) 'used for writing to the oil list
FilePath = .ListSubItems(1) 'used for FileNameSource array
tempFileName = .Text 'used for FileName array
End With
WriteTheFile
Next FilesToWrite
'write the total files added to the oil list
frmOilIni.rtfOil.SelText = "[NoFiles]" & vbNewLine & "No=" & NumberOfFiles
UpdateLog

Case 6
If DoCompress = True Then
CopyFilesAndZip 'call CopyFilesAndZip sub in Module1 if want to Zip
Else
CopyFiles 'call the CopyFiles sub in Module1 if don't want to zip
End If
SavePackedList

Case 7
End
End Select


'advance the FrameCount to show the next frame
FrameCount = FrameCount + 1
If FrameCount >= 8 Then Exit Sub
'move through the frames
If FrameCount < 1 Then
Frame(1).Visible = False
Else
Frame(FrameCount - 1).Visible = False
End If
Frame(FrameCount).Visible = True
'disable the next button until a file is added
If FrameCount = 4 Then
If listFiles.ListItems.Count > 0 Then
cmdNext.Enabled = True
Else
cmdNext.Enabled = False
End If
End If
If FrameCount = 2 Then cmdNext.Enabled = False
If FrameCount = 6 Then
If CanCompress = True Then  'disable ckZip if install from web was choosen
ckZip.Enabled = True
Else
ckZip.Enabled = False
End If
lblSave.Caption = "The package will be saved in a folder named " & AppName & "OIL in the Directory you choose"
End If
If FrameCount = 7 Then
cmdNext.Caption = "Finished"
lblWhere.Caption = frmMain.Dir.Path & "\" & AppName & "OIL"
End If
End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdView_Click()
Dim openit As String
openit = lblWhere.Caption
Shell "explorer " & openit & """"
End Sub

Private Sub Dir_Change()
txtSave.Text = Dir.Path
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

Private Sub Form_Load()
NoNewFolders = 1
FrameCount = 1
Me.cmImage.Filter = "Image (*.gif)|*.gif"
cmOpen.Filter = "Executable (*exe) |*.exe"
NumberOfFiles = 0
txtSave.Text = Dir.Path
LoadPackedList App.Path & "\PackedList.txt", cmbApp
txtCombo.LoadFile App.Path & "\PackedList.txt"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmOilIni
End Sub


Private Sub listFiles_Click()
      If listFiles.ListItems.Count = 0 Then
      cmdNext.Enabled = False
      Exit Sub
      End If
      If listFiles.SelectedItem.Checked = False Then
      listFiles.ListItems.Remove listFiles.SelectedItem.Index
      If listFiles.ListItems.Count = 0 Then
      cmdNext.Enabled = False 'disable cmdnext if no files
      End If
      End If

End Sub

Private Sub optWhere_Click(Index As Integer)
'choose where the Fuel Installer will install the files from
Select Case optWhere(Index).Index
Case 0 'WEB
lblUrl.Visible = True
txtUrl.Visible = True
InstallType = "Web"
lblInstallFromHelp.Caption = "Select ""Install from the Web"" if you intend to place the files packaged by OIL on a web server and have the Fuel Installer download them."
lblInstallFromHelp.Caption = lblInstallFromHelp.Caption & vbNewLine & "When Selecting this option the files CANNOT be complied into a Zip format."
CanCompress = False
Case 1 'File
InstallType = "File"
cmdNext.Enabled = True
lblUrl.Visible = False
txtUrl.Visible = False
lblInstallFromHelp.Caption = "Select ""Install from a File"" if you intend to distrbute the files packaged by OIL on a disk or have them downloaded independently from the Fuel Installer."
lblInstallFromHelp.Caption = lblInstallFromHelp.Caption & vbNewLine & "When Selecting this option the files CAN be complied into a Zip format."
CanCompress = True
End Select
End Sub

Private Sub txtSave_Change()
On Error Resume Next
 Dir.Path = txtSave.Text

End Sub

Private Sub txtUrl_Change()
'check to see if a url was added before enabling cmdNext
If txtUrl.Text = "" Then
cmdNext.Enabled = flase
Else
cmdNext.Enabled = True
End If
End Sub

Public Sub CopyFilesAndZip()
On Error Resume Next
Dim c As Integer
Dim FileSelected



MkDir frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" 'make the Oil folder
ThePath = frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" & "\" & "Files.zip" 'set the path to the new oil folder for copying

'loop through the NumberOfFiles and copy the the files to the new oil folder
For c = 0 To NumberOfFiles
If FileSelected = "" Then
            FileSelected = FileNameSource(c)
        Else
            FileSelected = FileSelected & "*" & FileNameSource(c)
        End If
    'End If
Next c
'zip it up
Print CompressFiles(FileSelected, ThePath)
  'save the text file
Open frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" & "\install.oil" For Output As #1
    Print #1, frmOilIni.rtfOil.Text
    Close #1
'copy the update installer
CopyFile App.Path & "\Fuel.exe", Dir.Path & "\" & AppName & "OIL" & "\Fuel.exe", 1
'copy the image if it was selected
If InstallerIcon <> "" Then CopyFile InstallerIcon, frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL\icon.gif", 1

End Sub

Private Function CompressFiles(theFiles, outputfile)
'*****************************************************
'thanks to OldManMarcin mpasek@polaccess.com
' for this java zip class
'*****************************************************
 Set javaObject = GetObject("java:ZipFunctions")
 strResult = javaObject.ZipFile(theFiles, outputfile)
 Set javaObject = Nothing
 CompressFiles = strResult
End Function


