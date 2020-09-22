VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   8385
   ClientTop       =   1005
   ClientWidth     =   4890
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4890
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   7560
      Width           =   6495
   End
   Begin VB.TextBox txtFileInstalling 
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   7200
      Width           =   5775
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3840
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2400
      Width           =   4515
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   735
      ScaleWidth      =   4890
      TabIndex        =   6
      Top             =   0
      Width           =   4890
   End
   Begin MSComDlg.CommonDialog cmFindLog 
      Left            =   5880
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblV 
      Caption         =   "lblNoFiles"
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   6720
      Width           =   1755
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7320
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblFiles 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblFiles"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   3675
   End
   Begin VB.Label lblNoFiles 
      Caption         =   "lblNoFiles"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   6720
      Width           =   795
   End
   Begin VB.Label LBLwin 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   5640
   End
   Begin VB.Label lblFileInstalling 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim nofiles As Integer      'number of files to install including new folders
Dim UpdateNo As String      'Update Version
Dim CurrentV As String      'current version of the app being updated
Dim InstallFrom As String   'where the file for the update are located online or on the computer
Dim UpdateLocation As String 'url of the files if they are online
Dim LenAppName As Integer
Dim appNameEXE As String
Dim AppLoc As String
Dim AppLocLen As Integer


Private Sub cmdCancel_Click()
Unload Me
Unload frmDownload
Unload frmUninstall
End

End Sub

Private Sub Getpath_SYSTEM()
Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                        
Dim Temp                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
                         'a variable for holding the the output of the function
Temp = GetSystemDirectory(WindirS, 255)      'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, Temp)                 'holds final path
LBLwin(4).Caption = Result
End Sub

Private Sub cmdInstall_Click()
cmdInstall.Enabled = False
If InstallFrom = "Web" Then 'if the files are online then
LoadWebIni
Else
LoadIni                     'if the files are on this computer or a disk then
End If
If UpdateLog = True Then SaveLog
End Sub




Private Sub Form_Load()
Me.Show
UpdateFolder = App.Path & "\Files"
On Error Resume Next
Pic.Picture = LoadPicture(App.Path & "\Icon.gif")

Getpath_SYSTEM                      'Calls the getpath_SYSTEM sub

  inipath$ = App.Path + "\install.oil"
 UpdateNo = GetFromINI("Update", "Update", inipath$)
 UpdateLog = GetFromINI("Log", "Update", inipath$)
'Load the number of files to install
 nofiles = GetFromINI("NoFiles", "No", inipath$)
 InstallFrom = GetFromINI("Type", "Type", inipath$)
 
  AppName = GetFromINI("AppName", "Name", inipath$) ' get the app name form ini
 lblFileInstalling.Caption = AppName
 
 appNameEXE = AppName & ".exe"
 LenAppName = Len(appNameEXE)
'find the app name in registry and get installed location
 AppLocation = GetAppPath(AppName & ".exe")
  txtFileInstalling.Text = AppLocation
'find the current installed version
  CurrentV = GetProductVersion(AppLocation)
lblV.Caption = CurrentV
'get rid of then appname and .exe, but leave the "\"
 AppLocLen = Len(AppLocation)
 AppLocLen = AppLocLen - LenAppName
 AppLocation = Left(AppLocation, AppLocLen)
 
 If AppLocation = "" Then   'if the app wasn't installed before don't install it now
 MsgBox AppName & " must be installed to use this update", vbOKOnly
 End
 End If
 If CurrentV = vbNullString Then
 well = MsgBox("The current version number of " & AppName & " could not be found." & vbNewLine & "Continue with the update?", vbYesNo)
    If well = vbYes Then
    GoTo continue:
    Else
    End
    End If
 End If
continue:
 If CurrentV > UpdateNo Then
 MsgBox "You have the most current Version" 'if the current app version is the same or newer don't install
 End
 End If

 Me.Caption = "Fuel for " & AppName
 lblTitle.Caption = "This will install " & AppName & " Version " & UpdateNo _
                    & " to it's default Directory" & vbCrLf & "Click ""Install"" to begin, " & " ""Cancel"" to Exit"
                  
 
 lblFiles.Caption = AppLocation & "ST6UNST.LOG"

 lblNoFiles.Caption = nofiles
 hold
End Sub

Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function

Public Sub LoadIni()
  inipath$ = App.Path + "\install.oil"

Dim i As Integer                    'number of files
Dim FName() As String               'array to split
Dim exten As String                 'file extension
Dim F As String                     'from ini file
For i = 1 To nofiles                'loop thru files to install
 F = GetFromINI("Files", "File" & i, inipath$)
        FName = Split(F, "/")       'split the ini string
        x = LBound(FName)
        FileName = FName(x)

        x = UBound(FName)
        loc = FName(x)
    exten = Right(FileName, 3)      ' get file extensions
On Error GoTo errr:
Select Case loc
    Case "AppPath"
   
        CopyFile UpdateFolder & "\" & FileName, AppLocation & "\" & FileName, 1
        txtList.Text = txtList.Text & "Copied..." & FileName & vbCrLf
        'we don't want to add the original exe to the uninstall log
        If FileName <> (AppName & ".exe") Then UpdateUninstallLog

    Case "System"
       
        RegDLL_OCX UpdateFolder & "\" & FileName, , True
        txtList.Text = txtList.Text & "Registered..." & FileName & vbCrLf

        CopyFile UpdateFolder & "\" & FileName, Result & "\" & FileName, 1
        txtList.Text = txtList.Text & "Copied..." & FileName & vbCrLf
        UpdateUninstallLog

    Case "NewFolder"
    On Error Resume Next
        
        MkDir AppLocation & "\" & FileName
        txtList.Text = txtList.Text & "New Folder..." & FileName & vbCrLf
        UpdateUninstallLog
        
    Case Else
    CopyFile UpdateFolder & "\" & FileName, AppLocation & "\" & loc & "\" & FileName, 1
    txtList.Text = txtList.Text & "Copied..." & FileName & vbCrLf
UpdateUninstallLog
    End Select
    
lblFileInstalling.Caption = "Coppying..." & FileName
    FileName = ""
'add files coppies to the list

Next i
txtList.Text = txtList.Text & "Installation Complete"

cmdCancel.Caption = "Exit"
cmdInstall.Enabled = False
Exit Sub
errr:
MsgBox "An error has occured." & vbNewLine & "The installation was not conpleted.", vbOKOnly, "Error"
txtList.Text = ""
txtList = "Update was Not complete"

End Sub


Public Sub LoadWebIni()
  inipath$ = App.Path + "\install.oil"
       On Error Resume Next
        MkDir App.Path & "\Files" 'make the "Files" folder to hold the downloaded files
UpdateFolder = App.Path & "\Files"

UpdateLocation = GetFromINI("Type", "From", inipath$)
Dim i As Integer                    'number of files
'Dim loc As String                   'install location
Dim FName() As String               'array to split
Dim FileName As String              'file nane
Dim exten As String                 'file extension
Dim F As String                     'from ini file
For i = 1 To nofiles                'loop thru files to install
 F = GetFromINI("Files", "File" & i, inipath$)
        FName = Split(F, "/")       'split the ini string
        x = LBound(FName)           'and get the file name
        FileName = FName(x)

        x = UBound(FName)
        loc = FName(x)
    exten = Right(FileName, 3)      ' get file extensions
On Error GoTo errr:
Select Case loc
    Case "AppPath"   'puts the file in the application path
      'download the file
      frmDownload.DownloadFile UpdateLocation & "/" & FileName, UpdateFolder & "\" & FileName
        'install the file
        CopyFile UpdateFolder & "\" & FileName, AppLocation & "\" & FileName, 1
        txtList.Text = txtList.Text & "Copied..." & FileName & vbCrLf
        'we don't want to add the original exe to the uninstall log it's already there
        If FileName <> (AppName & ".exe") Then UpdateUninstallLog

    Case "System" 'puts the file in the windows system folder and registers it
        frmDownload.DownloadFile UpdateLocation & "/" & FileName, UpdateFolder & "\" & FileName

        RegDLL_OCX UpdateFolder & "\" & FileName, , True
        txtList.Text = txtList.Text & "Registered..." & FileName & vbCrLf

        CopyFile UpdateFolder & "\" & FileName, Result & "\" & FileName, 1
        txtList.Text = txtList.Text & "Copied..." & FileName & vbCrLf
        UpdateUninstallLog

    Case "NewFolder"  'makes a new folder if needed
    On Error Resume Next
        
        MkDir AppLocation & "\" & FileName
        txtList.Text = txtList.Text & "New Folder..." & FileName & vbCrLf
        UpdateUninstallLog
    Case Else
    frmDownload.DownloadFile UpdateLocation & "/" & FileName, UpdateFolder & "\" & FileName

    CopyFile UpdateFolder & "\" & FileName, AppLocation & "\" & loc & "\" & FileName, 1
    txtList.Text = txtList.Text & "Copied..." & FileName & vbCrLf
UpdateUninstallLog
    End Select
    
lblFileInstalling.Caption = "Coppying..." & FileName
    FileName = ""
'add files coppies to the list
        

Next i
txtList.Text = txtList.Text & vbCrLf & "Deleteing Temporary Files"
KillFiles
txtList.Text = txtList.Text & vbCrLf & "Installation Complete"
cmdCancel.Caption = "Exit"
cmdInstall.Enabled = False

Exit Sub
errr:
MsgBox "An error has occured." & vbNewLine & "The installation was not conpleted.", vbOKOnly, "Error"
txtList.Text = ""
txtList = "Update was Not complete"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
Unload frmDownload
tempLog = ""
End

End Sub


Public Sub KillFiles() 'delete the donloaded file after the install
  inipath$ = App.Path + "\install.oil"
       On Error Resume Next
UpdateFolder = App.Path & "\Files"

Dim i As Integer                    'number of files
Dim FName() As String               'array to split
Dim FileName As String              'file nane
Dim F As String                     'from ini file
For i = 1 To nofiles                'loop thru files to install
 F = GetFromINI("Files", "File" & i, inipath$)
        FName = Split(F, "/")       'split the ini string
        x = LBound(FName)           'and get the file name
        FileName = FName(x)

   Kill (UpdateFolder + "\" + FileName) 'delete downloaded file
 txtList.Text = txtList.Text & vbCrLf & FileName & " Deleted"
    FileName = ""
Next i
RmDir App.Path & "\Files"

End Sub
