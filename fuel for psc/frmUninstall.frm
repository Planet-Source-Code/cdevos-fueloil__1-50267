VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmUninstall 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8085
   ClientLeft      =   15
   ClientTop       =   -180
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox txtLog 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmUninstall.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmUninstall.frx":0082
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLine 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   7800
      Width           =   1095
   End
End
Attribute VB_Name = "frmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
GetCurrentLine Me.txtLog
End Sub

Private Sub txtLog_SelChange()
GetCurrentLine Me.txtLog
TheLine = GetCurrentLine(Me.txtLog)
End Sub
