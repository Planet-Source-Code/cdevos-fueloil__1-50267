VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOilIni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Oil List"
   ClientHeight    =   5115
   ClientLeft      =   11235
   ClientTop       =   2580
   ClientWidth     =   3615
   ControlBox      =   0   'False
   Icon            =   "frmOilIni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtfOil 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8493
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmOilIni.frx":030A
   End
End
Attribute VB_Name = "frmOilIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
frmMain.SetFocus
End Sub

Private Sub Form_Load()
Me.Left = (frmMain.Left + frmMain.Width) + 50
Me.Top = frmMain.Top
End Sub
