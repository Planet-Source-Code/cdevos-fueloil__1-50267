Attribute VB_Name = "Module4"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9

Global NewLine As Long
Global loc As String            'install location
Global AppLocation As String    'where the app is installed
Global AppName As String        'name of the app
Global FileName As String       'file nane
Global UpdateFolder As String   'folder containing this installer and oil file
Global Result As String         'windows system dir
Global UpdateLog As Boolean     'update the uninstall log for the app


Public Sub UpdateUninstallLog()
'The position at which the text is added to the ST6UNST.LOG
'is very important.  Changing the position can cause the app to unistall
'with errors or not at all. When added in the current position, as below, after updating
'I have had all the apps I tested uninstall correctly

If UpdateLog = True Then

'find the were we want to insert the action
Select Case loc
Case "AppPath"
frmUninstall.txtLog.Find ("ACTION: CreateDir: " & """" & Left(AppLocation, Len(AppLocation) - 1) & """")
NewLine = frmUninstall.lblLine.Caption + 1  'advance to the next blank line
SetCursorAtLine NewLine, frmUninstall.txtLog 'move to the  line
'insert the action
frmUninstall.txtLog.SelText = vbNewLine & "ACTION: PrivateFile: " & """" & AppLocation & FileName & """"
frmUninstall.txtLog.SelText = vbNewLine & "(Updated by Fuel Installer -- new file copied)" & vbNewLine

Case "System"
frmUninstall.txtLog.Find "ACTION: SharedFile:"
NewLine = frmUninstall.lblLine.Caption + 2  'advance to the next blank line
SetCursorAtLine NewLine, frmUninstall.txtLog 'move to the line
'insert the action
frmUninstall.txtLog.SelText = vbNewLine & "ACTION: SharedFile: " & """" & Result & "\" & FileName & """"
frmUninstall.txtLog.SelText = vbNewLine & "(Updated by Fuel Installer -- new file copied)" & vbNewLine

Case "NewFolder"
frmUninstall.txtLog.Find ("ACTION: CreateDir: " & """" & Left(AppLocation, Len(AppLocation) - 1) & """")
NewLine = frmUninstall.lblLine.Caption + 1  'advance to the next blank line
SetCursorAtLine NewLine, frmUninstall.txtLog 'move to the  line
'insert the action
frmUninstall.txtLog.SelText = vbNewLine & "ACTION: CreateDir: " & """" & AppLocation & FileName & """" & vbNewLine


Case Else
frmUninstall.txtLog.Find ("ACTION: CreateDir: " & """" & AppLocation & loc & """")
NewLine = frmUninstall.lblLine.Caption + 1  'advance to the next blank line
SetCursorAtLine NewLine, frmUninstall.txtLog 'move to the line
'insert the action
frmUninstall.txtLog.SelText = vbNewLine & "ACTION: PrivateFile: " & """" & AppLocation & loc & "\" & FileName & """"
frmUninstall.txtLog.SelText = vbNewLine & "(Updated by Fuel Installer -- new file copied)" & vbNewLine

End Select
SetCursorAtLine 1, frmUninstall.txtLog
End If
End Sub

Public Sub hold()
On Error GoTo done:
Dim tempLog As String
tempLog = Form1.lblFiles.Caption
'frmUninstall.Show

frmUninstall.txtLog.LoadFile tempLog

Exit Sub
done:
Select Case Err
Case "75"
    Dim NL
    NL = MsgBox("The Unistall Log file, ST6UNST.LOG, for " & AppName & " could not be found" & vbNewLine & "Proceed without updating the ST6UNST.LOG?", vbYesNoCancel)
    Select Case NL
    Case vbYes
         Exit Sub
    Case vbNo
          ReOpen
        Exit Sub
    Case vbCancel
        Exit Sub
        
    End Select

End Select

End Sub

Public Sub ReOpen()
On Error GoTo done:
 Form1.cmFindLog.ShowOpen
  frmUninstall.txtLog.LoadFile Form1.cmFindLog.FileName
   cmdNext.Enabled = True
Exit Sub
done:

End Sub

Public Function GetCurrentLine(RichTextBox As RichTextBox)
    Dim CurrentLine As Long
    CurrentLine = SendMessage(RichTextBox.hwnd, EM_LINEFROMCHAR, -1, 0&) + 1
    frmUninstall.lblLine.Caption = Format(CurrentLine, "###,###,###,###")
End Function


Public Sub SetCursorAtLine(WhichLine As Long, WhichRTFText As RichTextBox)
'**********************************************
'thanks to Matthew Brown, MMComputers for this sub
'**********************************************
Dim Estimate As Long, StartP As Long, EndP As Long
Dim NumChars As Long

With WhichRTFText
    ' Maximum the estimate can be!
    NumChars = Len(.Text)

    ' Its already going to be on the right line!
    If NumChars = 0 Then
        Exit Sub
    End If
    
    ' Check if the given line is out of bounds, or Line 1
    If WhichLine <= 1 Then
        .SelStart = 0
        .SelLength = 0
        Exit Sub
    ElseIf WhichLine > (.GetLineFromChar(NumChars) + 1) Then
        .SelStart = NumChars
        .SelLength = 0
        Exit Sub
    End If
        
    ' Make first estimate
    Estimate = Int(NumChars / 2)
    StartP = 1
    EndP = NumChars

    Dim Finalised As Long ' This is not important - see later

    Do
        If WhichLine < (.GetLineFromChar(Estimate) + 1) Then
            ' estimate too big, refine...
            StartP = StartP
            EndP = Estimate
            Estimate = StartP + Int((EndP - StartP) / 2)
        ElseIf WhichLine > (.GetLineFromChar(Estimate) + 1) Then
            ' estimate too small, refine...
            StartP = Estimate
            EndP = EndP
            Estimate = StartP + Int((EndP - StartP) / 2)
        Else ' is equal! We've found the line
            Finalised = Estimate
            ' Although we know a character IN the line,
            ' this Do...Loop finds the first character on the line
            Do
                Finalised = Finalised - 1
                If Finalised = 0 Then
                    'Finalised = 1
                    .SelStart = Finalised
                    .SelLength = 0
                    Exit Do
                Else
                    If (.GetLineFromChar(Finalised) + 1) < WhichLine Then
                        Finalised = Finalised + 1
                        .SelStart = Finalised
                        .SelLength = 0
                        Exit Do
                    End If
                End If
            Loop
            Exit Do
        End If
    Loop
End With
End Sub


Public Sub SaveLog()
Open AppLocation & "ST6UNST.LOG" For Output As #1
    Print #1, frmUninstall.txtLog.Text
    Close #1
End Sub
