Attribute VB_Name = "Module1"
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Global FileToAdd As String
Global NumberOfFiles As Integer
Global RemDrive As String
Global RemDir As String
Global NoNewFolders As Integer
Global NewFolderName As String
Global AppName As String
Global AppVersion As String
'store the filepath of the files added for copying later
Global FileNameSource(1 To 40) As String 'sets the max number of file to add at 40
'store the name of the file for copying later
Global FileName(1 To 40) As String 'if change to FileNameSource, much change this as well
Global InstallerIcon As String
Global InstallType As String
Global InstallFrom As String
'used for copying frmMain.listFile to WriteTheFile Sub
Global FilePath As String
Global tempFileName As String
Global CanCompress As Boolean 'can the files be compressed choosen if "install from the web" or "install from a File"
Global DoCompress As Boolean 'did you decide to compress the files
Global NoLog As Boolean
Global addList As Long

Public Sub WriteTheFile()
'if we haven't added any files to the oil list write the header first
If NumberOfFiles = 0 Then
frmOilIni.rtfOil.SelText = "[Files]" & vbNewLine
NumberOfFiles = 1
frmOilIni.rtfOil.SelText = "File" & NumberOfFiles & "=" & FileToAdd & vbNewLine
Else
NumberOfFiles = NumberOfFiles + 1
frmOilIni.rtfOil.SelText = "File" & NumberOfFiles & "=" & FileToAdd & vbNewLine
End If
FileToAdd = ""
'stores the files we added to an FileNameSource array to use to copy files to Oil folder
FileNameSource(NumberOfFiles) = FilePath
'stores the file name we added to an FileName array to use to copy files to Oil folder
FileName(NumberOfFiles) = tempFileName
End Sub

Public Sub Remember()
RemDir = frmAddFile.Dir.Path
RemDrive = frmAddFile.Drive1.Drive
End Sub


Public Sub reSet()
With frmMain
.cmdAppName.Enabled = True
.cmdSave.Enabled = False
.rtfOil.Text = vbNullString
.cmdSave.Enabled = False
End With
End Sub

Public Sub WriteInstallType()
Select Case InstallType
Case "Web"
frmOilIni.rtfOil.SelText = "[Type]" & vbNewLine
frmOilIni.rtfOil.SelText = "Type=Web" & vbNewLine
frmOilIni.rtfOil.SelText = "From=" & InstallFrom & vbNewLine

Case "File"
frmOilIni.rtfOil.SelText = "[Type]" & vbNewLine
frmOilIni.rtfOil.SelText = "Type=File" & vbNewLine
frmOilIni.rtfOil.SelText = "From=" & vbNewLine

End Select
End Sub


Public Sub CopyFiles()
On Error Resume Next
Dim c As Integer

MkDir frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" 'make the Oil folder
'make the "Files" folder in the new Oil folder to hold our new files we will copy
MkDir frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" & "\" & "Files"
ThePath = frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" & "\" & "Files" 'set the path to the new oil folder for copying

'loop through the NumberOfFiles and copy the the files to the new oil folder
For c = 1 To NumberOfFiles
CopyFile FileNameSource(c), ThePath & "\" & FileName(c), 1

'Me.txt1.SelText = "coppied..  " & FileNameSource(c) & "  to  " & ThePath & "\" & FileName(c) & vbNewLine
Next c

  'save the text file
Open frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" & "\install.oil" For Output As #1
    Print #1, frmOilIni.rtfOil.Text
    Close #1
'copy the update installer
CopyFile App.Path & "\Fuel.exe", frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL" & "\Fuel.exe", 1
'copy the image if it was selected
If InstallerIcon <> "" Then CopyFile InstallerIcon, frmMain.Dir.Path & "\" & Replace(AppName, ".exe", vbNullString) & "OIL\icon.gif", 1


End Sub




Public Sub UpdateLog() 'write if fuel installer will update ST6UNST.LOG
If NoLog = False Then
frmOilIni.rtfOil.SelText = vbNewLine & "[Log]" & vbNewLine & "Update=True"
ElseIf NoLog = True Then
frmOilIni.rtfOil.SelText = vbNewLine & "[Log]" & vbNewLine & "Update=False"
End If
End Sub


Public Sub LoadPackedList(Path As String, Combo As ComboBox)
Dim What As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Combo.AddItem What$
    Wend
    Close #1
End Sub

Public Sub SavePackedList()
If addList = -1 Then
frmMain.txtCombo.SelText = vbNewLine & frmMain.cmbApp.Text
Open App.Path & "\PackedList.txt" For Output As #1
    Print #1, frmMain.txtCombo.Text
    Close #1

End If

End Sub


