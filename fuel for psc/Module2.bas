Attribute VB_Name = "Module2"
Option Explicit


Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const KEY_ALL_ACCESS = &H3F
    Private Const SE_ERR_FNF = 2&
    Private Const SE_ERR_PNF = 3&
    Private Const SE_ERR_ACCESSDENIED = 5&
    Private Const SE_ERR_OOM = 8&
    Private Const SE_ERR_DLLNOTFOUND = 32&
    Private Const SE_ERR_SHARE = 26&
    Private Const SE_ERR_ASSOCINCOMPLETE = 27&
    Private Const SE_ERR_DDETIMEOUT = 28&
    Private Const SE_ERR_DDEFAIL = 29&
    Private Const SE_ERR_DDEBUSY = 30&
    Private Const SE_ERR_NOASSOC = 31&
    Private Const ERROR_BAD_FORMAT = 11&
    Private Const SW_HIDE = 0
    Private Const SW_NORMAL = 1
    Private Const SW_SHOWNORMAL = 1
    Private Const SW_SHOWMINIMIZED = 2
    Private Const SW_SHOWMAXIMIZED = 3
    Private Const SW_RESTORE = 9



Public Function GetAppPath(ByVal AppName As String) As String
    'from the registry if it exists
    'returns vbNullstring if not
    On Error GoTo TheEnd:
    Dim TheResult As Long
    Dim Index As Long
    Dim TheEntry As String
    Dim EntryLength As Long
    Dim TheDataType As Long
    Dim TheByteArray(1 To 1024) As Byte
    Dim DataLength As Long
    Dim ByteValue As String
    Dim i As Integer
    Dim MainKey As Long
    Dim SubKey As String
    Dim mKey As Long


    'If LCase(Right(AppName, 4)) <> ".exe" Then
       ' AppName = AppName & ".exe"
    'End If
    MainKey = HKEY_LOCAL_MACHINE


SubKey = "Software\Microsoft\Windows\CurrentVersion\App Paths\" & AppName
    TheResult = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, mKey)
    If TheResult <> 0 Then Exit Function
    'looked for it and failed
    Index = 0


    Do
        EntryLength = 1024
        DataLength = 1024
        TheEntry = Space(EntryLength)
        TheResult = RegEnumValue(mKey, Index, TheEntry, EntryLength, 0, _
        TheDataType, TheByteArray(1), DataLength)
        'looks like we just have to pass just th
        '     e first element
        'of the array to have it filled...
        If TheResult <> 0 Then Exit Do
        TheEntry = Left(TheEntry, EntryLength)


        If Len(TheEntry) = 0 Then
            'looking for (Default), empty string
            ByteValue = ""


            For i = 1 To DataLength - 1
                ByteValue = ByteValue & Chr(TheByteArray(i))
            Next
            


            If ByteValue <> "" Then
                GetAppPath = LongName(ByteValue)
                RegCloseKey mKey
                Exit Function
            End If
        End If
        Index = Index + 1
    Loop
    GetAppPath = ""
    RegCloseKey mKey
    Exit Function
TheEnd:
    GetAppPath = ""
End Function


Public Function LongName(ShortName As String) As String
    'not my code
    
    Dim Temp As String
    Dim NewString As String
    Dim Searched As Boolean
    Dim i As Integer
    If Len(ShortName) = 0 Then Exit Function
    Temp = ShortName


    If Right(Temp, 1) = "\" Then
        Temp = Left(Temp, Len(Temp) - 1)
        Searched = True
    End If
    On Error GoTo NoFile:


    If InStr(Temp, "\") Then
        NewString = ""


        Do While InStr(Temp, "\")


            If Len(NewString) Then
                NewString = Dir(Temp, 55) & "\" & NewString
            Else
                NewString = Dir(Temp, 55)


                If NewString = "" Then
                    LongName = ShortName
                    Exit Function
                End If
            End If
            On Error Resume Next


            For i = Len(Temp) To 1 Step -1


                If ("\" = Mid(Temp, i, 1)) Then
                    Exit For
                End If
            Next
            Temp = Left(Temp, i - 1)
        Loop
        NewString = Temp & "\" & NewString
    Else
        NewString = Dir(Temp, 55)
    End If
Here:


    If Searched Then
        NewString = NewString & "\"
    End If
    LongName = PrettyPath(NewString)
    Exit Function
NoFile:
    NewString = ""
    Resume Here:
End Function

Public Function PrettyPath(ThePath As String) As String
    On Error GoTo TheEnd:
    Dim Path As String
    Dim Start As Integer
    Dim Temp As String
    Path = ThePath
    Path = LCase(Path)
    Temp = Left(Path, 1)
    Temp = UCase(Temp)
    Path = Temp & Right(Path, Len(Path) - 1)
    'got drive letter pretty
    Start = 1


    Do
        Start = InStr(Start, Path, "\")
        If Start = 0 Then Exit Do
        Mid(Path, Start + 1, 1) = UCase(Mid(Path, Start + 1, 1))
        Start = Start + 1
    Loop While Start < Len(ThePath)
    'put a cap after each backslash
    Start = 1


    Do
        Start = InStr(Start, Path, " ")
        If Start = 0 Then Exit Do
        Mid(Path, Start + 1, 1) = UCase(Mid(Path, Start + 1, 1))
        Start = Start + 1
    Loop While Start < Len(Path)
    'put a cap after each space
    PrettyPath = Path
    Exit Function
TheEnd:
    PrettyPath = ThePath
    'just in case
End Function


