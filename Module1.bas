Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Public Const SW_SHOWNA = 8

Const BIF_RETURNONLYFSDIRS = &H1
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5

Public Function GetFolder(title As String, parentHwnd As Long) As String
Dim bi As BROWSEINFO
Dim pidl As Long
Dim Folder As String
Folder = String$(255, Chr$(0))
With bi
    If IsNumeric(parentHwnd) Then .hOwner = parentHwnd
    .ulFlags = BIF_RETURNONLYFSDIRS
    .pidlRoot = 0
    If title <> "" Then
        .lpszTitle = title & Chr$(0)
    Else
        .lpszTitle = "Select a Folder" & Chr$(0)
    End If
    
End With

pidl = SHBrowseForFolder(bi)

If SHGetPathFromIDList(ByVal pidl, ByVal Folder) Then
    GetFolder = Left(Folder, InStr(Folder, Chr$(0)) - 1)
Else
    GetFolder = ""
End If

End Function


