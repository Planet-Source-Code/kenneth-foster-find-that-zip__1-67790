Attribute VB_Name = "modOpendiag"
Option Explicit
    
    Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
    
    Declare Function SHBrowseForFolder Lib "Shell32.dll" Alias _
    "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
    
    Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias _
    "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long
    
    Const DELETE = &H3
    Const ALLOWUNDO = &H40
    
    Declare Function SHFileOperation _
    Lib "Shell32.dll" _
    (FileOp As SHFILEOPSTRUCT) As Long
    
    Type SHFILEOPSTRUCT
    hWnd As Long
    lFunc As Long
    sForm As String
    Sto As String
    iFlags As Integer
    boolAnyOperationAborted As Boolean
    lNameMappings As Long
    sProgressTitle As String
End Type

Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName  As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type

Type SHITEMID
CB As Long
abID As Byte
End Type

Type ITEMIDLIST
mkid As SHITEMID
End Type

Const BIF_RETURNONLYFSDIRS = &H1

Public Function GetBrowseDirectory(Owner As Form) As String
    Dim Bi As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Dim r As Long
    Dim pidl As Long
    Dim tmpPath As String
    Dim pos As Integer
    
    Bi.hOwner = Owner.hWnd
    Bi.pidlRoot = 0&
    Bi.lpszTitle = "Choose a directory from the list."
    Bi.ulFlags = BIF_RETURNONLYFSDIRS
    pidl = SHBrowseForFolder(Bi)
    
    tmpPath = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl, ByVal tmpPath)
    
    If r Then
        pos = InStr(tmpPath, Chr$(0))
        tmpPath = Left(tmpPath, pos - 1)
        
        If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
        GetBrowseDirectory = tmpPath
    Else
        GetBrowseDirectory = ""
    End If
    
End Function

Public Function FileExists(FileName As String) As Boolean
    'This function checks the existance of a file
    On Error GoTo Handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
Handle:
    FileExists = False
End Function

Public Sub SendToBin(sFileName As String)
    Dim FileStruct As SHFILEOPSTRUCT
    Dim x As Long
    
    With FileStruct
        .iFlags = ALLOWUNDO
        .sForm = sFileName
        .lFunc = DELETE
    End With
    x = SHFileOperation(FileStruct)
End Sub
