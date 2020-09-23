<div align="center">

## Network Browser \- Updated


</div>

### Description

I needed a way to select a network share, As I couldn't find any source I had to put this together. So now I am sharing it for others,

enjoy
 
### More Info
 
Nothing, it couldn't be simpler

String of the chosen path


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Malcolm Clarke](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/malcolm-clarke.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/malcolm-clarke-network-browser-updated__1-9939/archive/master.zip)

### API Declarations

```
Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_PATH As Long = 260
Private Const CSIDL_NETWORK As Long = &H12
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Type BROWSEINFO 'BI
 hOwner As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As String
 ulFlags As Long
 lpfn As Long
 lParam As Long
 iImage As Long
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
 Alias "SHGetPathFromIDListA" _
 (ByVal pidl As Long, _
 ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
 Alias "SHBrowseForFolderA" _
 (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation _
 Lib "shell32.dll" _
 (ByVal hwndOwner As Long, _
 ByVal nFolder As Long, _
 pidl As Long) As Long
```


### Source Code

```
Public Function GetBrowseNetworkShare(ByVal hw As Variant) As String
  'returns only a valid share on a network server or workstation
  ' hw is a forms hWnd
  ' call: Text1.Text = GetBrowseNetworkShare(Me.hWnd)
  Dim BI As BROWSEINFO
  Dim pidl As Long
  Dim sPath As String
  Dim pos As Integer
  If SHGetSpecialFolderLocation(0, CSIDL_NETWORK, pidl) = ERROR_SUCCESS Then
    With BI
      .hOwner = hw
      .pidlRoot = pidl
      .pszDisplayName = Space$(MAX_PATH)
      .lpszTitle = "Select a network computer or share."
      .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    'show the browse dialog
    pidl = SHBrowseForFolder(BI)
    If pidl <> 0 Then
      sPath = Space$(MAX_PATH)
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
        pos = InStr(sPath, Chr$(0))
        GetBrowseNetworkShare = Left$(sPath, pos - 1)
      End If
    Else:
      GetBrowseNetworkShare = "\\" & BI.pszDisplayName
    End If 'If pidl
  End If 'If SHGetSpecialFolderLocation
End Function
```

