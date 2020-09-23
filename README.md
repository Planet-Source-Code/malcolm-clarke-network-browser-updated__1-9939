<div align="center">

## Network Browser - Updated<br/>by Malcolm Clarke

</div>

### Description

I needed a way to select a network share, As I couldn't find any source I had to put this together. So now I am sharing it for others,

enjoy

### More Info

nothing, it couldn't be simpler

string of the chosen path

### API Declarations

private const error_success as long = 0
private const max_path as long = 260
private const csidl_network as long = &h12
private const bif_returnonlyfsdirs as long = &h1
private const bif_browseforcomputer as long = &h1000
private type browseinfo 'bi
 howner as long
 pidlroot as long
 pszdisplayname as string
 lpsztitle as string
 ulflags as long
 lpfn as long
 lparam as long
 iimage as long
end type
private declare function shgetpathfromidlist lib "shell32.dll" _
 alias "shgetpathfromidlista" _
 (byval pidl as long, _
 byval pszpath as string) as long
private declare function shbrowseforfolder lib "shell32.dll" _
 alias "shbrowseforfoldera" _
 (lpbrowseinfo as browseinfo) as long
private declare function shgetspecialfolderlocation _
 lib "shell32.dll" _
 (byval hwndowner as long, _
 byval nfolder as long, _
 pidl as long) as long

### Source Code

```
public function getbrowsenetworkshare(byval hw as variant) as string
  'returns only a valid share on a network server or workstation
  ' hw is a forms hwnd
  ' call: text1.text = getbrowsenetworkshare(me.hwnd)
  dim bi as browseinfo
  dim pidl as long
  dim spath as string
  dim pos as integer
  if shgetspecialfolderlocation(0, csidl_network, pidl) = error_success then
    with bi
      .howner = hw
      .pidlroot = pidl
      .pszdisplayname = space$(max_path)
      .lpsztitle = "select a network computer or share."
      .ulflags = bif_returnonlyfsdirs
    end with
    'show the browse dialog
    pidl = shbrowseforfolder(bi)
    if pidl <> 0 then
      spath = space$(max_path)
      if shgetpathfromidlist(byval pidl, byval spath) then
        pos = instr(spath, chr$(0))
        getbrowsenetworkshare = left$(spath, pos - 1)
      end if
    else:
      getbrowsenetworkshare = "\\" & bi.pszdisplayname
    end if 'if pidl
  end if 'if shgetspecialfolderlocation
end function
```

