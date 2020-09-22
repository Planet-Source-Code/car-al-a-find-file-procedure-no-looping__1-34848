<div align="center">

## A Find File Procedure No Looping


</div>

### Description

This procedure finds a file and returns the full path, provided a partial path it will search all sub dirs. No looping only 1 API - this is the kind.

I can't really take credit for putting an api to use, just sharing something that you might find useful.
 
### More Info
 
sFile: File name to find.

sRootPath: Path to begin the search in

I know of two limitations:

1. The api has a maximum path depth of 32 directories (this is documented).

If a directory that the current user does not have access rights to resides in the directory that the search starts in the api call will fail without warning (this is NOT documented).

Returns the full file path if found otherwise returns null string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[carñal](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/car-al.md)
**Level**          |Beginner
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/car-al-a-find-file-procedure-no-looping__1-34848/archive/master.zip)

### API Declarations

```
Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long
```


### Source Code

```
Private Function FindFile(sFile As String, sRootPath As String) As String
 ' Search for the file specified and return the full path if found
 Dim sPathBuffer As String
 Dim iEnd As Integer
 'Allocate some buffer space (you may need more)
 sPathBuffer = Space(512)
 If SearchTreeForFile(sRootPath, sFile, sPathBuffer) Then
  'Strip off the null string that will be returned following the path name
  iEnd = InStr(1, sPathBuffer, vbNullChar, vbTextCompare)
  sPathBuffer = Left$(sPathBuffer, iEnd - 1)
  FindFile = sPathBuffer
 Else
  FindFile = vbNullString
 End If
End Function
```

