<div align="center">

## Launch any file\!


</div>

### Description

With this little function you can LAUNCH ANY TYPE OF FILE that windows reconizes. This means that you can open any type of file that windows has the application for it. For example: You can run a file named "movie.rm" (RealPlayer fomat), but if you don't have Real Player installed, the function just won't do anything! No bugs at all!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Yossi R](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/yossi-r.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/yossi-r-launch-any-file__1-4890/archive/master.zip)

### API Declarations

```
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```


### Source Code

```
Public Function Win32Keyword(ByVal URL As String) As Long
weburl = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
'For example: put the next code under a commad button:
Private Sub Command1_Click()
win32keyword("C:\bla\bla\movie.rm")
End Sub
```

