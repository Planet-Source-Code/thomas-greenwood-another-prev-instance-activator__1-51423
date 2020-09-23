<div align="center">

## Another Prev Instance Activator


</div>

### Description

This code will activate a previous instance of your program if detected. It uses subclassing of the main window.
 
### More Info
 
May crash the ide, so save your work before execution (did not crash mine however, but subclassing can be unstable)

Had to use end to close the next instance (not the prev instance). Using Unload Obj in Form_Load produced an error, it can be used if On Error Resume Next or a Error Handler is produced.


<span>             |<span>
---                |---
**Submitted On**   |2004-02-02 09:22:02
**By**             |[Thomas Greenwood](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-greenwood.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Another\_Pr170288222004\.zip](https://github.com/Planet-Source-Code/thomas-greenwood-another-prev-instance-activator__1-51423/archive/master.zip)

### API Declarations

```
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' ^ Used to change window proc
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ^ Used to call default window proc for window
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
' ^ Send Message to window (Proc)
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
' ^ Used, funnily enough...to flash the window :D
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' ^ Used to find the window if we have a prev instance
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
' ^ Check to see if we have a valid window.
```





