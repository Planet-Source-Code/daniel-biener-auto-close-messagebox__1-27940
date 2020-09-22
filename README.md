<div align="center">

## Auto close messagebox


</div>

### Description

This function replaces VB's msgbox function and closes itself after the parameter provided number of seconds. The syntax and return values are exactly the same as msgbox except the first parameter is the number of seconds to display. Just add this code to a module (not a cls or frm) in your project and call ACmsgbox. Thanks to Sparq's submission here for help in writing this.

with the added parameter of
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Biener](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-biener.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-biener-auto-close-messagebox__1-27940/archive/master.zip)

### API Declarations

```
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const NV_CLOSEMSGBOX As Long = &H5000&
Private sLastTitle As String
```


### Source Code

```
Public Function ACmsgbox(AutoCloseSeconds As Long, prompt As String, Optional buttons As Long, _
      Optional title As String, Optional helpfile As String, _
      Optional context As Long) As Long
  sLastTitle = title
  SetTimer Screen.ActiveForm.hWnd, NV_CLOSEMSGBOX, AutoCloseSeconds * 1000, AddressOf TimerProc
  ACmsgbox = MsgBox(prompt, buttons, title, helpfile, context)
End Function
Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  Dim hMessageBox As Long
  KillTimer hWnd, idEvent
  Select Case idEvent
  Case NV_CLOSEMSGBOX
    hMessageBox = FindWindow("#32770", sLastTitle)
    If hMessageBox Then
      Call SetForegroundWindow(hMessageBox)
      SendKeys "{enter}"
    End If
    sLastTitle = vbNullString
  End Select
End Sub
```

