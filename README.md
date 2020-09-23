<div align="center">

## Detect URLs in RichTextBox \(using EM\_AUTOURLDETECT\)


</div>

### Description

This code snippet makes use of RichEdit 2.0's new EM_AUTOURLDETECT message. When you type in a valid webiste address e.g. www.planetsourcecode.com, it will be coloured in blue, then underlined. When the mouse pointer is over it, it will change to a hand icon, and when you click it, it will open a new browser an navigate to the link.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Megatron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/megatron.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/megatron-detect-urls-in-richtextbox-using-em-autourldetect__1-36414/archive/master.zip)

### API Declarations

```
'ADD THE FOLLOWING TO A MODULE
'************************************************
'
'  Written By: Megatron
'
'  March 2002
'
'  Email: mega__tron@hotmail.com
'
'************************************************
Public Type CHARRANGE
  cpMin As Long
  cpMax As Long
End Type
Public Type NMHDR
  hwndFrom As Long
  wPad1 As Integer
  idfrom As Integer
  code As Integer
  wPad2 As Integer
End Type
Public Type ENLINK
  nm As NMHDR
  msg As Integer
  wPad1 As Integer
  wParam As Integer
  wPad2 As Integer
  lParam As Integer
  chrg As CHARRANGE
End Type
Public Type TEXTRANGE
  chrg As CHARRANGE
  lpstrText As String
End Type
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Const WM_SETFONT = &H30
Public Const WM_USER = &H400
Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_VISIBLE = &H10000000
Public Const ES_MULTILINE = &H4&
Public Const WS_CHILD = &H40000000
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const ENM_LINK = &H4000000
Public Const GWL_WNDPROC = (-4)
Public Const WM_NOTIFY = &H4E
Public Const EN_LINK = &H70B
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public IDC_RICHEDIT As Long
Public WinProcOld As Long
Public hwndRichEdit As Long
Public hModule As Long
Public Function WinProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim tNMH As NMHDR
  Dim tEN As ENLINK
  Dim strText As String
  ' If a notification message is recieved then...
  If wMsg = WM_NOTIFY Then
    RtlMoveMemory tNMH, ByVal lParam, Len(tNMH)
    If (tNMH.hwndFrom = hwndRichEdit) And (tNMH.code = EN_LINK) Then
      RtlMoveMemory tEN, ByVal lParam, Len(tEN)
      If tEN.msg = WM_LBUTTONUP Then
        strText = GetTextRange(tEN.chrg.cpMin, tEN.chrg.cpMax)
        If ShellExecute(hwnd, vbNullString, strText, vbNullString, vbNullString, vbNormalFocus) = 2 Then MsgBox "Link Failed", vbExclamation
      End If
    End If
  End If
  WinProc = CallWindowProc(WinProcOld&, hwnd&, wMsg&, wParam&, lParam&)
End Function
Sub SubClassWnd(hwnd As Long)
  WinProcOld& = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WinProc)
End Sub
Sub UnSubclassWnd(hwnd As Long)
  SetWindowLong hwnd, GWL_WNDPROC, WinProcOld&
  WinProcOld& = 0
End Sub
Public Function GetTextRange(nStart As Long, nEnd As Long) As String
  Dim lLen As Long
  Dim txt As TEXTRANGE
  txt.lpstrText = Space$(2000)
  txt.chrg.cpMax = nEnd
  txt.chrg.cpMin = nStart
  lLen = SendMessage(hwndRichEdit, EM_GETTEXTRANGE, 0, txt)
  Debug.Print lLen
  txt.lpstrText = Left(txt.lpstrText, lLen)
  GetTextRange = txt.lpstrText
End Function
Public Sub SetFont(nSize As Long, sName As String)
  Dim hFont As Long
  hFont = CreateFont(nSize, 400, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sName)
  SendMessage hwndRichEdit, WM_SETFONT, hFont, 0
End Sub
```


### Source Code

```
'ADD THE FOLLOWING TO YOUR FORM
Private Sub Form_Load()
  'Subclass the our main window so we can track when a link is hit
  SubClassWnd hwnd
  IDC_RICHEDIT = 4096
  'Load the richedit 2 library
  hModule = LoadLibrary("Riched20.dll")
  If hModule Then
    'Create the richedit window
    hwndRichEdit = CreateWindowEx(WS_EX_CLIENTEDGE, RICHEDIT_CLASSA, "(Type in a URL)", WS_CHILD Or WS_VISIBLE Or ES_MULTILINE, 32, 32, 200, 200, hwnd, IDC_RICHEDIT, App.hInstance, 0)
    'Set it up, such that it can automatically detect URLs
    SendMessage hwndRichEdit, EM_SETEVENTMASK, 0, ByVal ENM_LINK
    Call SendMessage(hwndRichEdit, EM_AUTOURLDETECT, 1, ByVal 0)
    'Change to a more appropiate font
    SetFont 12, "MS Sans Serif"
  Else
    MsgBox "Cannot initialize RichEdit."
    Unload Me
  End If
End Sub
Private Sub Form_Terminate()
  'Free the library from memory
  FreeLibrary hModule
End Sub
Private Sub Form_Unload(Cancel As Integer)
  'Unsubclass the window
  UnSubclassWnd hwnd
End Sub
```

