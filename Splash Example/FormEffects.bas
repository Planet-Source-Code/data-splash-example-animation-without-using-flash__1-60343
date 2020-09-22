Attribute VB_Name = "FormEffects"
Option Explicit

'declarations needed for rounding the form corners
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function CreateRoundRectRgn Lib _
    "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function SetWindowRgn Lib _
    "user32" (ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib _
    "gdi32" (ByVal hObject As Long) As Long


'declarations needed for moving the form
Public Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
ByVal lParam As Long) As Long

'const for rounding the form corners
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

'this enables you to use labels to move the form using _MouseDown
Sub MoveForm(TheForm As Form)
'put: MoveForm Me in the _MouseDown part of the sub
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

' Just a simple pause function
Public Function Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Function
