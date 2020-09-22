Attribute VB_Name = "Module1"
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)

Const EWX_FORCE = 4
Const EWX_LOGOFF = 0
Const EWX_REBOOT = 2
Const EWX_SHUTDOWN = 1
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Const VK_SPACE = &H20
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub MakeTopMost(TheForm As Form)
    SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MoveWithoutCap(ObjName As Object)
Dim lngReturnValue As Long
Call ReleaseCapture
lngReturnValue = SendMessage(ObjName.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub MakeCenter(lngHwnd As Form)
    lngHwnd.Top = (Screen.Height - lngHwnd.Height) / 2
    lngHwnd.Left = (Screen.Width - lngHwnd.Width) / 2
End Sub
