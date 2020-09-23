Attribute VB_Name = "modClock"
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const PI = 3.14159265
Public Sub drag(frm As Form)
'Drag the form
    ReleaseCapture
    SendMessage frm.hWnd, &HA1, 2, 0
End Sub



