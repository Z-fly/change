Attribute VB_Name = "Module1"
Option Base 1
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public i As Long, X As Long, Y As String, Z As String
Public counts As Long
Public strd1() As String
Public a(100000) As String
Public TT As Single, TL As Single, TW As Single, TH As Single
Public CMT As Single, CML As Single, CMW As Single, CMH As Single
