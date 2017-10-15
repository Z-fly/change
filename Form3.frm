VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1000
   ScaleMode       =   0  'User
   ScaleWidth      =   1000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "请选择您的密码本"
      Filter          =   "文本文件(*.txt) |*.txt|所有文件(*.*) |*.*"
      FilterIndex     =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "报码转汉字"
      Default         =   -1  'True
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function f(str As String) As String
    Dim TStr As String, re
    TStr = str
    Set re = CreateObject("Vbscript.Regexp")
    re.Pattern = "\D"
    re.IgnoreCase = True
    re.Global = True
    TStr = re.Replace(TStr, "")
    Set re = Nothing
    f = TStr
End Function
Private Sub Form_Load()
    TT = Text1.Top
    TL = Text1.Left
    TW = Text1.Width
    TH = Text1.Height
    CMT = Command1.Top
    CML = Command1.Left
    CMW = Command1.Width
    CMH = Command1.Height
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 1 Or 2
    counts = 0
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.FileName) = 0 Then
        End
    End If
    Open CommonDialog1.FileName For Input As #1
    Do While Not EOF(1)
        counts = counts + 1
        ReDim Preserve strd1(counts) As String
        Line Input #1, strd1(counts)
    Loop
    Close
End Sub
Private Sub Form_Resize()
    Me.ScaleHeight = 1000
    Me.ScaleWidth = 1000
    Text1.Top = TT
    Text1.Left = TL
    Text1.Width = TW
    Text1.Height = TH
    Command1.Top = CMT
    Command1.Left = CML
    Command1.Width = CMW
    Command1.Height = CMH
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
Private Sub Command1_Click()
    Y = f(Text1)
    Z = ""
    For i = 1 To Len(Y) Step 4
        a(i) = Mid(Y, i, 4)
    Next
    For i = 1 To Len(Y)
        For X = 1 To counts
            If a(i) = Mid(strd1(X), 2) Then
                a(i) = Left(strd1(X), 1)
                Exit For
            End If
            DoEvents
    Next X, i
    For i = 1 To Len(Y) Step 4
        Z = Z & a(i)
    Next
    Text1 = Z
    Clipboard.Clear
    Clipboard.SetText Text1
    Text1.SetFocus
End Sub
