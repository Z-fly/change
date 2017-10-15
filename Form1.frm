VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数字汉字双向工程之转变"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   420
      TabIndex        =   3
      ToolTipText     =   "关于"
      Top             =   360
      Width           =   4695
      Begin VB.Label Label1 
         Caption         =   "数字汉字双向工程           Copyright 2014，4，All Rights Reserved 张执 "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "汉字转报码"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "报码转汉字"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "按此退出，按上面的会……"
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub
Private Sub Command2_Click()
    Me.Hide
    Form3.Show
End Sub
Private Sub Command3_Click()
    Me.Hide
    Form2.Show
End Sub
Private Sub Form_Initialize()
    App.Title = ""
End Sub
Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = -1
End Sub
