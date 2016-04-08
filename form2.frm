VERSION 5.00
Begin VB.Form form2 
   Caption         =   "登录"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   ScaleHeight     =   2700
   ScaleWidth      =   4335
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   840
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
i = i + 1
If Trim(Text1.Text) = "Airing" And Trim(Text2.Text) = "123" Then
MsgBox "登录成功", 48 + 1, "提示"
Form1.Show
Else
MsgBox "输出错误，请重新输入", 32 + 1, "提示"
End If
If i = 3 Then
MsgBox "对不起，您无权使用本系统！", 16 + 1, "提示"
End
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
form2.Show
Text1.SetFocus
Text2.PasswordChar = "*"
End Sub

