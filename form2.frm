VERSION 5.00
Begin VB.Form form2 
   Caption         =   "��¼"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   ScaleHeight     =   2700
   ScaleWidth      =   4335
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���룺"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�û�����"
      BeginProperty Font 
         Name            =   "����"
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
MsgBox "��¼�ɹ�", 48 + 1, "��ʾ"
Form1.Show
Else
MsgBox "�����������������", 32 + 1, "��ʾ"
End If
If i = 3 Then
MsgBox "�Բ�������Ȩʹ�ñ�ϵͳ��", 16 + 1, "��ʾ"
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

