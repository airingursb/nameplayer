VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ע��"
   ClientHeight    =   5580
   ClientLeft      =   7455
   ClientTop       =   2280
   ClientWidth     =   4125
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   4125
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text6 
      DataField       =   "��"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "�˺�"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      DataField       =   "������ʾ"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "�û���"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ע��"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "*��    �ţ�"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��    ����"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "������ʾ��"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "*ȷ�����룺"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "*��    �룺"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "*�� �� ����"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ע��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Connstring
Dim db As ADODB.Connection
Dim Rs As ADODB.Recordset

Private Sub Command1_Click()
  Rs.AddNew                                                                '���������
         Rs!�˺� = Text1.Text
         Rs!�û��� = Text2.Text
         Rs!���� = Text3.Text
         Rs!������ʾ = Text5.Text
         Rs!�� = Text6.Text
         Rs!ʤ�� = "0"
         Rs!ʧ�� = "0"
         Rs!�ȼ� = 1
         Rs!����֮�� = "1"
         Rs!��Ǯ = "0"
         Rs!С��ҩ = "0"
         Rs!����ҩ = "1"
         Rs!���� = "0"
         Rs.Update                                                                '   ��������
         Rs.Close                                                                 '   �رձ��
         db.Close                                                               '   �ر����ݿ�
MsgBox "ע��ɹ���"
Unload Me
Form4.Show
End Sub

Private Sub Command2_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command3_Click()
Dim Sum&
Dim c
Dim Char
Sum = 0
For c = 1 To Len(Text2.Text)
Char = Mid(Text2.Text, c, 1)
If (AscW(Char) > -40870 And AscW(Char) < -19967) Or (AscW(Char) < 40870 And AscW(Char) > 19967) Then
Sum = Sum + 1
End If
Next c
If Sum = 0 Then
MsgBox "�û�������Ϊ���֣�����������������"
Text2.Text = ""
End If
If Sum = 1 Then
MsgBox "�û�������Ϊ�������֣�һ���ֲ��Ϸ�������������"
Text2.Text = ""
End If
End Sub

Private Sub Command4_Click()
If IsNumeric(Text1.Text) = False Then
    MsgBox "�˺ű���Ϊ��λ���ϰ�λ���µ�����"
    Text1.Text = ""
End If
If Len(Text1.Text) < 6 And Len(Text1.Text) >= 8 Then
    MsgBox "�˺ű���Ϊ��λ���ϰ�λ���µ�����"
    Text1.Text = ""
End If
Dim rc As ADODB.Recordset
Dim strsql As String
Set rc = New ADODB.Recordset
strsql = " select * from zc where �˺�=" & Text1
rc.Open strsql, db, adOpenStatic, adLockReadOnly
    If rc.EOF Then
        MsgBox "��ϲ���˺ſ���ʹ��", , "��ʾ"
    Else
        MsgBox "���ź����˺Ų���ʹ�ã�������ע��", vbCritical, "��ʾ"
        Set rc = Nothing
        'rc.Close
End If
End Sub

Private Sub Form_Load()
Dim strsql
Set db = New ADODB.Connection
Set Rs = New ADODB.Recordset
db.ConnectionString = "Provider=SQLOLEDB.1;Password=1123581321;Persist Security Info=True;User ID=hds1010886;Initial Catalog=hds1010886_db;Data Source=hds-101.hichina.com"
db.Open
    strsql = "select * from zc"                                    '�򿪱��
    Rs.Open strsql, db, 3, 3
If db.State = adStateOpen Then
'MsgBox "�ɹ�"
Else
MsgBox "����ʧ��"
End If
End Sub

Private Sub Text2_LostFocus()
Call Command3_Click
End Sub

Private Sub Text3_LostFocus()
If IsNumeric(Text3.Text) = False Then
    MsgBox "�������Ϊ��λ���ϵ�����"
    Text3.Text = ""
End If
If Len(Text3.Text) < 6 Then
    MsgBox "�������Ϊ��λ���ϵ�����"
    Text3.Text = ""
End If
End Sub

Private Sub Text4_LostFocus()
If Trim(Text4.Text) <> Trim(Text3.Text) Then
    MsgBox "������������벻ͬ��"
    Text3.Text = ""
    Text3.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Call Command4_Click
End Sub
