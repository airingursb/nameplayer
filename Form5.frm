VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "�һ�����"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4170
   LinkTopic       =   "Form5"
   ScaleHeight     =   3525
   ScaleWidth      =   4170
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��    ����"
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "������ʾ��"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��    �ţ�"
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����һ�ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1890
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Connstring
Dim db As ADODB.Connection
Dim Rs As ADODB.Recordset
Private Sub Command1_Click()
Dim rc As ADODB.Recordset
Dim pass As String
Dim strsql As String
Set rc = New ADODB.Recordset
strsql = " select * from zc where �˺�=" & Text1 & "  and �� ='" & Text3.Text & "' "
rc.Open strsql, db, adOpenStatic, adLockReadOnly
     If rc.EOF Then
         MsgBox "������������", vbCritical, "��ʾ"
   Else
         pass = rc.Fields("����").Value
         MsgBox "��������� " & pass & " �����Ʊ��ܣ�", vbInformation, "������ʾ"
         Set rc = Nothing
         'rc.Close
         pass = ""
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set Rs = New ADODB.Recordset
db.ConnectionString = "Provider=SQLOLEDB.1;Password=1123581321;Persist Security Info=True;User ID=hds1010886;Initial Catalog=hds1010886_db;Data Source=hds-101.hichina.com"
db.Open
If db.State = adStateOpen Then
'MsgBox "�ɹ�"
Else
MsgBox "����ʧ��"
End If
End Sub

Private Sub Text1_LostFocus()
Dim rc As ADODB.Recordset
Dim pass As String
Dim strsql As String
Set rc = New ADODB.Recordset
strsql = " select * from zc where �˺�=" & Text1
rc.Open strsql, db, adOpenStatic, adLockReadOnly
     If rc.EOF Then
         MsgBox "���޴˺�", vbCritical, "��ʾ"
   Else
         Text2.Text = rc.Fields("������ʾ").Value
         Set rc = Nothing
         'rc.Close
End If
End Sub

