VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "�û���"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2040
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "����"
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

Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("�˺�") = Text1.Text
Adodc1.Recordset.Fields("�û���") = Text2.Text
Adodc1.Recordset.Fields("����") = Text3.Text
Adodc1.Recordset.Fields("������ʾ") = Text5.Text
Adodc1.Recordset.Fields("��") = Text6.Text
Adodc1.Recordset.Fields("ʤ��") = "0"
Adodc1.Recordset.Fields("ʧ��") = "0"
Adodc1.Recordset.Fields("�ȼ�") = 1
Adodc1.Recordset.Fields("����֮��") = "1"
Adodc1.Recordset.Fields("��Ǯ") = "0"
Adodc1.Recordset.Fields("С��ҩ") = "0"
Adodc1.Recordset.Fields("����ҩ") = "1"
Adodc1.Recordset.UpdateBatch
Adodc1.Recordset.MoveLast
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
    MsgBox "�˺ű���Ϊ��λ���ϵ�����"
    Text1.Text = ""
End If
If Len(Text1.Text) < 6 Then
    MsgBox "�˺ű���Ϊ��λ���ϵ�����"
    Text1.Text = ""
End If
On Error Resume Next
Adodc1.RecordSource = "ע��"
Adodc1.Refresh
Adodc1.Recordset.Find "�˺�=" & Text1.Text
If Adodc1.Recordset.EOF Then
    MsgBox "��ϲ�㣬���˺ſ���ʹ�ã�"
Else
    Text1.DataField = ""
    Text1.Text = ""
    Text4.Text = ""
    Adodc1.Recordset.MoveFirst
    MsgBox "�˺Ų���ʹ�ã�������ע�ᣡ", vbOKOnly + vbCritical
End If
End Sub


Private Sub Form_Load()
Connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\zc.mdb;Jet OLEDB:Database password=123"
Adodc1.ConnectionString = Connstring
Adodc1.RecordSource = "ע��"
Adodc1.Refresh
Text1.DataField = ""
Text2.DataField = ""
Text3.DataField = ""
Text5.DataField = ""
Text6.DataField = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
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



