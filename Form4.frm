VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录系统"
   ClientHeight    =   4155
   ClientLeft      =   7215
   ClientTop       =   3210
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "请输入账号和密码"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "游客登录"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "退    出"
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "登    录"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         DataField       =   "密码"
         DataSource      =   "Adodc1"
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   "账号"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "找回密码"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3120
         TabIndex        =   8
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "注册账号"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "账号："
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "姓名大乐斗登录系统"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   2835
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Connstring
Dim db As ADODB.Connection
Dim Rs As ADODB.Recordset
Private Sub Command1_Click()
Call check
If Trim(Text1.Text) = "Airing" And Trim(Text2.Text) = "071515" Then
MsgBox "管理员登录成功！"
Form3.Show
Exit Sub
End If
Dim Name As String, Num As String
Dim Rs As ADODB.Recordset
Dim strsql As String
Dim temp As String
Name = Text1.Text
Num = Text2.Text
Set Rs = New ADODB.Recordset
    strsql = "select * from zc where 账号=" & Name & "  and 密码='" & Num & "'"
    Rs.Open strsql, db, adOpenStatic, adLockReadOnly 'Open table "DBser"
    If Rs.EOF Then
        MsgBox "用户名或密码错误", vbCritical, "提示"
    Else
        ID = Name
        Player1 = Rs.Fields("用户名").Value
        Form1.Show
    End If
End Sub

Private Sub check()
Dim Rs As ADODB.Recordset
Dim strsql As String
Dim Name As String, Num As String
Name = 3214555
Num = 3214555
Set Rs = New ADODB.Recordset
    strsql = "select * from zc where 账号=" & Name & "  and 密码='" & Num & "'"
    Rs.Open strsql, db, adOpenStatic, adLockReadOnly
    If Rs.EOF Then
        MsgBox "已是最新版本v1.5"
    Else
        MsgBox "已有更新，请去官网下载更新。"
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form8.Show
Unload Me
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set Rs = New ADODB.Recordset
db.ConnectionString = "Provider=SQLOLEDB.1;Password=1123581321;Persist Security Info=True;User ID=hds1010886;Initial Catalog=hds1010886_db;Data Source=hds-101.hichina.com"
db.Open
If db.State = adStateOpen Then
'MsgBox "成功"
Else
MsgBox "失败"
End If
End Sub

Private Sub Label4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Label5_Click()
Form5.Show
Unload Me
End Sub
