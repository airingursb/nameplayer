VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   Caption         =   "�̵�"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6735
   LinkTopic       =   "Form10"
   ScaleHeight     =   4530
   ScaleWidth      =   6735
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "����֮��"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "��Ǯ"
      Height          =   735
      Left            =   5400
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "0"
         DataSource      =   "Adodc1"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��   Ʒ"
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "100"
         Height          =   180
         Index           =   6
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ�"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "0"
         DataSource      =   "Adodc1"
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   90
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   4920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ʣ��"
         Height          =   180
         Index           =   4
         Left            =   3000
         TabIndex        =   5
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ч��"
         Height          =   180
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "HP+50"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "С��ҩ"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   180
      Index           =   7
      Left            =   5400
      TabIndex        =   13
      Top             =   480
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
      Height          =   180
      Left            =   5880
      TabIndex        =   12
      Top             =   480
      Width           =   90
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Connstring
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0: Cmd0
    End Select
End Sub

Private Sub Cmd0()
If Label3 >= 100 Then
Label2(0) = Val(Label2(0)) + 1
Label3 = Val(Label3) - 100
Money = Label3
Else
MsgBox "��Ľ�Ǯ���㣡"
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command3_Click()
Unload Me
Form7.Show
End Sub

Private Sub Form_Load()
Label4.Caption = ID
Connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\zc.mdb;Jet OLEDB:Database password=123"
Adodc1.ConnectionString = Connstring
Adodc1.RecordSource = "ע��"
Adodc1.Refresh
Label2(0).DataField = ""
Label3.DataField = ""
Label4.DataField = ""
'On Error Resume Next
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "�˺�=" & Label4
Label3.Caption = Adodc1.Recordset!��Ǯ
Label2(0).Caption = Adodc1.Recordset!С��ҩ
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a As String
Dim b As String
a = "update ע�� set ��Ǯ= '" & Label3 & "',С��ҩ='" & Label2(0) & "' where �˺�=" & Label4
Call CnSql(a, 2)
b = "select * from ע�� where �˺�=" & Label4
Call CnSql(b, 1)
End Sub
