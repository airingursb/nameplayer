VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   2520
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Selectsql(SQL As String) As ADODB.Recordset '返回ADODB.Recordset对象
Dim ConnStr As String
Dim Conn As ADODB.Connection
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set Conn = New ADODB.Connection

'On Error GoTo MyErr:
ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;Password=1123581321;Initial Catalog=test;Data Source=localhost" '这是连接SQL数据库的语句
Conn.Open ConnStr
rs.CursorLocation = adUseClient
rs.Open Trim$(SQL), Conn, adOpenDynamic, adLockOptimistic
End Function

Private Sub Command1_Click()
SQL = "SELECT * FROM text WHERE 账号='" & Text1.Text & "' AND 密码='" & Text2.Text & "' "
Set rs = Selectsql(SQL)
End Sub

