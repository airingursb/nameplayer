VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "’“ªÿ√‹¬Î"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4170
   LinkTopic       =   "Form5"
   ScaleHeight     =   3525
   ScaleWidth      =   4170
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command2 
      Caption         =   "∑µªÿ"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»∑∂®"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
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
         Name            =   "ÀŒÃÂ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Caption         =   "¥    ∞∏£∫"
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "√‹¬ÎÃ· æ£∫"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "’À    ∫≈£∫"
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "√‹¬Î’“ªÿœµÕ≥"
      BeginProperty Font 
         Name            =   "¡• È"
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

Private Sub Command1_Click()
Dim Name As String, Num As String
Adodc1.RecordSource = "◊¢≤·"
Adodc1.Refresh
Adodc1.Recordset.Find "’À∫≈=" & Text1.Text
If Text2.Text = "" Then
MsgBox "’“≤ªªÿ√‹¬Î"
Else
If Text3.Text = Adodc1.Recordset!¥∞∏ Then
MsgBox "√‹¬Î «" & Adodc1.Recordset!√‹¬Î
Unload Me
Form4.Show
Else
MsgBox "¥∞∏¥ÌŒÛ"
End If
End If
Text3.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
Connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\zc.mdb;Jet OLEDB:Database password=123"
Adodc1.ConnectionString = Connstring
Adodc1.RecordSource = "◊¢≤·"
Adodc1.Refresh
Text1.DataField = ""
Text2.DataField = ""
Text1.Text = ""
Text2.Text = ""
Adodc1.Recordset.MoveFirst
End Sub


Private Sub Text1_LostFocus()
Dim Name As String, Num As String
Adodc1.RecordSource = "◊¢≤·"
Adodc1.Refresh
Adodc1.Recordset.Find "’À∫≈=" & Text1.Text
If Adodc1.Recordset.EOF Then
MsgBox "’À∫≈≤ª¥Ê‘⁄"
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Text2_GotFocus()
Dim Name As String, Num As String
Adodc1.RecordSource = "◊¢≤·"
Adodc1.Refresh
Adodc1.Recordset.Find "’À∫≈=" & Text1.Text
Text2.Text = Adodc1.Recordset!√‹¬ÎÃ· æ
End Sub
