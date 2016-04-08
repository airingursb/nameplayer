VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "µ«¬ºœµÕ≥"
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
   Begin VB.Frame Frame1 
      Caption         =   "«Î ‰»Î’À∫≈∫Õ√‹¬Î"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "”ŒøÕµ«¬º"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ÕÀ    ≥ˆ"
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "µ«    ¬º"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Caption         =   "’“ªÿ√‹¬Î"
         BeginProperty Font 
            Name            =   "∫⁄ÃÂ"
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
         Caption         =   "◊¢≤·’À∫≈"
         BeginProperty Font 
            Name            =   "∫⁄ÃÂ"
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
         Caption         =   "√‹¬Î£∫"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "’À∫≈£∫"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "–’√˚¥Û¿÷∂∑v1.4µ«¬ºœµÕ≥"
      BeginProperty Font 
         Name            =   "¡• È"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Connstring

Private Sub Command1_Click()
If Trim(Text1.Text) = "Airing" And Trim(Text2.Text) = "071515" Then
MsgBox "π‹¿Ì‘±µ«¬º≥…π¶£°"
Form3.Show
Exit Sub
End If
On Error Resume Next
Dim Name As String, Num As String
Adodc1.RecordSource = "◊¢≤·"
Adodc1.Refresh
Adodc1.Recordset.Find "’À∫≈=" & Text1.Text
If Adodc1.Recordset.EOF Or Adodc1.Recordset!√‹¬Î <> Text2.Text Then
    Text1.DataField = ""
    Text2.DataField = ""
    Text1.Text = ""
    Text2.Text = ""
    Adodc1.Recordset.MoveFirst
    MsgBox "’À∫≈ªÚ√‹¬Î¥ÌŒÛ£°", vbOKOnly + vbCritical
Else
    MsgBox "µ«¬º≥…π¶£°"
    Player1 = Adodc1.Recordset!”√ªß√˚
    ID = Val(Text1.Text)
    Form1.Show
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

Private Sub Label4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Label5_Click()
Form5.Show
Unload Me
End Sub
