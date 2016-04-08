VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5595
   LinkTopic       =   "Form3"
   ScaleHeight     =   3540
   ScaleWidth      =   5595
   StartUpPosition =   3  '窗口缺省
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Data\zc.mdb"
Data1.RecordSource = "注册"
End Sub
