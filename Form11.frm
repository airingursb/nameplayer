VERSION 5.00
Begin VB.Form Form11 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "姓名大乐斗 v1.5"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command9 
      Caption         =   "商店"
      Height          =   375
      Left            =   1800
      TabIndex        =   51
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "属性"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   44
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Height          =   180
      Left            =   12000
      TabIndex        =   29
      Top             =   6960
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "复位"
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   " ‖ "
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "》"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   26
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "《"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   6240
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   12240
      Top             =   7320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   5
      Left            =   10920
      TabIndex        =   21
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   4
      Left            =   10920
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   3
      Left            =   10920
      TabIndex        =   19
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   2
      Left            =   10920
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   1
      Left            =   10920
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   0
      Left            =   10920
      TabIndex        =   16
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   4935
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   720
      Width           =   6855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "擂台战"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5040
      TabIndex        =   62
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   4
      Left            =   11280
      TabIndex        =   61
      Top             =   5880
      Width           =   90
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   3
      Left            =   10440
      TabIndex        =   60
      Top             =   6600
      Width           =   90
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   2
      Left            =   10200
      TabIndex        =   59
      Top             =   5880
      Width           =   90
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Index           =   1
      Left            =   10440
      TabIndex        =   58
      Top             =   6240
      Width           =   90
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   8880
      TabIndex        =   57
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "ID："
      Height          =   180
      Index           =   9
      Left            =   8400
      TabIndex        =   56
      Top             =   360
      Width           =   360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "金钱："
      Height          =   180
      Index           =   8
      Left            =   10680
      TabIndex        =   55
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "失败："
      Height          =   180
      Index           =   7
      Left            =   9840
      TabIndex        =   54
      Top             =   6600
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "胜场："
      Height          =   180
      Index           =   6
      Left            =   9840
      TabIndex        =   53
      Top             =   6240
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Lv："
      Height          =   180
      Index           =   5
      Left            =   9840
      TabIndex        =   52
      Top             =   5880
      Width           =   360
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
      Height          =   180
      Index           =   4
      Left            =   1560
      TabIndex        =   50
      Top             =   5760
      Width           =   90
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "金钱："
      Height          =   180
      Index           =   4
      Left            =   960
      TabIndex        =   49
      Top             =   5760
      Width           =   540
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
      Height          =   180
      Index           =   3
      Left            =   840
      TabIndex        =   48
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "失败："
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   47
      Top             =   6480
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Lv："
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   46
      Top             =   5760
      Width           =   360
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
      Height          =   180
      Index           =   2
      Left            =   720
      TabIndex        =   45
      Top             =   5760
      Width           =   90
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "ID："
      Height          =   180
      Index           =   1
      Left            =   1800
      TabIndex        =   43
      Top             =   360
      Width           =   360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "胜场："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   42
      Top             =   6120
      Width           =   540
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   41
      Top             =   6120
      Width           =   90
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
      Height          =   180
      Index           =   0
      Left            =   2400
      TabIndex        =   40
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "游戏选项："
      Height          =   180
      Left            =   3360
      TabIndex        =   39
      Top             =   7080
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "擂主"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10920
      TabIndex        =   38
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "玩家"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   37
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "运气"
      Height          =   180
      Left            =   10200
      TabIndex        =   36
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "敏捷"
      Height          =   180
      Left            =   10200
      TabIndex        =   35
      Top             =   4320
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "速度"
      Height          =   180
      Left            =   10200
      TabIndex        =   34
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "防御"
      Height          =   180
      Left            =   10200
      TabIndex        =   33
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "攻击"
      Height          =   180
      Left            =   10200
      TabIndex        =   32
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "HP值"
      Height          =   180
      Left            =   10200
      TabIndex        =   31
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   10200
      TabIndex        =   30
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "游戏速度："
      Height          =   180
      Left            =   3360
      TabIndex        =   24
      Top             =   6360
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "HP值"
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "攻击"
      Height          =   180
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "速度"
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "敏捷"
      Height          =   180
      Left            =   1200
      TabIndex        =   3
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "运气"
      Height          =   180
      Left            =   1200
      TabIndex        =   2
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "防御"
      Height          =   180
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   360
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim fhy
Dim a As Integer '左边随即优先攻击
Dim b As Integer '右边随即优先攻击
Dim k As Integer '攻击值数目

Dim c As String '左边向右边先攻击
Dim d As String '右边向左边先攻击
Dim c0 As String
Dim d0 As String

Dim e As String '攻击方式
Dim f As Integer '下降属性值

Dim hp(1) As Integer '生命值
Dim gj(1) As Integer '攻击值
Dim fy(1) As Integer '防御值
Dim sd(1) As Integer '速度值
Dim mz(1) As Integer '命中值
Dim yq(1) As Integer '运气值
Dim Tur As Integer '静态变量
Dim Flag As Integer    '攻击方式产生概率
Dim Die
Dim hp0, Jc                '金蝉脱壳
Dim hpc1, hpc2         '回光返照
Dim Round1, Round2  '使用八门遁甲死亡倒计时
Dim R1 As Boolean, R2 As Boolean   '启动死亡倒计时
Dim SP(0 To 60)       '各技能使用上限
Dim Connstring
Dim Money2

Private Sub Command1_Click()
If Money >= 100 And MsgBox("您确定花费100金币购买挑战资格吗？", vbYesNo, "提示") = vbYes Then
    Money = Money - 100
    Label19(4).Caption = Money
Else
    MsgBox "您没有挑战资格！", , "提示"
    Exit Sub
End If
Randomize
Timer1.Enabled = True
Command9.Enabled = False
R1 = False
R2 = False
SP(1) = 0
SP(2) = 0
Lv1 = Val(Label19(2))
Lv2 = 1
Timer1.Interval = 500
Dim lngReturn As Long
Dim I

If Name_Do(Val(Len(Text1) + Len(Text2))) > 0 Then
MsgBox "请输入汉字！", , "提示"
Exit Sub
Else
If Text1 = "" And Text2 = "" Then
MsgBox "请输入名字！", , "提示"
Exit Sub
End If
c = "[" & Text1 & "]" & "向" & "[" & Text2 & "]"
d = "[" & Text2 & "]" & "向" & "[" & Text1 & "]"
c0 = "[" & Text1 & "]" & "向" & "自己"
d0 = "[" & Text2 & "]" & "向" & "自己"
For I = 1 To Len(Trim(Text1))
lngReturn = CLng("&h" & Hex((AscW(Mid(Text1, I, 1)))))
If I = 1 Then
Text4(0).Text = Mid(lngReturn, 1, 3)
Text4(1).Text = Val(Mid(lngReturn, 3, 2) + 30)
Text4(4).Text = Val(Mid(lngReturn, 2, 2) + 50)
End If
If I = 2 Then
Text4(2).Text = Val(Mid(lngReturn, 1, 2) + 30)
Text4(3).Text = Val(Mid(lngReturn, 2, 2) + 40)
End If
Next I

For I = 1 To Len(Trim(Text2))
lngReturn = CLng("&h" & Hex((AscW(Mid(Text2, I, 1)))))
If I = 1 Then
Text5(0).Text = Mid(lngReturn, 1, 3)
Text5(1).Text = Val(Mid(lngReturn, 3, 2) + 30)
Text5(4).Text = Val(Mid(lngReturn, 2, 2) + 50)
End If
If I = 2 Then
Text5(2).Text = Val(Mid(lngReturn, 1, 2) + 30)
Text5(3).Text = Val(Mid(lngReturn, 2, 2) + 40)
End If
Next I
Text4(5).Text = Int(Rnd * 100)
Text5(5).Text = Int(Rnd * 100)

Call sx(0, 0, 0)

Text3.Text = "姓名大作战 VB版" & vbCrLf & vbCrLf
Text3.Text = Text3.Text + Text1 & "  " & "HP：" & Text4(0) & "  " & "攻：" & Text4(1) & "  " & "防：" & Text4(2) & "  " & "速：" & Text4(3) & "  " & "技：" & Text4(4) & "  " & "运：" & Text4(5) & vbCrLf
Text3.Text = Text3.Text + Text2 & "  " & "HP：" & Text5(0) & "  " & "攻：" & Text5(1) & "  " & "防：" & Text5(2) & "  " & "速：" & Text5(3) & "  " & "技：" & Text5(4) & "  " & "运：" & Text5(5) & vbCrLf & vbCrLf

If (Text4(5).Text + Text4(3).Text) > (Text5(5).Text + Text5(3).Text) Then '战斗先机
Tur = 2
ElseIf (Text4(5).Text + Text4(3).Text) < (Text5(5).Text + Text5(3).Text) Then
Tur = 1
Else
MsgBox "无法确定先手，请再来一次！", , "提示"
Text4(5).Text = Int(Rnd * 100)
Text5(5).Text = Int(Rnd * 100)
End If
Timer1.Enabled = True
Command9.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command3_Click()          '减慢
Timer1.Interval = Timer1.Interval - 500
End Sub

Private Sub Command4_Click()          '属性栏
Form6.Show
End Sub

Private Sub Command5_Click()          '加快
Timer1.Interval = Timer1.Interval + 500
End Sub

Private Sub Command6_Click()          '暂停
Timer1.Interval = 0
End Sub

Private Sub Command7_Click()                   '复位键
MsgBox "请重新输入对手姓名", , "提示"
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command8_Click()          '作弊开启
If Text1.Text = "邓国雄" Then
Text4(0).Text = Text4(0).Text + 500
ElseIf Text2.Text = "邓国雄" Then
Text5(0).Text = Text5(0).Text + 500
Else
MsgBox "您无权使用作弊器！", , "警告"
End If
End Sub

Private Sub Command9_Click()
Unload Me
Form10.Show
End Sub

Private Sub Text2_LostFocus()
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
MsgBox "对手名必须为汉字！"
Text2.Text = ""
End If
If Sum = 1 Then
MsgBox "对手名至少为两个字！请重新输入！"
Text2.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
a = Int(Rnd * 20)
b = Int(Rnd * 20)
Flag = Int(Rnd * 100)
Die = Int(Rnd * 100) '金蝉脱壳释放几率
Call fLv(8, Lv1)
Call fLv(9, Lv2)
If Text4(0) <= 0 Then
Call fLv(5, Lv1)
Call fLv(6, Lv1)
If Die >= Jc And SP(3) < 1 Then
Text4(0) = hp0
Text3.Text = Text3.Text + "[" & Text1 & "]" & "使用金蝉脱壳！" + vbCrLf
SP(3) = SP(3) + 1
Else
Text4(0) = 0
Text3.Text = Text3.Text + "[" & Text1 & "]" & "被打败！"
Timer1.Enabled = False
Command9.Enabled = True
Label19(3).Caption = Val(Label19(3)) + 1 '计算失败
Label21(1).Caption = Val(Label21(1)) + 1
Call Lv02
Call fail
End If
Exit Sub

ElseIf Text5(0) <= 0 Then
Call fLv(5, Lv2)
Call fLv(6, Lv2)
If Die >= Jc And SP(4) < 1 Then
Text5(0) = hp0
Text3.Text = Text3.Text + "[" & Text2 & "]" & "使用金蝉脱壳！" + vbCrLf
SP(4) = SP(4) + 1
Else
Text5(0) = 0
Text3.Text = Text3.Text + "[" & Text2 & "]" & "被打败！"
Timer1.Enabled = False
Command9.Enabled = True
Label19(1).Caption = Val(Label19(1)) + 1 '计算胜场
Label21(3).Caption = Val(Label21(3)) + 1
Call Lv
Call win
End If
Exit Sub

ElseIf hpc1 > Text4(0) > 0 Then
Call fLv(7, Lv1)
Text3.Text = Text3.Text + "[" & Text1 & "]" & "垂死挣扎，提升属性值" & f & "点"
Call sx(f, 0, 0)
ElseIf hpc2 > Text5(0) > 0 Then
Call fLv(7, Lv2)
Text3.Text = Text3.Text + "[" & Text2 & "]" & "垂死挣扎，提升属性值" & f & "点"
Call sx(f, 0, 1)

Else
If Tur = 1 Then '战斗循环
If R2 = True Then
Round2 = Round2 - 1
    If Round2 = 1 Then
    Text3.Text = Text3.Text + "[" & Text2 & "]" & "八门遁甲使用时间过长，功力枯竭身亡！"
    Label19(1) = Val(Label19(1)) + 1                                   ' 自爆身亡算失败
    Label21(3) = Val(Label21(3)) + 1
    Call Lv
    Call win
    Timer1.Enabled = False
    Command9.Enabled = True
    Exit Sub
End If
End If
Call Skill(0, 0, Tur)
If SP(1) < 1 And Flag >= 60 And Flag < 65 Then
Text3.Text = Text3.Text + d0 & e & vbCrLf
Round2 = 3
R2 = True
SP(1) = SP(1) + 1
ElseIf SP(2) < 3 And Flag >= 93 And Flag < 100 Then
Text3.Text = Text3.Text + d & e & vbCrLf
SP(2) = SP(2) + 1
Else
Text3.Text = Text3.Text + d & e & vbCrLf
End If
Tur = 2
Exit Sub

ElseIf Tur = 2 Then
If R1 = True Then
Round1 = Round1 - 1
    If Round1 = 1 Then
    Text3.Text = Text3.Text + "[" & Text1 & "]" & "八门遁甲使用时间过长，功力枯竭身亡！"
    Timer1.Enabled = False
    Command9.Enabled = True
    Label19(3).Caption = Val(Label19(3)) + 1
    Label21(1) = Val(Label21(1)) + 1
    Call Lv02
    Call fail
    Exit Sub
End If
End If
Call Skill(0, 0, Tur)
If SP(1) < 1 And Flag >= 60 And Flag < 65 Then
Text3.Text = Text3.Text + c0 & e & vbCrLf
Round1 = 3
R1 = True
SP(1) = SP(1) + 1
ElseIf SP(2) < 3 And Flag >= 93 And Flag < 100 Then
Text3.Text = Text3.Text + c & e & vbCrLf
SP(2) = SP(2) + 1
Else
Text3.Text = Text3.Text + c & e & vbCrLf
End If
Tur = 1
Exit Sub
End If
End If
End Sub

Private Sub Form_Load()
Dim I
Dim strsql1 As String
Dim strsql2 As String

MsgBox "欢迎游戏姓名大乐斗擂台战(PVP)~", , "温馨提示"
Randomize

Label19(0).Caption = ID
Set db = New ADODB.Connection
Set Rs = New ADODB.Recordset
db.ConnectionString = "Provider=SQLOLEDB.1;Password=1123581321;Persist Security Info=True;User ID=hds1010886;Initial Catalog=hds1010886_db;Data Source=hds-101.hichina.com"
db.Open
If db.State = adStateOpen Then
'MsgBox "成功"
Else
MsgBox "连接失败"
End If
'         strsql = "select * from zc"                                    '打开表格
'         Rs.Open strsql, db, 3, 3
Set Rs = New ADODB.Recordset
strsql1 = " select * from zc where 账号=" & ID
Rs.Open strsql1, db, adOpenStatic, adLockReadOnly
Label19(1).Caption = Rs!胜场
Label19(2).Caption = Rs!等级
Label19(3).Caption = Rs!失败
Label19(4).Caption = Rs!金钱
fhy = Rs!复活药
Money = Label19(4).Caption

Set Rs = New ADODB.Recordset
strsql2 = " select * from zc where 擂主=" & "2"
Rs.Open strsql2, db, adOpenStatic, adLockReadOnly
Player2 = Rs!用户名
Label21(0).Caption = Rs!账号
Label21(1).Caption = Rs!胜场
Label21(2).Caption = Rs!等级
Label21(3).Caption = Rs!失败
Label21(4).Caption = Rs!金钱

If Label19(1) < 3 Then Lv1 = 1
If Label19(1) >= 3 And Label19(1) < 10 Then Lv1 = 2
If Label19(1) >= 10 And Label19(1) < 20 Then Lv1 = 3
If Label19(1) >= 20 Then Lv1 = 4
Label19(2) = Lv1
Money = Val(Label19(4))

If Label21(1) < 3 Then Lv2 = 1
If Label21(1) >= 3 And Label21(1) < 10 Then Lv2 = 2
If Label21(1) >= 10 And Label21(1) < 20 Then Lv2 = 3
If Label21(1) >= 20 Then Lv2 = 4
Label21(2) = Lv2
Money2 = Val(Label21(4))

Text1.Text = Trim(Player1)
Text2.Text = Trim(Player2)
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4(0).Locked = True
Text4(1).Locked = True
Text4(2).Locked = True
Text4(3).Locked = True
Text4(4).Locked = True
Text4(5).Locked = True
Text5(0).Locked = True
Text5(1).Locked = True
Text5(2).Locked = True
Text5(3).Locked = True
Text5(4).Locked = True
Text5(5).Locked = True
a = Int(Rnd * 10)
b = Int(Rnd * 10)
Flag = Int(Rnd * 100)

Text3.ForeColor = vbBlue
For I = 0 To 5
Text4(I).ForeColor = vbRed
Text5(I).ForeColor = vbRed
Next I
End Sub

Function Skill(Fis1 As Integer, Fis2 As Integer, who As Integer) '攻击方式
If SP(2) < 3 And Flag >= 93 And Flag < 100 Then
e = "使用吸星大法，"
Select Case who
    Case 2
    Call fLv(4, Lv1)
    e = e + "[" & Text2 & "]" & "功力被吸走一成"
    Text4(0).Text = Text4(0) + Val(Text5(0).Text) * f
    Text4(1).Text = Text4(1) + Val(Text5(1).Text) * f
    Text4(2).Text = Text4(2) + Val(Text5(2).Text) * f
    Text4(3).Text = Text4(3) + Val(Text5(3).Text) * f
    Text4(4).Text = Text4(4) + Val(Text5(4).Text) * f
    Text5(0).Text = Val(Text5(0).Text) * 0.9
    Text5(1).Text = Text5(1) - 1
    Text5(2).Text = Text5(2) - 1
    Text5(3).Text = Text5(3) - 1
    Text5(4).Text = Text5(4) - 1

    Case 1
    Call fLv(4, Lv2)
    e = e + "[" & Text1 & "]" & "功力被吸走一成"
    Text5(0).Text = Text5(0) + Val(Text4(0).Text) * f
    Text5(1).Text = Text5(1) + Val(Text4(1).Text) * f
    Text5(2).Text = Text5(2) + Val(Text4(2).Text) * f
    Text5(3).Text = Text5(3) + Val(Text4(3).Text) * f
    Text5(4).Text = Text5(4) + Val(Text4(4).Text) * f
    Text4(0).Text = Val(Text4(0).Text) * 0.9
    Text4(1).Text = Text4(1) - 1
    Text4(2).Text = Text4(2) - 1
    Text4(3).Text = Text4(3) - 1
    Text4(4).Text = Text4(4) - 1
End Select
Exit Function

ElseIf SP(1) < 1 And Flag >= 60 And Flag < 65 Then
e = "使用八门遁甲，"
Select Case who
    Case 2
    Call fLv(2, Lv1)
    e = e + "[" & Text1 & "]" & "的所有属性值上升" & f & "点"
    Call sx(f, 0, 0)

    Case 1
    Call fLv(2, Lv2)
    e = e + "[" & Text2 & "]" & "的所有属性值上升" & f & "点"
    Call sx(f, 0, 1)
End Select
Exit Function


ElseIf Flag >= 75 And Flag < 80 Then
e = "使用九阳神功，"
Select Case who
    Case 2
    Call fLv(1, Lv1)
    e = e + "[" & Text2 & "]" & "的所有属性值下降" & f & "点"
    Call sx(f, 1, 0)

    Case 1
    Call fLv(1, Lv2)
    e = e + "[" & Text1 & "]" & "的所有属性值下降" & f & "点"
    Call sx(f, 1, 1)
End Select
Exit Function

ElseIf Flag >= 80 And Flag < 90 Then
e = "发起攻击，"
Select Case who
    Case 2
    Call fLv(3, Lv1)
    e = e + "[" & Text2 & "]" & "被打晕了，" & "[" & Text1 & "]" & "趁机恢复体力" & f & "点"
    Text4(0).Text = Val(Text4(0).Text + f)
    
    Case 1
    Call fLv(3, Lv2)
    e = e + "[" & Text1 & "]" & "被打晕了，" & "[" & Text2 & "]" & "趁机恢复体力" & f & "点"
    Text5(0).Text = Val(Text5(0).Text + f)
End Select
Exit Function

Else
e = "发起攻击，"
Select Case who
    Case 2
    Call fLv(60, Lv1)
    If gj(0) > fy(1) Then
    e = e + "[" & Text2 & "]" & "受到" & Val(Abs(gj(0) - fy(1)) + k) & "点攻击"
    Text5(0).Text = Val(Text5(0).Text - Val(Abs(gj(0) - fy(1))) - k)
    Exit Function
    Else
    e = e + "[" & Text2 & "]" & "受到" & Val(k) & "点攻击"
    Text5(0).Text = Val(Text5(0).Text - k)
    Exit Function
    End If

    Case 1
    Call fLv(60, Lv2)
    If gj(1) > fy(0) Then
    e = e + "[" & Text1 & "]" & "受到" & Val(Abs(gj(1) - fy(0)) + k) & "点攻击"
    Text4(0).Text = Val(Text4(0).Text - Val(Abs(gj(1) - fy(0))) - k)
    Exit Function
    Else
    e = e + "[" & Text1 & "]" & "受到" & Val(k) & "点攻击"
    Text4(0).Text = Val(Text4(0).Text - k)
    Exit Function
    End If
End Select
End If
End Function

Function fLv(K1 As Integer, Lv As Integer) As Integer
If K1 = 60 Then                                                  '普通攻击
If Lv = 1 Then k = Int(Rnd * 30)
If Lv = 2 Then k = Int(Rnd * 25 + 5)
If Lv >= 3 Then k = Int(Rnd * 30 + 5)

ElseIf K1 = 1 Then                                            '九阳神功
If Lv1 = 1 Then f = Val(Int(Rnd * 15 + 1))
If Lv1 = 2 Then f = Val(Int(Rnd * 15 + 2))
If Lv1 >= 3 Then f = Val(Int(Rnd * 17 + 3))

ElseIf K1 = 2 Then                                            '八门遁甲
If Lv1 = 1 Then f = Val(Int(Rnd * 101 + 300))
If Lv1 = 2 Then f = Val(Int(Rnd * 101 + 320))
If Lv1 >= 3 Then f = Val(Int(Rnd * 101 + 340))

ElseIf K1 = 3 Then                                            '快速治疗
If Lv1 = 1 Then f = Val(Int(Rnd * 30))
If Lv1 = 2 Then f = Val(Int(Rnd * 30 + 10))
If Lv1 >= 3 Then f = Val(Int(Rnd * 30 + 20))

ElseIf K1 = 4 Then                                            '吸星大法
If Lv1 = 1 Then f = 0.1
If Lv1 = 2 Then f = 0.11
If Lv1 >= 3 Then f = 0.12

ElseIf K1 = 5 Then                                        '金蝉脱壳触发几率
If Lv = 1 Then Jc = 95
If Lv = 2 Then Jc = 94
If Lv >= 3 Then Jc = 93

ElseIf K1 = 6 Then                                       '金蝉脱壳回复血量
If Lv = 1 Then hp0 = 20
If Lv = 2 Then hp0 = 30
If Lv >= 3 Then hp0 = 40

ElseIf K1 = 7 Then                                            '回光返照提升属性值
If Lv = 1 Then f = Val(Int(Rnd * 10))
If Lv = 2 Then f = Val(Int(Rnd * 13 + 2))
If Lv >= 3 Then f = Val(Int(Rnd * 15 + 5))

ElseIf K1 = 8 Then                                            '回光返照触发血量
If Lv = 1 Then hpc1 = 15
If Lv = 2 Then hpc2 = 17
If Lv >= 3 Then hpc2 = 20

ElseIf K1 = 9 Then                                            '回光返照触发血量
If Lv = 1 Then hpc1 = 15
If Lv = 2 Then hpc2 = 17
If Lv >= 3 Then hpc2 = 20
End If
End Function

Function sx(zj As Integer, jj As Integer, kk As Integer) As Integer
hp(0) = Val(Text4(0).Text)  '生命值
Text4(0).Text = hp(0)
gj(0) = Val(Text4(1).Text) '攻击值
fy(0) = Val(Text4(2).Text)  '防御值
sd(0) = Val(Text4(3).Text) '速度值
mz(0) = Val(Text4(4).Text) '命中值

hp(1) = Val(Text5(0).Text) '生命值
Text5(0).Text = hp(1)
gj(1) = Val(Text5(1).Text) '攻击值
fy(0) = Val(Text4(2).Text) '防御值
sd(0) = Val(Text4(3).Text) '速度值
mz(0) = Val(Text4(4).Text) '命中值


yq(0) = Val(Text4(5).Text) '运气值
Text4(5).Text = yq(0)
yq(1) = Val(Text5(5).Text)
Text5(5).Text = yq(1)

If jj = 0 Then
If kk = 0 Then
Text4(1).Text = gj(0) + zj
gj(0) = Val(Text4(1).Text) '攻击值
Text4(2).Text = fy(0) + zj
fy(0) = Val(Text4(2).Text) '防御值
Text4(3).Text = sd(0) + zj
sd(0) = Val(Text4(3).Text) '速度值
Text4(4).Text = mz(0) + zj
mz(0) = Val(Text4(4).Text) '命中值


ElseIf kk = 1 Then
Text5(1).Text = gj(1) + zj
gj(1) = Val(Text5(1).Text) '攻击值
Text5(2).Text = fy(1) + zj
fy(1) = Val(Text5(2).Text) '防御值
Text5(3).Text = sd(1) + zj
sd(1) = Val(Text5(3).Text) '速度值
Text5(4).Text = mz(1) + zj
mz(1) = Val(Text5(4).Text) '命中值
End If

ElseIf jj = 1 Then
If kk = 1 Then
Text4(1).Text = gj(0) - zj
gj(0) = Val(Text4(1).Text) '攻击值
Text4(2).Text = fy(0) - zj
fy(0) = Val(Text4(2).Text) '防御值
Text4(3).Text = sd(0) - zj
sd(0) = Val(Text4(3).Text) '速度值
Text4(4).Text = mz(0) - zj
mz(0) = Val(Text4(4).Text) '命中值

ElseIf kk = 0 Then
Text5(1).Text = gj(1) - zj
gj(1) = Val(Text5(1).Text) '攻击值
Text5(2).Text = fy(1) - zj
fy(1) = Val(Text5(2).Text) '防御值
Text5(3).Text = sd(1) - zj
sd(1) = Val(Text5(3).Text) '速度值
Text5(4).Text = mz(1) - zj
mz(1) = Val(Text5(4).Text) '命中值
End If
End If

End Function

Private Function Name_Do(str As Integer)
Dim s, n, y, t, k, I
t = 0
n = 0
y = 0
k = 0
For I = 1 To str
s = Mid(Trim(Text1 & Text2), I, 1)
If s = " " Then
k = k + 1
Else
If s = "0" Or Val(s) > 0 Then
n = n + 1
Else
If Asc(s) >= 65 And Asc(s) <= 90 Or Asc(s) >= 97 And Asc(s) <= 122 Then
y = y + 1
ElseIf Asc(s) > 65 Then
t = t + 1
End If
End If
End If
Next I
Name_Do = k + y + n
End Function

Public Sub Lv()
If Label19(1) < 3 Then Lv1 = 1
If Label19(1) >= 3 And Label19(1) < 10 Then Lv1 = 2
If Label19(1) >= 10 And Label19(1) < 20 Then Lv1 = 3
If Label19(1) >= 20 Then Lv1 = 4
Label19(2) = Lv1
Money = Money + Int(Rnd * 40 + 10)
Label19(4).Caption = Money
End Sub

Public Sub Lv02()
If Label21(1) < 3 Then Lv2 = 1
If Label21(1) >= 3 And Label21(1) < 10 Then Lv2 = 2
If Label21(1) >= 10 And Label21(1) < 20 Then Lv2 = 3
If Label21(1) >= 20 Then Lv2 = 4
Label21(2) = Lv2
Money2 = Money2 + Int(Rnd * 40 + 10)
Label21(4).Caption = Money2
End Sub

Public Sub win()
Dim strsql1, strsql2
MsgBox "恭喜您挑战成功，成为擂主！！奖励金币200，复活药*1", , "恭喜"
Money = Money + 200
Label19(4).Caption = Money
fhy = fhy + 1
'Set Rs = New ADODB.Recordset
'strsql1 = " select * from zc where 账号=" & ID
'Rs.Open strsql1, db, adOpenStatic, adLockReadOnly
'Rs!擂主 = 2
Dim a As String
Dim b As String
a = "update zc set 胜场= '" & Label19(1) & "',失败='" & Label19(3) & "',金钱='" & Label19(4) & "',复活药='" & fhy & "',擂主='" & "2" & "',等级=" & Val(Label19(2)) & " where 账号=" & Val(Label19(0))
Call CnSql(a, 2)
b = "select * from zc where 账号=" & Val(Label19(0))
Call CnSql(b, 1)


'Set Rs = New ADODB.Recordset
'strsql2 = " select * from zc where 账号=" & Val(Label21(0))
'Rs.Open strsql2, db, adOpenStatic, adLockReadOnly
'Rs!擂主 = 0
Dim a2 As String
Dim b2 As String
a2 = "update zc set 胜场= '" & Label21(1) & "',失败='" & Label21(3) & "',金钱='" & Label21(4) & "',擂主='" & "0" & "',等级=" & Val(Label21(2)) & " where 账号=" & Val(Label21(0))
Call CnSql(a2, 2)
b2 = "select * from zc where 账号=" & Val(Label21(0))
Call CnSql(b2, 1)
End Sub

Public Sub fail()
MsgBox "很遗憾您挑战失败，请尝试再次挑战！"
Dim a2 As String
Dim b2 As String
a2 = "update zc set 胜场= '" & Label21(1) & "',失败='" & Label21(3) & "',金钱='" & Label21(4) & "',等级=" & Val(Label21(2)) & " where 账号=" & Val(Label21(0))
Call CnSql(a2, 2)
b2 = "select * from zc where 账号=" & Val(Label21(0))
Call CnSql(b2, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer) '卸载窗体事件
Dim a As String
Dim b As String
a = "update zc set 胜场= '" & Label19(1) & "',失败='" & Label19(3) & "',金钱='" & Label19(4) & "',等级=" & Val(Label19(2)) & " where 账号=" & Val(Label19(0))
Call CnSql(a, 2)
b = "select * from zc where 账号=" & Val(Label19(0))
Call CnSql(b, 1)

Dim a2 As String
Dim b2 As String
a2 = "update zc set 胜场= '" & Label21(1) & "',失败='" & Label21(3) & "',金钱='" & Label21(4) & "',等级=" & Val(Label21(2)) & " where 账号=" & Val(Label21(0))
Call CnSql(a2, 2)
b2 = "select * from zc where 账号=" & Val(Label21(0))
Call CnSql(b2, 1)

End Sub

