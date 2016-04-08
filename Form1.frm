VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   12990
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Height          =   180
      Left            =   12000
      TabIndex        =   30
      Top             =   6960
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "再次战斗"
      Height          =   615
      Left            =   1920
      TabIndex        =   29
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "暂停"
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "快"
      Height          =   375
      Left            =   7200
      TabIndex        =   27
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "中"
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "慢"
      Height          =   375
      Left            =   5520
      TabIndex        =   25
      Top             =   5760
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4080
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出游戏"
      Height          =   615
      Left            =   360
      TabIndex        =   23
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始战斗"
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   5
      Left            =   2160
      TabIndex        =   21
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   20
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   19
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   4935
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   360
      Width           =   6855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "游戏速度："
      Height          =   495
      Left            =   4320
      TabIndex        =   24
      Top             =   5760
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   360
      Y2              =   8040
   End
   Begin VB.Label Label7 
      Caption         =   "HP值"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "攻击"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "速度"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "命中"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "运气"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "防御"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer '左边随即优先攻击
Dim b As Integer '右边随即优先攻击
Dim k As Integer '攻击值数目

Dim c As String '左边向右边先攻击
Dim d As String '右边向左边先攻击

Dim e As String '攻击方式
Dim f As Integer '下降属性值

Dim hp(1) As Integer '生命值
Dim gj(1) As Integer '攻击值
Dim fy(1) As Integer '防御值
Dim sd(1) As Integer '速度值
Dim mz(1) As Integer '命中值
Dim yq(1) As Integer '运气值
Dim Tur As Integer '静态变量


Private Sub Command1_Click()
Dim lngReturn As Long

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
For i = 1 To Len(Trim(Text1))
lngReturn = CLng("&h" & Hex((AscW(Mid(Text1, i, 1)))))
If i = 1 Then
Text4(0).Text = Mid(lngReturn, 1, 3)
Text4(1).Text = Val(Mid(lngReturn, 3, 2) + 30)
Text4(4).Text = Val(Mid(lngReturn, 2, 2) + 50)
End If
If i = 2 Then
Text4(2).Text = Val(Mid(lngReturn, 1, 2) + 30)
Text4(3).Text = Val(Mid(lngReturn, 2, 2) + 40)
End If
Next i

For i = 1 To Len(Trim(Text2))
lngReturn = CLng("&h" & Hex((AscW(Mid(Text2, i, 1)))))
If i = 1 Then
Text5(0).Text = Mid(lngReturn, 1, 3)
Text5(1).Text = Val(Mid(lngReturn, 3, 2) + 30)
Text5(4).Text = Val(Mid(lngReturn, 2, 2) + 50)
End If
If i = 2 Then
Text5(2).Text = Val(Mid(lngReturn, 1, 2) + 30)
Text5(3).Text = Val(Mid(lngReturn, 2, 2) + 40)
End If
Next i
Text4(5).Text = Int(Rnd * 100)
Text5(5).Text = Int(Rnd * 100)


Call sx(0, 0, 0)

Text3.Text = "姓名大作战 VB版" & Label11 & vbCrLf & vbCrLf
Text3.Text = Text3.Text + Text1 & "  " & "HP：" & Text4(0) & "  " & "攻：" & Text4(1) & "  " & "防：" & Text4(2) & "  " & "速：" & Text4(3) & "  " & "技：" & Text4(4) & "  " & "运：" & Text4(5) & vbCrLf
Text3.Text = Text3.Text + Text2 & "  " & "HP：" & Text5(0) & "  " & "攻：" & Text5(1) & "  " & "防：" & Text5(2) & "  " & "速：" & Text5(3) & "  " & "技：" & Text5(4) & "  " & "运：" & Text5(5) & vbCrLf & vbCrLf

If Text4(5).Text > Text5(5).Text Then '战斗先机
Tur = 2
ElseIf Text4(5).Text < Text5(5).Text Then
Tur = 1
Else
MsgBox "运气都一样，在来一次！", , "提示"
Text4(5).Text = Int(Rnd * 100)
Text5(5).Text = Int(Rnd * 100)
End If
Timer1.Enabled = True

End If


End Sub


Private Sub Command3_Click()          '减慢
Timer1.Interval = 2000

End Sub

Private Sub Command4_Click()          '正常
Timer1.Interval = 1500
End Sub

Private Sub Command5_Click()          '加快
Timer1.Interval = 1000
End Sub

Private Sub Command6_Click()          '暂停
Timer1.Interval = 0
End Sub

Private Sub Command7_Click()                   '复位键
MsgBox "请重新输入两人姓名", , "提示"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub Command8_Click()          '作弊开启
If Text1.Text = "邓国雄" Then
Text4(0).Text = 500
Text4(5).Text = 100
End If

End Sub

Private Sub Timer1_Timer()
a = Int(Rnd * 20)
b = Int(Rnd * 20)
If Text4(0) < 0 Then
Text4(0) = 0
Text3.Text = Text3.Text + "[" & Text1 & "]" & "被打败！"
Timer1.Enabled = False
Exit Sub
ElseIf Text5(0) < 0 Then
Text5(0) = 0
Text3.Text = Text3.Text + "[" & Text2 & "]" & "被打败！"
Timer1.Enabled = False
Exit Sub
ElseIf 10 > Text4(0) > 0 Then
f = Val(Int(Rnd * 10))
Text3.Text = Text3.Text + "[" & Text1 & "]" & "垂死挣扎，提升属性值" & f & "点"
Call sx(f, 0, 0)
ElseIf 10 > Text5(0) > 0 Then
f = Val(Int(Rnd * 10))
Text3.Text = Text3.Text + "[" & Text2 & "]" & "垂死挣扎，提升属性值" & f & "点"
Call sx(f, 0, 1)
Else
If Tur = 1 Then '战斗循环
Call Skill(0, 0, Tur)
Text3.Text = Text3.Text + d & e & vbCrLf
Tur = 2
Exit Sub
ElseIf Tur = 2 Then
Call Skill(0, 0, Tur)
Text3.Text = Text3.Text + c & e & vbCrLf
Tur = 1
Exit Sub
End If
End If
End Sub

Private Sub Form_Load()
a = Int(Rnd * 10)
b = Int(Rnd * 10)

Text3.ForeColor = vbBlue
For i = 0 To 5
Text4(i).ForeColor = vbRed
Text5(i).ForeColor = vbRed
Next i
End Sub

Function Skill(Fis1 As Integer, Fis2 As Integer, who As Integer) '攻击方式

If Val(Fis1 + a) > Val(Fis2 + b) Then
e = "发起攻击，"
     k = Val(Int(Rnd * 10))
Select Case who
       Case 2
     If gj(0) > fy(1) Then
     e = e + "[" & Text2 & "]" & "受到" & Val(Abs(gj(0) - fy(1)) + k) & "点攻击"
     Text5(0).Text = Val(Text5(0).Text - Val(Abs(gj(0) - fy(1))) + k)
     Exit Function
     Else
     e = e + "[" & Text2 & "]" & "受到" & Val(k) & "点攻击"
     Text5(0).Text = Val(Text5(0).Text - k)
     Exit Function
     End If

       Case 1
     If gj(1) > fy(0) Then
     e = e + "[" & Text1 & "]" & "受到" & Val(Abs(gj(1) - fy(0)) + k) & "点攻击"
     Text4(0).Text = Val(Text4(0).Text - Val(Abs(gj(1) - fy(0))) + k)
     Exit Function
     Else
     e = e + "[" & Text1 & "]" & "受到" & Val(k) & "点攻击"
     Text4(0).Text = Val(Text4(0).Text - k)
     Exit Function
     End If

End Select

ElseIf Val(Fis1 + a) < Val(Fis2 + b) Then
e = "使用降龙十八掌，"
Select Case who
       Case 2
     e = e + "[" & Text2 & "]" & "功力下降两成"
     Text5(0).Text = Val(Text5(0).Text) / 2

       Case 1
     e = e + "[" & Text1 & "]" & "功力下降三成"
     Text4(0).Text = Val(Text4(0).Text) / 2

End Select
Exit Function

Else
a = Int(Rnd * 20)
b = Int(Rnd * 20)
If a > b Then
e = "使用九阳神功，"
     f = Val(Int(Rnd * 10))
Select Case who
       Case 2

     e = e + "[" & Text2 & "]" & "的所有属性值下降" & f & "点"
     Call sx(f, 1, 0)

       Case 1

     e = e + "[" & Text1 & "]" & "的所有属性值下降" & f & "点"
     Call sx(f, 1, 1)

End Select
Exit Function

ElseIf b > a Then
e = "发起攻击，"
     f = Val(Int(Rnd * 20))
Select Case who
       Case 2

     e = e + "[" & Text2 & "]" & "被打晕了，" & "[" & Text1 & "]" & "趁机恢复体力" & f & "点"
     Text4(0).Text = Val(Text4(0).Text + f)

       Case 1

     e = e + "[" & Text1 & "]" & "被打晕了，" & "[" & Text2 & "]" & "趁机恢复体力" & f & "点"
     Text5(0).Text = Val(Text5(0).Text + f)

End Select
Exit Function
End If
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
Dim s, n, y, t, k
t = 0
n = 0
y = 0
k = 0
For i = 1 To str
s = Mid(Trim(Text1 & Text2), i, 1)
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
Next i
Name_Do = k + y + n
End Function
Private Sub Form_Unload(Cancel As Integer) '卸载窗体事件
MsgBox "谢谢游戏！欢迎继续游玩！", , "提示"
End Sub

