VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ֶ� v1.5  ����֮��"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "��Ʒ��"
      Height          =   1575
      Left            =   8400
      TabIndex        =   55
      Top             =   6000
      Width           =   2655
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ʹ��"
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   66
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "0"
         DataSource      =   "Adodc1"
         Height          =   180
         Index           =   7
         Left            =   1800
         TabIndex        =   65
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   6
         Left            =   960
         TabIndex        =   64
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "����ҩ"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ʹ��"
         Height          =   240
         Index           =   0
         Left            =   2160
         TabIndex        =   62
         Top             =   720
         Width           =   420
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "0"
         DataSource      =   "Adodc1"
         Height          =   180
         Index           =   6
         Left            =   1800
         TabIndex        =   61
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "HP+50"
         Height          =   180
         Index           =   4
         Left            =   960
         TabIndex        =   60
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "С��ҩ"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2760
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   1680
         TabIndex        =   58
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Ч��"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   57
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "�̵�"
      Height          =   375
      Left            =   1800
      TabIndex        =   51
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
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
      Caption         =   "����"
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   " �� "
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   26
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��"
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
      Left            =   11880
      Top             =   6120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
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
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "��"
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
      Index           =   6
      Left            =   6000
      TabIndex        =   54
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      DataSource      =   "Adodc1"
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
      Index           =   5
      Left            =   5280
      TabIndex        =   53
      Top             =   240
      Width           =   165
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "��"
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
      Index           =   5
      Left            =   4560
      TabIndex        =   52
      Top             =   240
      Width           =   315
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
      Caption         =   "��Ǯ��"
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
      Caption         =   "ʧ�ܣ�"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   47
      Top             =   6480
      Width           =   540
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Lv��"
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
      Caption         =   "ID��"
      Height          =   180
      Index           =   1
      Left            =   1800
      TabIndex        =   43
      Top             =   360
      Width           =   360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "ʤ����"
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
      Caption         =   "��Ϸѡ�"
      Height          =   180
      Left            =   3360
      TabIndex        =   39
      Top             =   7080
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   38
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   37
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   10200
      TabIndex        =   36
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   10200
      TabIndex        =   35
      Top             =   4320
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "�ٶ�"
      Height          =   180
      Left            =   10200
      TabIndex        =   34
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   10200
      TabIndex        =   33
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   10200
      TabIndex        =   32
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "HPֵ"
      Height          =   180
      Left            =   10200
      TabIndex        =   31
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   10200
      TabIndex        =   30
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "��Ϸ�ٶȣ�"
      Height          =   180
      Left            =   3360
      TabIndex        =   24
      Top             =   6360
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "HPֵ"
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "�ٶ�"
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1200
      TabIndex        =   3
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1200
      TabIndex        =   2
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   360
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim a As Integer '����漴���ȹ���
Dim b As Integer '�ұ��漴���ȹ���
Dim k As Integer '����ֵ��Ŀ

Dim c As String '������ұ��ȹ���
Dim d As String '�ұ�������ȹ���
Dim c0 As String
Dim d0 As String

Dim e As String '������ʽ
Dim f As Integer '�½�����ֵ

Dim hp(1) As Integer '����ֵ
Dim gj(1) As Integer '����ֵ
Dim fy(1) As Integer '����ֵ
Dim sd(1) As Integer '�ٶ�ֵ
Dim mz(1) As Integer '����ֵ
Dim yq(1) As Integer '����ֵ
Dim Tur As Integer '��̬����
Dim Flag As Integer    '������ʽ��������
Dim Die
Dim hp0, Jc                 '����ѿ�
Dim hpc1, hpc2           '�عⷵ��
Dim slta
Dim Round1, Round2  'ʹ�ð��Ŷݼ���������ʱ
Dim R1 As Boolean, R2 As Boolean   '������������ʱ
Dim SP(0 To 60)       '������ʹ������
Dim Connstring

Private Sub Command1_Click()
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
MsgBox "�����뺺�֣�", , "��ʾ"
Exit Sub
Else
If Text1 = "" And Text2 = "" Then
MsgBox "���������֣�", , "��ʾ"
Exit Sub
End If
c = "[" & Text1 & "]" & "��" & "[" & Text2 & "]"
d = "[" & Text2 & "]" & "��" & "[" & Text1 & "]"
c0 = "[" & Text1 & "]" & "��" & "�Լ�"
d0 = "[" & Text2 & "]" & "��" & "�Լ�"
For I = 1 To Len(Trim(Text1))
lngReturn = CLng("&h" & Hex((AscW(Mid(Text1, I, 1)))))
If I = 1 Then
'Text4(0).Text = Mid(lngReturn, 1, 3)
Text4(1).Text = Val(Mid(lngReturn, 3, 2) + 30)
Text4(4).Text = Val(Mid(lngReturn, 2, 2) + 50)
End If
If I = 2 Then
Text4(2).Text = Val(Mid(lngReturn, 1, 2) + 30)
Text4(3).Text = Val(Mid(lngReturn, 2, 2) + 40)
End If
Next I

'For I = 1 To Len(Trim(Text2))
'lngReturn = CLng("&h" & Hex((AscW(Mid(Text2, I, 1)))))
'If I = 1 Then
'Text5(0).Text = Mid(lngReturn, 1, 3)
'Text5(1).Text = Val(Mid(lngReturn, 3, 2) + 30)
'Text5(4).Text = Val(Mid(lngReturn, 2, 2) + 50)
'End If
'If I = 2 Then
'Text5(2).Text = Val(Mid(lngReturn, 1, 2) + 30)
'Text5(3).Text = Val(Mid(lngReturn, 2, 2) + 40)
'End If
'Next I
Text4(5).Text = Int(Rnd * 100)
Text5(5).Text = Int(Rnd * 100)

Call sx(0, 0, 0)

Text3.Text = "��������ս VB��" & vbCrLf & vbCrLf
Text3.Text = Text3.Text + Text1 & "  " & "HP��" & Text4(0) & "  " & "����" & Text4(1) & "  " & "����" & Text4(2) & "  " & "�٣�" & Text4(3) & "  " & "����" & Text4(4) & "  " & "�ˣ�" & Text4(5) & vbCrLf
Text3.Text = Text3.Text + Text2 & "  " & "HP��" & Text5(0) & "  " & "����" & Text5(1) & "  " & "����" & Text5(2) & "  " & "�٣�" & Text5(3) & "  " & "����" & Text5(4) & "  " & "�ˣ�" & Text5(5) & vbCrLf & vbCrLf

If (Text4(5).Text + Text4(3).Text) > (Text5(5).Text + Text5(3).Text) Then 'ս���Ȼ�
Tur = 2
ElseIf (Text4(5).Text + Text4(3).Text) < (Text5(5).Text + Text5(3).Text) Then
Tur = 1
Else
MsgBox "�޷�ȷ�����֣�������һ�Σ�", , "��ʾ"
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

Private Sub Command3_Click()          '����
Timer1.Interval = Timer1.Interval - 500
End Sub

Private Sub Command4_Click()          '������
Form6.Show
End Sub

Private Sub Command5_Click()          '�ӿ�
Timer1.Interval = Timer1.Interval + 500
End Sub

Private Sub Command6_Click()          '��ͣ
Timer1.Interval = 0
End Sub

Private Sub Command7_Click()
Timer1.Enabled = False
MsgBox "����������ս��������һ�㣡"
Label19(5) = 1
Call Save
Call Form_Load
End Sub

Private Sub Command8_Click()          '���׿���
If Text1.Text = "�˹���" Then
Text4(0).Text = Text4(0).Text + 500
ElseIf Text2.Text = "�˹���" Then
Text5(0).Text = Text5(0).Text + 500
Else
MsgBox "����Ȩʹ����������", , "����"
End If
End Sub

Private Sub Command9_Click()
Unload Me
Form10.Show
End Sub

Private Sub Label22_Click(Index As Integer)
    Select Case Index
        Case 0: Cmd0
        'Case 1: Cmd1
    End Select
End Sub

Private Sub Cmd0()
If Label19(6) > 0 Then
Text4(0).Text = Val(Text4(0)) + 50
Label19(6) = Label19(6) - 1
Text3.Text = Text3.Text + vbCrLf + "[" + Text1.Text + "] ʹ����С��ҩ��HP+50��" + vbCrLf
Else
MsgBox "���С��ҩû��ʣ�࣡"
End If
End Sub
Private Sub Cmd1()
If Label19(7) > 0 Then
Dim lngReturn As Long
Dim I

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

Label19(7) = Label19(7) - 1
Text3.Text = Text3.Text + vbCrLf + "[" + Text1.Text + "] ʹ���˸���ҩ��ԭ����Ѫ���" + vbCrLf
Timer1.Enabled = True
Else
MsgBox "��ĸ���ҩû��ʣ�࣡"
End If
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
MsgBox "����������Ϊ���֣�"
Text2.Text = ""
End If
If Sum = 1 Then
MsgBox "����������Ϊ�����֣����������룡"
Text2.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
a = Int(Rnd * 20)
b = Int(Rnd * 20)
Flag = Int(Rnd * 100)
Die = Int(Rnd * 100) '����ѿ��ͷż���
Call fLv(8, Lv1)
Call fLv(9, Lv2)
If Text4(0) <= 0 Then
Call fLv(5, Lv1)
Call fLv(6, Lv1)
If Die >= Jc And SP(3) < 1 Then
Text4(0) = hp0
Text3.Text = Text3.Text + "[" & Text1 & "]" & "ʹ�ý���ѿǣ�" + vbCrLf
SP(3) = SP(3) + 1
Else
Text4(0) = 0
Text3.Text = Text3.Text + "[" & Text1 & "]" & "����ܣ�"
Timer1.Enabled = False
Command9.Enabled = True
Label19(3).Caption = Val(Label19(3)) + 1 '����ʧ��
Call fail
End If
Exit Sub

ElseIf Text5(0) <= 0 Then
Call fLv(5, Lv2)
Call fLv(6, Lv2)
If Die >= Jc And SP(4) < 1 Then
Text5(0) = hp0
Text3.Text = Text3.Text + "[" & Text2 & "]" & "ʹ�ý���ѿǣ�" + vbCrLf
SP(4) = SP(4) + 1
Else
Text5(0) = 0
Text3.Text = Text3.Text + "[" & Text2 & "]" & "����ܣ�"
Timer1.Enabled = False
Command9.Enabled = True
Label19(1).Caption = Val(Label19(1)) + 1 '����ʤ��
Call win
Call Lv
End If
Exit Sub

ElseIf hpc1 > Text4(0) > 0 Then
Call fLv(7, Lv1)
Text3.Text = Text3.Text + "[" & Text1 & "]" & "������������������ֵ" & f & "��"
Call sx(f, 0, 0)
ElseIf hpc2 > Text5(0) > 0 Then
Call fLv(7, Lv2)
Text3.Text = Text3.Text + "[" & Text2 & "]" & "������������������ֵ" & f & "��"
Call sx(f, 0, 1)

Else
If Tur = 1 Then 'ս��ѭ��
If R2 = True Then
Round2 = Round2 - 1
    If Round2 = 99 Then
    Text3.Text = Text3.Text + "[" & Text2 & "]" & "���Ŷݼ�ʹ��ʱ������������ݽ�������"
    Label19(1) = Label19(1) + 1                                   ' �Ա�������ʧ��
    Call win
    Call Lv
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
    If Round1 = 99 Then
    Text3.Text = Text3.Text + "[" & Text1 & "]" & "���Ŷݼ�ʹ��ʱ������������ݽ�������"
    Timer1.Enabled = False
    Command9.Enabled = True
    Label19(3).Caption = Val(Label19(3)) + 1
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
Dim J
Dim strsql
MsgBox "��ӭ��������֮��~", , "��ܰ��ʾ"
Randomize

Label19(0).Caption = ID
Set db = New ADODB.Connection
Set Rs = New ADODB.Recordset
db.ConnectionString = "Provider=SQLOLEDB.1;Password=1123581321;Persist Security Info=True;User ID=hds1010886;Initial Catalog=hds1010886_db;Data Source=hds-101.hichina.com"
db.Open
If db.State = adStateOpen Then
'MsgBox "�ɹ�"
Else
MsgBox "����ʧ��"
End If
Set Rs = New ADODB.Recordset
strsql = " select * from zc where �˺�=" & ID
Rs.Open strsql, db, adOpenStatic, adLockReadOnly
Label19(1).Caption = Rs!ʤ��
Label19(2).Caption = Rs!�ȼ�
Label19(3).Caption = Rs!ʧ��
Label19(4).Caption = Rs!��Ǯ
Label19(5).Caption = Rs!����֮��
Label19(6).Caption = Rs!С��ҩ
Label19(7).Caption = Rs!����ҩ
Text1.Text = Trim(Player1)
Dim lngReturn As Long
Dim I

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

Text2.Text = ""
Text5(0).Text = ""
Text5(1).Text = ""
Text5(2).Text = ""
Text5(3).Text = ""
Text5(4).Text = ""
If Label19(1) < 3 Then Lv1 = 1
If Label19(1) >= 3 And Label19(1) < 10 Then Lv1 = 2
If Label19(1) >= 10 And Label19(1) < 20 Then Lv1 = 3
If Label19(1) >= 20 Then Lv1 = 4
Label19(2) = Lv1
Lv2 = 1
slta = Label19(5)
If slta = 1 Then Call slta1
If slta = 2 Then Call slta2
If slta = 3 Then Call slta3
If slta = 4 Then Call slta4
If slta = 5 Then Call slta5
If slta = 6 Then Call slta6
If slta = 7 Then Call slta7
If slta = 8 Then Call slta8
If slta = 9 Then Call slta9
If slta = 10 Then Call slta10

Text1.Locked = True
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
For J = 0 To 5
Text4(J).ForeColor = vbRed
Text5(J).ForeColor = vbRed
Next J
End Sub

Function Skill(Fis1 As Integer, Fis2 As Integer, who As Integer) '������ʽ
If SP(2) < 3 And Flag >= 93 And Flag < 100 Then
e = "ʹ�����Ǵ󷨣�"
Select Case who
    Case 2
    Call fLv(4, Lv1)
    e = e + "[" & Text2 & "]" & "����������һ��"
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
    e = e + "[" & Text1 & "]" & "����������һ��"
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

'ElseIf SP(1) < 1 And Flag >= 60 And Flag < 65 Then
'e = "ʹ�ð��Ŷݼף�"
'Select Case who
'    Case 2
'    Call fLv(2, Lv1)
'    e = e + "[" & Text1 & "]" & "����������ֵ����" & f & "��"
'    Call sx(f, 0, 0)
'
'    Case 1
'    Call fLv(2, Lv2)
'    e = e + "[" & Text2 & "]" & "����������ֵ����" & f & "��"
'    Call sx(f, 0, 1)
'End Select
'Exit Function


ElseIf Flag >= 75 And Flag < 80 Then
e = "ʹ�þ����񹦣�"
Select Case who
    Case 2
    Call fLv(1, Lv1)
    e = e + "[" & Text2 & "]" & "����������ֵ�½�" & f & "��"
    Call sx(f, 1, 0)

    Case 1
    Call fLv(1, Lv2)
    e = e + "[" & Text1 & "]" & "����������ֵ�½�" & f & "��"
    Call sx(f, 1, 1)
End Select
Exit Function

ElseIf Flag >= 80 And Flag < 90 Then
e = "���𹥻���"
Select Case who
    Case 2
    Call fLv(3, Lv1)
    e = e + "[" & Text2 & "]" & "�������ˣ�" & "[" & Text1 & "]" & "�û��ָ�����" & f & "��"
    Text4(0).Text = Val(Text4(0).Text + f)
    
    Case 1
    Call fLv(3, Lv2)
    e = e + "[" & Text1 & "]" & "�������ˣ�" & "[" & Text2 & "]" & "�û��ָ�����" & f & "��"
    Text5(0).Text = Val(Text5(0).Text + f)
End Select
Exit Function

Else
e = "���𹥻���"
Select Case who
    Case 2
    Call fLv(60, Lv1)
    If gj(0) > fy(1) Then
    e = e + "[" & Text2 & "]" & "�ܵ�" & Val(Abs(gj(0) - fy(1)) + k) & "�㹥��"
    Text5(0).Text = Val(Text5(0).Text - Val(Abs(gj(0) - fy(1))) - k)
    Exit Function
    Else
    e = e + "[" & Text2 & "]" & "�ܵ�" & Val(k) & "�㹥��"
    Text5(0).Text = Val(Text5(0).Text - k)
    Exit Function
    End If

    Case 1
    Call fLv(60, Lv2)
    If gj(1) > fy(0) Then
    e = e + "[" & Text1 & "]" & "�ܵ�" & Val(Abs(gj(1) - fy(0)) + k) & "�㹥��"
    Text4(0).Text = Val(Text4(0).Text - Val(Abs(gj(1) - fy(0))) - k)
    Exit Function
    Else
    e = e + "[" & Text1 & "]" & "�ܵ�" & Val(k) & "�㹥��"
    Text4(0).Text = Val(Text4(0).Text - k)
    Exit Function
    End If
End Select
End If
End Function

Function fLv(K1 As Integer, Lv As Integer) As Integer
If K1 = 60 Then                                                  '��ͨ����
If Lv = 1 Then k = Int(Rnd * 30)
If Lv = 2 Then k = Int(Rnd * 25 + 5)
If Lv >= 3 Then k = Int(Rnd * 30 + 5)

ElseIf K1 = 1 Then                                            '������
If Lv1 = 1 Then f = Val(Int(Rnd * 15 + 1))
If Lv1 = 2 Then f = Val(Int(Rnd * 15 + 2))
If Lv1 >= 3 Then f = Val(Int(Rnd * 17 + 3))

ElseIf K1 = 2 Then                                            '���Ŷݼ�
If Lv1 = 1 Then f = Val(Int(Rnd * 101 + 300))
If Lv1 = 2 Then f = Val(Int(Rnd * 101 + 320))
If Lv1 >= 3 Then f = Val(Int(Rnd * 101 + 340))

ElseIf K1 = 3 Then                                            '��������
If Lv1 = 1 Then f = Val(Int(Rnd * 30))
If Lv1 = 2 Then f = Val(Int(Rnd * 30 + 10))
If Lv1 >= 3 Then f = Val(Int(Rnd * 30 + 20))

ElseIf K1 = 4 Then                                            '���Ǵ�
If Lv1 = 1 Then f = 0.1
If Lv1 = 2 Then f = 0.11
If Lv1 >= 3 Then f = 0.12

ElseIf K1 = 5 Then                                        '����ѿǴ�������
If Lv = 1 Then Jc = 95
If Lv = 2 Then Jc = 94
If Lv >= 3 Then Jc = 93

ElseIf K1 = 6 Then                                       '����ѿǻظ�Ѫ��
If Lv = 1 Then hp0 = 20
If Lv = 2 Then hp0 = 30
If Lv >= 3 Then hp0 = 40
End If
End Function

Function sx(zj As Integer, jj As Integer, kk As Integer) As Integer
hp(0) = Val(Text4(0).Text)  '����ֵ
Text4(0).Text = hp(0)
gj(0) = Val(Text4(1).Text) '����ֵ
fy(0) = Val(Text4(2).Text)  '����ֵ
sd(0) = Val(Text4(3).Text) '�ٶ�ֵ
mz(0) = Val(Text4(4).Text) '����ֵ

hp(1) = Val(Text5(0).Text) '����ֵ
Text5(0).Text = hp(1)
gj(1) = Val(Text5(1).Text) '����ֵ
fy(0) = Val(Text4(2).Text) '����ֵ
sd(0) = Val(Text4(3).Text) '�ٶ�ֵ
mz(0) = Val(Text4(4).Text) '����ֵ


yq(0) = Val(Text4(5).Text) '����ֵ
Text4(5).Text = yq(0)
yq(1) = Val(Text5(5).Text)
Text5(5).Text = yq(1)

If jj = 0 Then
If kk = 0 Then
Text4(1).Text = gj(0) + zj
gj(0) = Val(Text4(1).Text) '����ֵ
Text4(2).Text = fy(0) + zj
fy(0) = Val(Text4(2).Text) '����ֵ
Text4(3).Text = sd(0) + zj
sd(0) = Val(Text4(3).Text) '�ٶ�ֵ
Text4(4).Text = mz(0) + zj
mz(0) = Val(Text4(4).Text) '����ֵ


ElseIf kk = 1 Then
Text5(1).Text = gj(1) + zj
gj(1) = Val(Text5(1).Text) '����ֵ
Text5(2).Text = fy(1) + zj
fy(1) = Val(Text5(2).Text) '����ֵ
Text5(3).Text = sd(1) + zj
sd(1) = Val(Text5(3).Text) '�ٶ�ֵ
Text5(4).Text = mz(1) + zj
mz(1) = Val(Text5(4).Text) '����ֵ
End If

ElseIf jj = 1 Then
If kk = 1 Then
Text4(1).Text = gj(0) - zj
gj(0) = Val(Text4(1).Text) '����ֵ
Text4(2).Text = fy(0) - zj
fy(0) = Val(Text4(2).Text) '����ֵ
Text4(3).Text = sd(0) - zj
sd(0) = Val(Text4(3).Text) '�ٶ�ֵ
Text4(4).Text = mz(0) - zj
mz(0) = Val(Text4(4).Text) '����ֵ

ElseIf kk = 0 Then
Text5(1).Text = gj(1) - zj
gj(1) = Val(Text5(1).Text) '����ֵ
Text5(2).Text = fy(1) - zj
fy(1) = Val(Text5(2).Text) '����ֵ
Text5(3).Text = sd(1) - zj
sd(1) = Val(Text5(3).Text) '�ٶ�ֵ
Text5(4).Text = mz(1) - zj
mz(1) = Val(Text5(4).Text) '����ֵ
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
Public Sub win()
Label19(5).Caption = Val(Label19(5)) + 1
slta = Label19(5)
If slta <= 10 Then
MsgBox "��ϲ����ս�ɹ���������һ��", , "��ϲ"
    If slta > 3 Then
        MsgBox "��ϲ����������˽��*200�����һظ�״̬ HP+300", , "��ϲ"
        Money = Money + 200
        Label19(4).Caption = Money
        Text4(0).Text = Val(Text4(0).Text) + 300
        Call Save
    End If
If slta = 1 Then Call slta1
If slta = 2 Then Call slta2
If slta = 3 Then Call slta3
If slta = 4 Then Call slta4
If slta = 5 Then Call slta5
If slta = 6 Then Call slta6
If slta = 7 Then Call slta7
If slta = 8 Then Call slta8
If slta = 9 Then Call slta9
If slta = 10 Then Call slta10
Else
MsgBox "��ϲ����ս�ɹ�������5000��ң�С��ҩ10��������ҩ�����棩3����", , "��ϲ"
Label19(5) = 1
Label19(6) = Label19(6) + 10
Label19(7) = Label19(7) + 3
Money = Money + 5000
Label19(4).Caption = Money
Call Save
Call Form_Load
End If
End Sub
Public Sub fail()
If MsgBox("�Ƿ�ʹ�ø���ҩ", vbYesNo, "��ʾ") = vbYes Then
Call Cmd1
Else
MsgBox "��սʧ�ܣ����ص�һ��", , "��ʾ"
Label19(5) = 1
If Money >= 100 And MsgBox("��ȷ������100��ҹ�����ս�ʸ���", vbYesNo, "��ʾ") = vbYes Then
    Money = Money - 100
    Unload Me
    Form7.Show
Else
    MsgBox "��û����ս�ʸ�", , "��ʾ"
    Unload Me
    Form1.Show
    End If
Call Save
Call Form_Load
End If
End Sub
Public Sub slta1()
Text2.Text = "������"
Text3.Text = "��ӭ��������֮����һ�㣬�������������ݣ����Գ��ͣ����Կ�����ɱ��������ܰ��ʾ������ˢǮ����������"
Text5(0).Text = 50
Text5(1).Text = 20
Text5(2).Text = 10
Text5(3).Text = 15
Text5(4).Text = 14
End Sub
Public Sub slta2()
Text2.Text = "���ټ�"
Text3.Text = "��ӭ��������֮���ڶ��㣬�����������ټ�ͯЬ�����ڳ��˵����ԣ������㲻�������⣡��"
Text5(0).Text = 100
Text5(1).Text = 50
Text5(2).Text = 10
Text5(3).Text = 42
Text5(4).Text = 41
End Sub
Public Sub slta3()
Text2.Text = "�Կ���"
Text3.Text = "��ӭ��������֮�������㣬���������Կ��ܣ�����ȫȫ��һ����ܣ�����Ѫ����������ֵ��Ϊ0�������������κ��˺���ͨ��ֻ��ʱ�����⣡��"
Text5(0).Text = 2500
Text5(1).Text = 0
Text5(2).Text = 0
Text5(3).Text = 0
Text5(4).Text = 0
End Sub
Public Sub slta4()
Text2.Text = "������"
Text3.Text = "��ӭ��������֮�����Ĳ㣬��������С�������������ԣ���"
Text5(0).Text = 300
Text5(1).Text = 100
Text5(2).Text = 20
Text5(3).Text = 50
Text5(4).Text = 50
End Sub
Public Sub slta5()
Text2.Text = "������"
Text3.Text = "��ӭ��������֮������㣬��������С�������߹������ԣ�С�ı���ɱ����"
Text5(0).Text = 100
Text5(1).Text = 300
Text5(2).Text = 100
Text5(3).Text = 50
Text5(4).Text = 50
End Sub
Public Sub slta6()
Text2.Text = "������"
Text3.Text = "��ӭ��������֮�������㣬��������С��Ƥ�����߷���������Ѫ��������û�й���������ǿ��̹�ˣ���"
Text5(0).Text = 2000
Text5(1).Text = 0
Text5(2).Text = 500
Text5(3).Text = 0
Text5(4).Text = 0
End Sub
Public Sub slta7()
Text2.Text = "÷����"
Text3.Text = "��ӭ��������֮�����߲㣬��������ú�ϰ壬���ڳ��˵����ԣ���"
Text5(0).Text = 500
Text5(1).Text = 200
Text5(2).Text = 200
Text5(3).Text = 50
Text5(4).Text = 50
End Sub
Public Sub slta8()
Text2.Text = "�P����"
Text3.Text = "��ӭ��������֮���ڰ˲㣬��������С�P���ٶȳ��죡�����ٶȿ���ʲô�ã�"
Text5(0).Text = 300
Text5(1).Text = 200
Text5(2).Text = 100
Text5(3).Text = 5000
Text5(4).Text = 50
End Sub
Public Sub slta9()
Text2.Text = "������"
Text3.Text = "��ӭ��������֮���ھŲ㣬��������Сè��׼Boss���𣡣�"
Text5(0).Text = 3000
Text5(1).Text = 100
Text5(2).Text = 200
Text5(3).Text = 50
Text5(4).Text = 50
End Sub
Public Sub slta10()
Text2.Text = "�˹���"
Text3.Text = "��ӭ��������֮����ʮ�㣬����������Ⱥ�����ˣ������񼶱𣡣����һ���ˣ���Ӯ�����ܵõ����������"
Text5(0).Text = 5000
Text5(1).Text = 500
Text5(2).Text = 500
Text5(3).Text = 500
Text5(4).Text = 500
End Sub

Private Sub Form_Unload(Cancel As Integer) 'ж�ش����¼�
Timer1.Enabled = False
On Error Resume Next
Call Save
End Sub

Private Sub Save()
Dim a As String
Dim b As String
a = "update zc set ʤ��= '" & Label19(1) & "',ʧ��='" & Label19(3) & "',����֮��='" & Label19(5) & "',С��ҩ='" & Label19(6) & "',����ҩ='" & Label19(7) & "',��Ǯ='" & Label19(4) & "',�ȼ�=" & Val(Label19(2)) & " where �˺�=" & Val(Label19(0))
Call CnSql(a, 2)
b = "select * from zc where �˺�=" & Val(Label19(0))
Call CnSql(b, 1)
End Sub

