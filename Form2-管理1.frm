VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "教务管理"
   ClientHeight    =   3735
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   3735
   ScaleWidth      =   6510
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Menu a 
      Caption         =   "文件（A）"
      Index           =   1
      Begin VB.Menu a1 
         Caption         =   "个人信息"
      End
      Begin VB.Menu a2 
         Caption         =   "修改信息"
      End
      Begin VB.Menu a3 
         Caption         =   "复制"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu a4 
         Caption         =   "黏贴"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu b 
      Caption         =   "系统管理(B)"
      Index           =   2
      Begin VB.Menu b1 
         Caption         =   "返回登入界面"
      End
      Begin VB.Menu b2 
         Caption         =   "关于"
         Enabled         =   0   'False
      End
      Begin VB.Menu b3 
         Caption         =   "管理员"
      End
      Begin VB.Menu b4 
         Caption         =   "退出系统"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu c 
      Caption         =   "课程管理(C)"
      Begin VB.Menu c1 
         Caption         =   "课程查询"
      End
      Begin VB.Menu c2 
         Caption         =   "课程管理"
         Shortcut        =   ^A
      End
      Begin VB.Menu c3 
         Caption         =   "课程表"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu d 
      Caption         =   "成绩管理（D）"
      Begin VB.Menu d1 
         Caption         =   "学生查询"
         Shortcut        =   ^X
      End
      Begin VB.Menu d2 
         Caption         =   "班级成绩查询"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu f 
      Caption         =   "教师查询（F）"
      Begin VB.Menu f1 
         Caption         =   "教师查询"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a1_Click()
Form2.Hide
Form5.Show

End Sub

Private Sub a2_Click()
Form2.Hide
Form6.Show
End Sub

Private Sub b1_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub b3_Click()
Dim a$
a = InputBox("请输入管理员的密码 ", "密码")
If a <> "01" Then
MsgBox "密码错误，请重新输入", vbDefaultButton1
a = InputBox("请输入管理员的密码 ", "密码")
Else
Form2.Hide
Form4.Show
End If
End Sub

Private Sub b4_Click()
End
End Sub

Private Sub c1_Click()
Form2.Hide
form8.Show
End Sub

Private Sub c2_Click()
Form2.Hide
form7.Show
End Sub


Private Sub d1_Click()
Form2.Hide
form9.Show
End Sub




Private Sub f1_Click()
Form2.Hide
jiaoshi.Show
End Sub

Private Sub Form_Load()
'Label1.Left = Form2.Width - Label1.Width
'label1.Top = Form2.Height - Label1.Height

End Sub

Private Sub Timer1_Timer()
Label1.Caption = Now
End Sub
