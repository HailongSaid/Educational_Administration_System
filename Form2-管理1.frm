VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "�������"
   ClientHeight    =   3735
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   3735
   ScaleWidth      =   6510
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�ļ���A��"
      Index           =   1
      Begin VB.Menu a1 
         Caption         =   "������Ϣ"
      End
      Begin VB.Menu a2 
         Caption         =   "�޸���Ϣ"
      End
      Begin VB.Menu a3 
         Caption         =   "����"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu a4 
         Caption         =   "���"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu b 
      Caption         =   "ϵͳ����(B)"
      Index           =   2
      Begin VB.Menu b1 
         Caption         =   "���ص������"
      End
      Begin VB.Menu b2 
         Caption         =   "����"
         Enabled         =   0   'False
      End
      Begin VB.Menu b3 
         Caption         =   "����Ա"
      End
      Begin VB.Menu b4 
         Caption         =   "�˳�ϵͳ"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu c 
      Caption         =   "�γ̹���(C)"
      Begin VB.Menu c1 
         Caption         =   "�γ̲�ѯ"
      End
      Begin VB.Menu c2 
         Caption         =   "�γ̹���"
         Shortcut        =   ^A
      End
      Begin VB.Menu c3 
         Caption         =   "�γ̱�"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu d 
      Caption         =   "�ɼ�����D��"
      Begin VB.Menu d1 
         Caption         =   "ѧ����ѯ"
         Shortcut        =   ^X
      End
      Begin VB.Menu d2 
         Caption         =   "�༶�ɼ���ѯ"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu f 
      Caption         =   "��ʦ��ѯ��F��"
      Begin VB.Menu f1 
         Caption         =   "��ʦ��ѯ"
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
a = InputBox("���������Ա������ ", "����")
If a <> "01" Then
MsgBox "�����������������", vbDefaultButton1
a = InputBox("���������Ա������ ", "����")
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
