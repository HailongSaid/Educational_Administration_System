VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "登入界面"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6690
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4680
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注册"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   3960
      Top             =   2400
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "1"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登入"
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
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
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "欢迎来到教务系统"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "密码："
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名："
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As New ADODB.Connection
Private Rs As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
Dim a$
Dim b$
'Data1.Recordset.FindFirst "UserName =" & "'" & Text4.Text & "'"

Select Case Index
Case 0
Text2.Text = ""
Text1.Text = ""
Text1.SetFocus
Case 1
Set cnn = New ADODB.Connection
'cnn.Open "DSN=db2;User ID=;pwd="
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
Set Rs = New ADODB.Recordset
Rs.Open "select * from Users where UserName ='" & Text1.Text & "'and Password = '" & Text2.Text & "' ", cnn, 1, 1

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "用户名或者密码不能为空", vbOKOnly
Text1.SetFocus
Exit Sub
End If
If Rs.RecordCount = 0 Then
MsgBox "用户名或者密码错误，请重新输入", vbOKOnly
Text1.SetFocus
Exit Sub
End If
If Rs.RecordCount <> 0 Then
Rs.Close
cnn.Close
Form1.Hide
Form2.Show
End If
Case 2
Form1.Hide
Form3.Show
End Select
End Sub



Private Sub Label3_Click()
If Label1.Font.Size < 40 Then
Label1.Font.Size = Label1.Font.Size + 1
Else
Label1.Font.Size = Label1.Font.Size - 1
End If
End Sub

Private Sub Text2_Change()
Text2.PasswordChar = "*"
End Sub
Private Sub Form_Load()

'rs.Open "select * from Users ", "DSN=db2;User ID=;pwd=", 3, 3
'Text2.Text = ""
'Text1.Text = ""
'Label4.Left = Form1.Width - Label4.Width
'Label4.Top = Form1.Height - Label4.Height
End Sub

Private Sub Timer1_Timer()
If Label3.Left >= Form1.Width Then
Label3.Left = 0
End If
If Label3.Left = 0 Then Label3.Left = Label3.Left + 20
If Label3.Left <> 0 And Label3.Left < Form1.Width Then Label3.Left = Label3.Left + 20
End Sub

Private Sub Timer2_Timer()
Label4.Caption = Now
End Sub
