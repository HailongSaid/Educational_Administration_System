VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "信息修改"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form3"
   ScaleHeight     =   5010
   ScaleWidth      =   6420
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   495
      Index           =   2
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000C0C0&
      DataField       =   "Password"
      DataSource      =   "Data1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000C0C0&
      DataField       =   "UserID"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0000C0C0&
      DataField       =   "Description"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0000C0C0&
      DataField       =   "UserName"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Description :"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "UserName :"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password :"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "UserID :"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cnn As New ADODB.Connection
Private Rs As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)

Select Case Index
Case 0
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "UserID Password UserName 不能为空，请输入", vbOKOnly
End If
Rs.Update
MsgBox " 修改成功！"
Case 1
Rs.Cancel
Case 2
Rs.Close
cnn.Close
Form6.Hide
Form2.Show
End Select
End Sub

Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
'cnn.Open "DSN=db2;User ID=;pwd="
Rs.Open "select * from Users where UserName='" & Form1.Text1.Text & "'and Password = '" & Form1.Text2.Text & "'", cnn, 3, 3
Call ShowData
End Sub
Private Sub ShowData()
Text1.Text = Rs.Fields(0).Value
Text2.Text = Rs.Fields(1).Value
Text3.Text = Rs.Fields(2).Value
Text4.Text = Rs.Fields(3).Value
End Sub
