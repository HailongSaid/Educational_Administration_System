VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "用户注册"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form3"
   ScaleHeight     =   5025
   ScaleWidth      =   6465
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "UserID"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   615
      Index           =   2
      Left            =   2160
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   615
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "立即注册"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      DataField       =   "Description"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "UserName"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008080&
      Caption         =   "Password again:"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008080&
      Caption         =   "Description :"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008080&
      Caption         =   "UserName :"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008080&
      Caption         =   "Password :"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   "UserID :"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private cnn As New ADODB.Connection
Private Rs As New ADODB.Recordset
Private Sub Command1_Click(Index As Integer)
Dim a%
a = 1
Select Case Index
Case 0

If Text5.Text = "" Or Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "UserID Password UserName 不能为空，请输入", vbOKOnly
a = 0
Text1.SetFocus
Exit Sub
End If
If Text2.Text <> Text5.Text Then
MsgBox " Password 和 Password again 不一致，请正确输入", vbOKOnly
a = 0
Text2.SetFocus
Exit Sub
End If
If a = 1 Then
cnn.Execute "insert into Users values ('" & Text1.Text & "','" & Text2.Text & " ','" & Text3.Text & "','" & Text4.Text & "')"
MsgBox "注册成功", vbOKOnly
End If
Case 1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Case 2
Rs.Close
cnn.Close
Form3.Hide
Form1.Show
End Select
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
'cnn.Open "DSN=db2;User ID=;pwd="
Rs.Open "select * from Users", cnn, 3, 3
End Sub

