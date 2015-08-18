VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "个人信息"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form3"
   ScaleHeight     =   5010
   ScaleWidth      =   6420
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "UserID"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "Description"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "UserName"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008080&
      Caption         =   "Description :"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008080&
      Caption         =   "UserName :"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008080&
      Caption         =   "Password :"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   "UserID :"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cnn As New ADODB.Connection
Private Rs As New ADODB.Recordset

Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Form5.Hide
Form2.Show
End Sub
Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
'cnn.Open "DSN=db2;User ID=;pwd="
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
Rs.Open "select * from Users where UserName='" & Form1.Text1.Text & "' and Password = '" & Form1.Text2.Text & "' ", cnn, 3, 3
Call ShowData
Rs.Close
cnn.Close
End Sub
Private Sub ShowData()
Text1.Text = Rs.Fields(0).Value
Text2.Text = Rs.Fields(1).Value
Text3.Text = Rs.Fields(2).Value
Text4.Text = Rs.Fields(3).Value
End Sub

