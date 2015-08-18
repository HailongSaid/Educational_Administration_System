VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form7 
   Caption         =   "课程管理"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8715
   StartUpPosition =   1  '所有者中心
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   495
      Left            =   3720
      TabIndex        =   20
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "最后一条"
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一条"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   18
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一条"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   17
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "第一条"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "课程信息"
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   8535
      Begin VB.TextBox Text6 
         BackColor       =   &H00008080&
         DataField       =   "Description"
         DataSource      =   "Data1"
         Height          =   1575
         Left            =   4560
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00008080&
         DataField       =   "Teacher"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00008080&
         DataField       =   "Period"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   1200
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00008080&
         DataField       =   "Credit"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00008080&
         DataField       =   "CourseName"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00008080&
         DataField       =   "CourseID"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         Height          =   180
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Period"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "CourseName"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "CourseID"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Teacher"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "form7"
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
'rs.EditMode
Case 1
'Rs.AddNew
Case 2
If Text1.Text = "" Then
    MsgBox "课程编号不能为空！", , "输入错误"
    Text1.SetFocus
    Exit Sub
  ElseIf Text2.Text = "" Then
    MsgBox "课程名称不能为空！", , "输入错误"
    Text2.SetFocus
    Exit Sub
  ElseIf Text3.Text = 0 Then
    MsgBox "该课程学分不能为空！", , "输入错误"
    Text3.SetFocus
    Exit Sub
  ElseIf Text4.Text = 0 Then
    MsgBox "该课程学时数不能为空！", , "输入错误"
    Text4.SetFocus
    Exit Sub
  ElseIf Text5.Text = "" Then
    MsgBox "任课老师不能为空！", , "输入错误"
    Text5.SetFocus
    Exit Sub
  Else
  'Rs.Update
 ' Rs.Save
 
  End If
End Select
Rs.Close



End Sub

Private Sub Command2_Click(Index As Integer)

Select Case Index
Case 0
Rs.MoveFirst
Call ShowData
Command2(1).Enabled = False
Command2(0).Enabled = True
Command2(2).Enabled = True
Command2(3).Enabled = True
Case 1
Rs.MovePrevious
Call ShowData
Command2(0).Enabled = True
Command2(1).Enabled = True
Command2(2).Enabled = True
Command2(3).Enabled = True
Case 2
Rs.MoveNext
Call ShowData
Command2(0).Enabled = True
Command2(1).Enabled = True
Command2(2).Enabled = True
Command2(3).Enabled = True
Case 3
Rs.MoveLast
Call ShowData
Command2(2).Enabled = False
Command2(0).Enabled = True
Command2(1).Enabled = True
Command2(3).Enabled = True
End Select

End Sub

Private Sub Command3_Click()
form7.Hide
Form2.Show
cnn.Close
End Sub

Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
Rs.Open "select * from Course", cnn, 3, 3
If Rs.AbsolutePosition = Rs.BOF Then Command2(0).Enabled = False
If Rs.AbsolutePosition = Rs.EOF Then Command2(2).Enabled = False
Call ShowData
Call ShowData1(Rs, MSFlexGrid1)
End Sub
Private Sub ShowData()
Text1.Text = Rs.Fields(0).Value
Text2.Text = Rs.Fields(1).Value
Text3.Text = Rs.Fields(2).Value
Text4.Text = Rs.Fields(3).Value
Text5.Text = Rs.Fields(4).Value
Text6.Text = Rs.Fields(5).Value

End Sub
Private Sub updata()
 Rs.Fields(0).Value = Text1.Text
 Rs.Fields(1).Value = Text2.Text
 Rs.Fields(2).Value = Text3.Text
 Rs.Fields(3).Value = Text4.Text
Rs.Fields(4).Value = Text5.Text
 Rs.Fields(5).Value = Text6.Text

End Sub



