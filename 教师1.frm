VERSION 5.00
Begin VB.Form jiaoshi 
   Caption         =   "教师查询"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   5790
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "返回"
      Height          =   495
      Left            =   4800
      TabIndex        =   39
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "记录操作"
      Height          =   735
      Left            =   120
      TabIndex        =   34
      Top             =   6000
      Width           =   4815
      Begin VB.CommandButton Command4 
         Caption         =   "取消"
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "确定"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "修改记录"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "添加记录"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "按教师编号查找"
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   4815
      Begin VB.CommandButton Command3 
         Caption         =   "最后一条"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "下一条"
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "上一条"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "第一条"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Text            =   "Text10"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "教师 编号"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "按姓名查找"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   4815
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Text            =   "Text9"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "最后一条"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "下一条"
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上一条"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "第一条"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "姓名"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "移动到"
      Height          =   495
      Index           =   4
      Left            =   3240
      TabIndex        =   17
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最后一条"
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   16
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一条"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一条"
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一条"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000C0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "工资"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "单位"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "职称"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "出生日期"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "性别"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "姓名"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "教师编号"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "jiaoshi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As New ADODB.Connection
Private Rs As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
Rs.Open "select * from 教师表", cnn, 3, 3

Dim n%
n = Val(Text8.Text)
 Select Case Index
   Case 0
   Rs.MoveFirst
   Call ShowData
   Command1(1).Enabled = False
   Command1(2).Enabled = True
   Command1(3).Enabled = True
   Case 1
   Rs.MovePrevious
   Call ShowData
   Command1(0).Enabled = True
   Command1(1).Enabled = True
   Command1(2).Enabled = True
   Command1(3).Enabled = True
   Case 2
   Rs.MoveNext
   Call ShowData
   Command1(0).Enabled = True
   Command1(1).Enabled = True
   Command1(2).Enabled = True
   Command1(3).Enabled = True
   Case 3
   Rs.MoveLast
   Call ShowData
   Command1(0).Enabled = True
   Command1(1).Enabled = True
   Command1(2).Enabled = False
   Command1(3).Enabled = True
   Case 4
   If n > Rs.RecordCount Then
   MsgBox " 此表没有那么多记录！"
   Else
 Rs.Move (n)
   Call ShowData
   End If
   End Select
   Rs.Close
End Sub

Private Sub Command2_Click(Index As Integer)
Rs.Open "select * from 教师表 where 姓名 like " & "'" & Text9.Text & "'", cnn, 3, 3
If Rs.RecordCount = 0 Then
MsgBox " 此表不存在这个姓名！"
Rs.Close
Else
Select Case Index
Case 0
  Rs.MoveFirst
     Call ShowData

  Case 1
   Rs.MovePrevious
      Call ShowData
 Case 2
   Rs.MoveNext
      Call ShowData
Case 3
  Rs.MoveLast
      Call ShowData
End Select
Rs.Close
End If

End Sub



Private Sub Command3_Click(Index As Integer)

Rs.Open "select * from 教师表 where 教师编号 like " & "'" & Text10.Text & "'", cnn, 3, 3
If Rs.RecordCount = 0 Then
MsgBox " 此表不存在这个姓名！"
Rs.Close
Else
Select Case Index
Case 0
  Rs.MoveFirst
     Call ShowData

  Case 1
   Rs.MovePrevious
      Call ShowData
 Case 2
   Rs.MoveNext
      Call ShowData
Case 3
  Rs.MoveLast
      Call ShowData
End Select
Rs.Close
End If
End Sub

Private Sub Command4_Click(Index As Integer)
Rs.Open "select * from 教师表", cnn, 3, 3
Select Case Index
Case 0
 'Rs.AddNew
   Command4(1).Enabled = False
  Case 1
 'Rs.Edit
     Command4(0).Enabled = False
 Case 2

 'cnn.Execute "insert into 教师表 value（'" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "')"

  Rs.Update
     Command4(0).Enabled = True
       Command4(1).Enabled = True
         Command4(2).Enabled = True
           Command4(3).Enabled = True
           
Case 3
 ' Rs.Save
   Command4(0).Enabled = True
       Command4(1).Enabled = True
         Command4(2).Enabled = True
           Command4(3).Enabled = True
          
End Select
Rs.Close
End Sub

Private Sub Command5_Click()
cnn.Close
jiaoshi.Hide
Form2.Show
End Sub

Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
 cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
' cnn.Open "DSN=db1;User ID=;pwd="
Rs.Open "select * from 教师表", cnn, 3, 3
Call ShowData
Rs.Close
End Sub
Private Sub ShowData()
Text2.Text = Rs.Fields("姓名").Value
Text1.Text = Rs.Fields("教师编号").Value
Text3.Text = Rs.Fields("性别").Value
'Text4.Text = rs.Fields("出生日期").Value
Text5.Text = Rs.Fields("职称").Value
Text6.Text = Rs.Fields("单位").Value
Text7.Text = Rs.Fields("工资").Value
End Sub
Private Sub updata()
 Rs.Fields("姓名").Value = Text2.Text
 Rs.Fields("教师编号").Value = Text1.Text
 Rs.Fields("性别").Value = Text3.Text
 Rs.Fields("出生日期").Value = Text4.Text
Rs.Fields("职称").Value = Text5.Text
 Rs.Fields("单位").Value = Text6.Text
 Rs.Fields("工资").Value = Text7.Text
End Sub

