VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form8 
   Caption         =   "课程查询"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8145
   StartUpPosition =   1  '所有者中心
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1720
      _Version        =   393216
      BackColor       =   16448
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame fmCourse 
      Caption         =   "课程信息"
      Height          =   1455
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "任课老师"
         Height          =   255
         Left            =   30
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "课程编号"
         Height          =   255
         Left            =   30
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "课程名称"
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "学分"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "学时数"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblTeacher 
         BackColor       =   &H00004040&
         Caption         =   "任课老师"
         DataField       =   "Teacher"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   930
         TabIndex        =   6
         Top             =   1080
         Width           =   1770
      End
      Begin VB.Label lblCourseID 
         BackColor       =   &H00004040&
         Caption         =   "课程编号"
         DataField       =   "CourseID"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label lblCourseName 
         BackColor       =   &H00004040&
         Caption         =   "课程名称"
         DataField       =   "CourseName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   930
         TabIndex        =   4
         Top             =   720
         Width           =   1770
      End
      Begin VB.Label lblCredit 
         BackColor       =   &H00004040&
         Caption         =   "学分"
         DataField       =   "Credit"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3690
         TabIndex        =   3
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblPeriod 
         BackColor       =   &H00004040&
         Caption         =   "学时数"
         DataField       =   "Period"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3690
         TabIndex        =   2
         Top             =   720
         Width           =   930
      End
   End
   Begin VB.Label Label2 
      Caption         =   "课程编号"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cnn As New ADODB.Connection
Private Rs As New ADODB.Recordset
Private Rs1 As New ADODB.Recordset


Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
Rs.Open "select * from Course where CourseID like" & "'" & Text1.Text & "'", cnn, 3, 3

If Rs.RecordCount = 0 Then
MsgBox "不存在这个编号！，请重新输入"
Text1.Text = ""
Text1.SetFocus
Else
Call ShowData
Call ShowData1(Rs, MSFlexGrid1)
End If
Rs.Close
'Data1.Recordset.FindFirst "CourseID like" & "'" & Text1.Text & "'"
Case 1
Rs1.Close
cnn.Close
form8.Hide
Form2.Show
End Select
End Sub



Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
'cnn.Open "DSN=db2;User ID=;pwd="
Rs.Open "select * from Course", cnn, 3, 3
Rs1.Open "select * from Course", cnn, 3, 3
Call ShowData
Call ShowData1(Rs, MSFlexGrid1)
Rs.Close
End Sub
Private Sub ShowData()
lblCourseID.Caption = Rs.Fields(0).Value
lblCourseName.Caption = Rs.Fields(1).Value
lblCredit.Caption = Rs.Fields(2).Value
lblPeriod.Caption = Rs.Fields(3).Value
lblTeacher.Caption = Rs.Fields(4).Value
End Sub


