VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmScoreClass 
   Caption         =   "班级成绩分析"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9375
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "退出(&Q)"
      Height          =   375
      Left            =   8160
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fmCourse 
      Caption         =   "课程信息"
      Height          =   2415
      Left            =   6360
      TabIndex        =   5
      Top             =   600
      Width           =   2895
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "任课老师"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "课程编号"
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "课程名称"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "学分"
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "学时数"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblTeacher 
         BackColor       =   &H8000000E&
         Caption         =   "任课老师"
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
         Left            =   1050
         TabIndex        =   10
         Top             =   1920
         Width           =   1770
      End
      Begin VB.Label lblCourseID 
         BackColor       =   &H8000000E&
         Caption         =   "课程编号"
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
         Left            =   1050
         TabIndex        =   9
         Top             =   480
         Width           =   1770
      End
      Begin VB.Label lblCourseName 
         BackColor       =   &H8000000E&
         Caption         =   "课程名称"
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
         Left            =   1050
         TabIndex        =   8
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label lblCredit 
         BackColor       =   &H8000000E&
         Caption         =   "学分"
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
         Left            =   1050
         TabIndex        =   7
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label lblPeriod 
         BackColor       =   &H8000000E&
         Caption         =   "学时数"
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
         Left            =   1050
         TabIndex        =   6
         Top             =   1560
         Width           =   1770
      End
   End
   Begin MSDataGridLib.DataGrid dgScore 
      Height          =   6135
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   10821
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "考试成绩统计"
      Height          =   3135
      Left            =   6360
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
      Begin VB.Label lblJGL 
         BackColor       =   &H8000000E&
         Caption         =   "及格率"
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
         Left            =   1500
         TabIndex        =   28
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "及格率"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblRS 
         BackColor       =   &H8000000E&
         Caption         =   "考试人数"
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
         Left            =   1500
         TabIndex        =   26
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblAvg 
         BackColor       =   &H8000000E&
         Caption         =   "平均分"
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
         Left            =   1500
         TabIndex        =   25
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "平均分"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblMax 
         BackColor       =   &H8000000E&
         Caption         =   "最高分"
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
         Left            =   1500
         TabIndex        =   23
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label lblBJG 
         BackColor       =   &H8000000E&
         Caption         =   "不及格人数"
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
         Left            =   1500
         TabIndex        =   22
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label lblMin 
         BackColor       =   &H8000000E&
         Caption         =   "最低分"
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
         Left            =   1500
         TabIndex        =   21
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "最高分"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "不及格人数"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "参加考试总人数"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "最低分"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.ListBox lstCourse 
      Height          =   5100
      Left            =   120
      TabIndex        =   2
      Top             =   1155
      Width           =   2055
   End
   Begin VB.ComboBox cmbClass 
      Height          =   300
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "课程列表"
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "班级编号"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmScoreClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub cmbClass_Click()
  FullCourse
  lstCourse_Click
End Sub

Private Sub cmdQuit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set Conn = New ADODB.Connection
  Set rs = New ADODB.Recordset
  Conn.Open ConnString
  
  FullClass
  cmbClass_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set dgScore.DataSource = Nothing
  rs.Close
  Conn.Close
  Set rs = Nothing
  Set Conn = Nothing
End Sub

Private Sub lstCourse_Click()
  FullGrid
  CourseInfo
  StatisticalInfo
End Sub

Private Sub FullClass()
  Dim rsClass As ADODB.Recordset
  Dim strSQL$
  
  Set rsClass = New ADODB.Recordset
  strSQL = "SELECT * FROM Class"
  rsClass.Open strSQL, Conn, 1, 1
  cmbClass.Clear
  While Not rsClass.EOF
    cmbClass.AddItem rsClass("ClassID").Value
    rsClass.MoveNext
  Wend
  If cmbClass.ListCount > 0 Then cmbClass.ListIndex = 0
  rsClass.Close
  Set rsClass = Nothing
End Sub

Private Sub FullCourse()
  Dim rsCourse As ADODB.Recordset
  Dim strSQL$
  
  Set rsCourse = New ADODB.Recordset
  strSQL = "SELECT * FROM CourseSelect WHERE ClassID='" & cmbClass.Text & "'"
  rsCourse.Open strSQL, Conn, 1, 1
  lstCourse.Clear
  While Not rsCourse.EOF
    lstCourse.AddItem rsCourse("CourseID").Value
    rsCourse.MoveNext
  Wend
  If lstCourse.ListCount > 0 Then lstCourse.ListIndex = 0
  rsCourse.Close
  Set rsCourse = Nothing
End Sub

Private Sub FullGrid()
  If rs.State <> adStateClosed Then rs.Close
  
  strSQL = "SELECT Score.ID, Score.StudentID, Student.Name, Score.Score  FROM Score, Student "
  strSQL = strSQL & " WHERE Student.StudentID = Score.StudentID"
  strSQL = strSQL & " AND Score.CourseID='" & lstCourse.Text & "'"
  strSQL = strSQL & " AND Student.ClassID='" & cmbClass.Text & "'"
  rs.Open strSQL, Conn, 3, 3
  
  Set dgScore.DataSource = rs
  With dgScore
    .Columns(0).Visible = False
    .Columns(1).Width = 900
    .Columns(1).Caption = "学号"
    .Columns(2).Width = 900
    .Columns(2).Caption = "姓名"
    .Columns(3).Width = 900
    .Columns(3).Caption = "成绩"
  End With
End Sub

Private Sub CourseInfo()
  Dim rsCourse As ADODB.Recordset
  Dim strSQL$
  
  Set rsCourse = New ADODB.Recordset
  strSQL = "SELECT * FROM Course WHERE CourseID='" & lstCourse.Text & "';"
  rsCourse.Open strSQL, Conn, 1, 1
  If Not rsCourse.EOF Then
    lblCourseID = rsCourse("CourseID").Value
    lblCourseName = rsCourse("CourseName").Value
    lblCredit = Str(rsCourse("Credit").Value)
    lblPeriod = Str(rsCourse("Period").Value)
    lblTeacher = rsCourse("Teacher").Value
  Else
    lblCourseID = ""
    lblCourseName = ""
    lblCredit = ""
    lblPeriod = ""
    lblTeacher = ""
  End If
  rsCourse.Close
  Set rsCourse = Nothing
End Sub

Private Sub StatisticalInfo()
  Dim rsCourse As ADODB.Recordset
  Dim strSQL$
  Dim intRS!, intBJG!, fRate As Single
  
  Set rsCourse = New ADODB.Recordset
  strSQL = "SELECT Count(*) as RS, Max(Score) as Max, Min(Score) as Min, Avg(Score) as Avg  FROM Score WHERE CourseID='" & lstCourse.Text & "';"
  rsCourse.Open strSQL, Conn, 1, 1
  If Not rsCourse.EOF Then
    intRS = rsCourse("RS").Value
    If intRS > 0 Then
      lblRS = Str(rsCourse("RS").Value)
      lblMax = Str(rsCourse("Max").Value)
      lblMin = Str(rsCourse("Min").Value)
      lblAvg = Str(rsCourse("Avg").Value)
    Else
      lblRS = ""
      lblMax = ""
      lblMin = ""
      lblAvg = ""
    End If
  Else
    lblRS = ""
    lblMax = ""
    lblMin = ""
    lblAvg = ""
  End If
  rsCourse.Close
  
  Set rsCourse = New ADODB.Recordset
  strSQL = "SELECT Count(*) as BJG FROM Score WHERE Score < 60 AND CourseID='" & lstCourse.Text & "';"
  rsCourse.Open strSQL, Conn, 1, 1
  If Not rsCourse.EOF And intRS > 0 Then
    intBJG = rsCourse("BJG").Value
    fRate = Int((intRS - intBJG) / intRS * 1000) / 10
    lblJGL = Str(fRate) & " %"
    lblBJG = Str(intBJG)
  Else
    lblJGL = ""
    lblBJG = ""
  End If
  rsCourse.Close
  
  Set rsCourse = Nothing
End Sub
