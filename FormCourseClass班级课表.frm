VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCourseClass 
   Caption         =   "班级课表"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8580
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fmClass 
      Caption         =   "班级信息"
      Height          =   1455
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "班级编号"
         Height          =   255
         Left            =   30
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "班级名称"
         Height          =   255
         Left            =   30
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "所属院系"
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "班主任"
         Height          =   255
         Left            =   3030
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "入学日期"
         Height          =   255
         Left            =   3030
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "人数"
         Height          =   255
         Left            =   3030
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblClassID 
         BackColor       =   &H80000009&
         Caption         =   "班级编号"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label lblClassName 
         BackColor       =   &H80000009&
         Caption         =   "班级名称"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label lblDepartName 
         BackColor       =   &H80000009&
         Caption         =   "所属院系"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label lblMaster 
         BackColor       =   &H80000009&
         Caption         =   "班主任"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label lblBeginDate 
         BackColor       =   &H80000009&
         Caption         =   "入学日期"
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
         Left            =   3960
         TabIndex        =   6
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label lblCount 
         BackColor       =   &H80000009&
         Caption         =   "人数"
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
         Left            =   3960
         TabIndex        =   5
         Top             =   1080
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "退出(&Q)"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgCourseSelect 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6376
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
   Begin VB.ComboBox cmbClass 
      Height          =   300
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1695
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
Attribute VB_Name = "frmCourseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub cmbClass_Click()
  ClassInfo
  FullGrid
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

Private Sub ClassInfo()
  Dim rsClass As ADODB.Recordset
  Dim strSQL$
  
  Set rsClass = New ADODB.Recordset
  strSQL = "SELECT Class.*, Department.DepartName FROM Class, Department"
  strSQL = strSQL & " WHERE Class.DepartID=Department.DepartID "
  strSQL = strSQL & " AND ClassID='" & cmbClass.Text & "';"
  rsClass.Open strSQL, Conn, 1, 1
  If Not rsClass.EOF Then
    lblClassID = rsClass("ClassID").Value
    lblClassName = rsClass("ClassName").Value
    lblDepartName = rsClass("DepartName").Value
    lblMaster = rsClass("Master").Value
    lblBeginDate = Str(rsClass("BeginDate").Value)
  Else
    lblClassID = ""
    lblClassName = ""
    lblDepartName = ""
    lblMaster = ""
    lblBeginDate = ""
  End If
  rsClass.Close
  
  strSQL = "SELECT Count(*) as StudentCount FROM Student WHERE ClassID='" & cmbClass.Text & "';"
  rsClass.Open strSQL, Conn, 1, 1
  lblCount = Str(rsClass("StudentCount").Value)
  rsClass.Close
  
  Set rsClass = Nothing
End Sub

Private Sub FullGrid()
  If rs.State <> adStateClosed Then rs.Close
  
  strSQL = "SELECT CourseSelect.ID, CourseName, ClassRoom, ClassDate, Credit, Period, Teacher FROM Course, CourseSelect "
  strSQL = strSQL & " WHERE Course.CourseID = CourseSelect.CourseID"
  strSQL = strSQL & " AND CourseSelect.ClassID='" & cmbClass.Text & "'"
  rs.Open strSQL, Conn, 3, 3
  
  Set dgCourseSelect.DataSource = rs
  With dgCourseSelect
    .Columns(0).Visible = False
    .Columns(1).Width = 2100
    .Columns(1).Caption = "课程名称"
    .Columns(2).Width = 900
    .Columns(2).Caption = "上课地点"
    .Columns(3).Width = 900
    .Columns(3).Caption = "上课时间"
    .Columns(4).Width = 900
    .Columns(4).Caption = "学分"
    .Columns(5).Width = 900
    .Columns(5).Caption = "学时数"
    .Columns(6).Width = 1500
    .Columns(6).Caption = "任课老师"
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set dgCourseSelect.DataSource = Nothing
  rs.Close
  Conn.Close
  Set rs = Nothing
  Set Conn = Nothing
End Sub
