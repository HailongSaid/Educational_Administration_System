VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form9 
   Caption         =   "学生成绩查询"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7575
   StartUpPosition =   1  '所有者中心
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   28
      Top             =   3840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16448
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   22
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   21
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "统计信息"
      Height          =   2415
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   3135
      Begin VB.Label Label12 
         BackColor       =   &H00004040&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00004040&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00004040&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00004040&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "最低分"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "不及格门数"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "最高分"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "平均分"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "个人信息"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
      Begin VB.Label lblStudentID 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "学号"
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "出生日期"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "姓名"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "性别"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "所属院系"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "所在班级"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "班主任"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label lblName 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label lblBirthday 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label lblDepart 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Label lblClass 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label lblMaster 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1200
         TabIndex        =   3
         Top             =   2520
         Width           =   1800
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "输入学号"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "form9"
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
Rs.Open "select Student.StudentID,Student.Name,Student.Sex,Student.Birthday , Department.DepartName,Class.ClassName,Class.Master from Student,Class,Department where Student.ClassID=Class.ClassID and Class.DepartID=Department.DepartID and Student.StudentID='" & Text1.Text & "'", cnn, 3, 3
If Rs.RecordCount = 0 Then
MsgBox "不存在这个学号，请重新输入！"
Text1.SetFocus
Exit Sub
Rs.Close
Else
Call ShowData
Call ShowData1(Rs, MSFlexGrid1)
Rs.Close
Rs.Open "select count(*) from Score where Score.Score<60 and Score.StudentID='" & Text1.Text & "' group by StudentID  ", cnn, 3, 3
'Label12(0).Caption = Rs.Fields(0).Value

Rs.Close
Rs.Open "select max(Score),avg(Score),min(Score) from Score where Score.StudentID='" & Text1.Text & "'", cnn, 3, 3
If Rs.RecordCount = 0 Then
MsgBox "不存在这个学号，请重新输入！"
Text1.Text = ""
Text1.SetFocus
Exit Sub
Rs.Close
Else
Label12(1).Caption = Rs.Fields(0).Value
Label12(2).Caption = Rs.Fields(1).Value
Label12(3).Caption = Rs.Fields(2).Value
Rs.Close
End If
End If
'Data1.Recordset.FindFirst "StudentID like" & "'" & Text1.Text & "'"
'Data2.Recordset.FindFirst "ClassID like" & "'" & Text2.Text & "'"
'Data3.Recordset.FindFirst "DepartID like" & "'" & Text3.Text & "'"
Case 1
Rs1.Close
cnn.Close
form9.Hide
Form2.Show
End Select

End Sub



Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set Rs = New ADODB.Recordset
cnn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\db2.mdb"
 'cnn.Open "DSN=db2;User ID=;pwd="
Rs.Open "select Student.StudentID,Student.Name,Student.Sex,Student.Birthday , Department.DepartName,Class.ClassName,Class.Master from Student,Class,Department  where Student.ClassID=Class.ClassID and Class.DepartID=Department.DepartID", cnn, 3, 3
Rs1.Open "select Student.StudentID,Student.Name,Student.Sex,Student.Birthday , Department.DepartName,Class.ClassName,Class.Master from Student,Class,Department  where Student.ClassID=Class.ClassID and Class.DepartID=Department.DepartID", cnn, 3, 3
Call ShowData1(Rs1, MSFlexGrid1)
Call ShowData
Rs.Close
Rs.Open "select count(*) from Score where Score.Score<60 and Score.StudentID='" & lblStudentID.Caption & "' group by StudentID  ", cnn, 3, 3
Label12(0).Caption = Rs.Fields(0).Value
Rs.Close
Rs.Open "select max(Score),avg(Score),min(Score) from Score where Score.StudentID='" & lblStudentID.Caption & "'", cnn, 3, 3
Label12(1).Caption = Rs.Fields(0).Value
Label12(2).Caption = Rs.Fields(1).Value
Label12(3).Caption = Rs.Fields(2).Value
Rs.Close
Rem Data4.RecordSource = "select count(*) from  Score where StudentID = '" & Text1.Text & " and Score <60 group by StudentID "
Rem Data5.RecordSource = "select max(Score) as max min(Score) as min  avg(Score) as avg from  Score where StudentID = '" & Text1.Text & " group by StudentID "
Rem Label12(0).DataField=
End Sub

Private Sub ShowData()
lblStudentID.Caption = Rs.Fields(0).Value
lblName.Caption = Rs.Fields(1).Value
lblSex.Caption = Rs.Fields(2).Value
lblBirthday.Caption = Rs.Fields(3).Value
lblDepart.Caption = Rs.Fields(4).Value
lblClass.Caption = Rs.Fields(5).Value
lblMaster.Caption = Rs.Fields(6).Value
End Sub


