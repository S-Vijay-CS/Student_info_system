VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmattendance 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form2"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9990
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   9990
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   0
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Attendance"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmattendance1.frx":0000
      Height          =   1695
      Left            =   600
      TabIndex        =   20
      Top             =   6720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H0080FF80&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   5880
      Picture         =   "frmattendance1.frx":0015
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   2880
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5880
      Picture         =   "frmattendance1.frx":1602F
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   12
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtpercentage 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox txtabsentday 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txtworkingday 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Absent Day"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Present days"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stu _ Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stu _ ID"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDANCE DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "frmattendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim str As String
Dim cn As New ADODB.Connection

Private Sub cmdfirst_Click(Index As Integer)
rs.MoveFirst
display
End Sub

Private Sub cmdlast_Click(Index As Integer)
rs.MoveLast
display
End Sub

Private Sub cmdnext_Click(Index As Integer)
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If

End Sub

Private Sub cmdpre_Click(Index As Integer)
rs.MovePrevious
If Not rs.BOF Then
display
Else
rs.MoveLast
display
End If

End Sub

Private Sub cmdupdate_Click()
rs.Fields("Stu_ID").Value = txtid.Text
rs.Fields("Stu_Name").Value = txtname.Text
rs.Fields("Course").Value = Combo1.Text
rs.Fields("TotalPresent").Value = txtworkingday.Text
rs.Fields("TotalAbsent").Value = txtabsentday.Text
rs.Fields("Percentage").Value = txtpercentage.Text
MsgBox "Information saved ", vbInformation, Attendance_details
rs.Update
End Sub

Private Sub Command1_Click()
rs.AddNew
rs.Fields("Stu_ID").Value = txtid.Text
rs.Fields("Stu_Name").Value = txtname.Text
rs.Fields("Course").Value = Combo1.Text
rs.Fields("TotalPresent").Value = txtworkingday.Text
rs.Fields("TotalAbsent").Value = txtabsentday.Text
rs.Fields("Percentage").Value = txtpercentage.Text
MsgBox "Information saved ", vbInformation, Attendance_details
rs.Update
End Sub

Private Sub Command2_Click()
txtworkingday.Text = ""
txtabsentday.Text = ""
txtpercentage.Text = ""
End Sub

Private Sub Command3_Click()
confirm = MsgBox("Do you want to delete the STUDENT DETAILS", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox " Record has been deleted successfully", vbInformation, "Message"
rs.Update
Else
MsgBox "Record not deleted...!!!", vbInformation, "Message"
End If

End Sub

Private Sub Command4_Click()
cn.Close
Unload Me
End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
cn.Open
rs.Open "select * from Attendance", cn, adOpenDynamic, adLockPessimistic
Adodc1.Visible = False

End Sub
Sub reload()
rs.Close
rs.Open "select*from Attendance", con, adOpenDynamic, adLockPessimistic

End Sub

Sub display()
txtid.Text = rs!AdmissionNo
txtname.Text = rs!Name
Combo1.Text = rs!Course
txtworkingday.Text = rs!Totalworkingday
txtabsentday.Text = rs!Totalabsentday
txtpercentage.Text = rs!percentage
Image1.Picture = LoadPicture(rs!Photo)
End Sub

Private Sub Picture1_Click()
rs.Close
rs.Open "Select*from Stu_regi where AdmissionNo='" + txtid.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found..!!!", vbInformation
End If

End Sub

Private Sub Picture2_Click()
rs.Close
rs.Open "Select*from AdmissionTBL where AdmissionNo='" + txtname.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found..!!!", vbInformation
End If

End Sub

Private Sub txtpercentage_GotFocus()
Dim y As Integer
y = txtworkingday.Text + txtabsentday.Text / 100
txtpercentage = y
End Sub
