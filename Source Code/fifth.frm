VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fifth 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form15"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   LinkTopic       =   "Form15"
   ScaleHeight     =   10140
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "fifth.frx":0000
      Height          =   2175
      Left            =   0
      TabIndex        =   29
      Top             =   7920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   5520
      Top             =   6120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      RecordSource    =   "Sem5"
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
   Begin VB.ComboBox cmbcourse 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   17
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtsem1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   15
      Text            =   "0"
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtsem2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   14
      Text            =   "0"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtsem3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   13
      Text            =   "0"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtsem4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   12
      Text            =   "0"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtsem5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   11
      Text            =   "0"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtsem6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   10
      Text            =   "0"
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox txttotal 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   9
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox txtaverage 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   8
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   0
      Picture         =   "fifth.frx":0015
      ScaleHeight     =   1035
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SEMESTER 5 MARK DETAILS"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Student RollNo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   240
      TabIndex        =   28
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Average"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   27
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   26
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   720
      TabIndex        =   25
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   24
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   23
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   720
      TabIndex        =   22
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   720
      TabIndex        =   21
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   720
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1200
      TabIndex        =   19
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   360
      TabIndex        =   18
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "fifth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim str As String
Dim cn As New ADODB.Connection

Private Sub cmdclear_Click()
clear
End Sub

Private Sub cmdedit_Click()
rs.Fields("Stu_ID").Value = txtid.Text
rs.Fields("Stu_Name").Value = txtname.Text
rs.Fields("Course").Value = cmbcourse.Text
rs.Fields("Sem1").Value = txtsem1.Text
rs.Fields("Sem2").Value = txtsem2.Text
rs.Fields("Sem3").Value = txtsem3.Text
rs.Fields("Sem4").Value = txtsem4.Text
rs.Fields("Sem5").Value = txtsem5.Text
rs.Fields("Sem6").Value = txtsem6.Text
rs.Fields("Overalltotal").Value = txttotal.Text
rs.Fields("Average").Value = txtaverage.Text
rs.Update
MsgBox "All Semester Mark was Updated ", vbInformation

End Sub

Private Sub cmdexit_Click()
Unload Me
cn.Close
End Sub

Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = txtid.Text
rs.Fields(1) = txtname.Text
rs.Fields(2) = cmbcourse.Text
rs.Fields(3) = txtsem1.Text
rs.Fields(4) = txtsem2.Text
rs.Fields(5) = txtsem3.Text
rs.Fields(6) = txtsem4.Text
rs.Fields(7) = txtsem5.Text
rs.Fields(8) = txtsem6.Text
rs.Fields(9) = txttotal.Text
rs.Fields(10) = txtaverage.Text
rs.Update
MsgBox "All Semester Mark was Updated ", vbInformation
rs.Close
rs.Open "select * from Marks_Details", cn, adOpenDynamic, adLockPessimistic
clear
End Sub

Private Sub Command1_Click()
Dim sum As Integer
txtsem1.Text = Val(txtsem1.Text)
txtsem2.Text = Val(txtsem2.Text)
txtsem3.Text = Val(txtsem3.Text)
txtsem4.Text = Val(txtsem4.Text)
txtsem5.Text = Val(txtsem5.Text)
txtsem6.Text = Val(txtsem6.Text)
sum = Val(txtsem1.Text) + Val(txtsem2.Text) + Val(txtsem3.Text) + Val(txtsem4.Text) + Val(txtsem5.Text) + Val(txtsem6.Text)
txttotal = sum
txtsem1.Text = Val(txtsem1.Text)
txtsem2.Text = Val(txtsem2.Text)
txtsem3.Text = Val(txtsem3.Text)
txtsem4.Text = Val(txtsem4.Text)
txtsem5.Text = Val(txtsem5.Text)
txtsem6.Text = Val(txtsem6.Text)
sum = (Val(txtsem1.Text) + Val(txtsem2.Text) + Val(txtsem3.Text) + Val(txtsem4.Text) + Val(txtsem5.Text) + Val(txtsem6.Text)) / 6
txtaverage = sum

End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
cn.Open
rs.Open "select * from Sem5", cn, adOpenDynamic, adLockPessimistic
Adodc1.Visible = False
End Sub
Sub clear()
txtid.Text = ""
txtname.Text = ""
cmbcourse.Text = ""
txtsem1.Text = "0"
txtsem2.Text = "0"
txtsem3.Text = "0"
txtsem4.Text = "0"
txtsem5.Text = "0"
txtsem6.Text = "0"
txttotal.Text = ""
txtaverage.Text = ""

End Sub
