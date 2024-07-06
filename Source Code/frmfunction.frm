VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13725
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   13725
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8640
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "function_det1"
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
      Bindings        =   "frmfunction.frx":0000
      Height          =   2175
      Left            =   480
      TabIndex        =   24
      Top             =   7080
      Width           =   11655
      _ExtentX        =   20558
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
   Begin VB.CommandButton Command6 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame 
      BackColor       =   &H80000002&
      Caption         =   "Navigator"
      Height          =   975
      Left            =   8400
      TabIndex        =   13
      Top             =   4800
      Width           =   3735
      Begin VB.CommandButton cmdfirst 
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdlast 
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdpre 
         Caption         =   "Pre"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox f6 
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox f5 
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox f4 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox f3 
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   2670
   End
   Begin VB.TextBox f2 
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox f1 
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCTION TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CHIEF GUEST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCTION TYPE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCTION NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCTION NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCTION DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmdfirst_Click(Index As Integer)
If rs.BOF = False Then
f1.Text = 1
MsgBox ("No more Records")
Else
rs.MoveFirst
End If
End Sub

Private Sub cmdlast_Click(Index As Integer)
rs.MoveLast
If rs.EOF = True Then
rs.MovePrevious
ElseIf rs.RecordCount = 0 Then
f1.Text = 0
Else
rs.MoveLast
f1.Text = rs(0)

MsgBox ("This is the Final Record")

End If

rs.MovePrevious

End Sub

Private Sub cmdnext_Click(Index As Integer)

rs.MovePrevious
If f1 < rs.EOF = True Then
f1 = f1 + 1
Else
MsgBox ("This is the Final Record")
End If

End Sub

Private Sub cmdpre_Click(Index As Integer)


rs.MovePrevious
If f1.Text = 1 Then
MsgBox ("This is the 1st Record")
Else
f1.Text = f1 - 1
End If

End Sub

Private Sub Command1_Click()
Dim nums As String
Dim numss As Long
With rs
If rs.EOF Then
nums = "000" + "00"
f1.Text = nums
Else
numss = Right(!f1_no, 3) + 1
            nums = "000" + Right("00" & numss, 3)
            End If
            f1.Text = nums
        End With
        Set rs = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from function_det1", cn, adOpenDynamic, adLockPessimistic
If f1.Text = "" Or f2.Text = "" Or f3.Text = "" Or f4.Text = "" Or f5.Text = "" Or f6.Text = "" Then
MsgBox "Please Fill All Items Else U Can't Add"
End If
With rs
.AddNew
.Fields(0) = f1.Text
.Fields(1) = f2.Text
.Fields(2) = f3.Text
.Fields(3) = f4.Text
.Fields(4) = f5.Text
.Fields(5) = f6.Text
.Update
End With
MsgBox "Records Added"
Set rs = Nothing

End Sub

Private Sub Command4_Click()

If f1.Text = "" Or f2.Text = "" Or f3.Text = "" Or f4.Text = "" Or f5.Text = "" Or f6.Text = "" Then
MsgBox "Please Fill All Items Else U Can't Add"
Else
rs.Fields(0) = f1.Text
rs.Fields(1) = f2.Text
rs.Fields(2) = f3.Text
rs.Fields(3) = f4.Text
rs.Fields(4) = f5.Text
rs.Fields(5) = f6.Text

rs.Update
MsgBox "Records Added"
rs.Close
rs.Open "select * from function_det1", cn, adOpenDynamic, adLockPessimistic
End If

End Sub

Private Sub Command5_Click()
f1.Text = ""
f2.Text = ""
f3.Text = ""
f4.Text = ""
f5.Text = ""
f6.Text = ""

End Sub

Private Sub Command6_Click()
confirm = MsgBox("Do you want to delete the STUDENT DETAILS", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox " Record has been deleted successfully", vbInformation, "Message"
rs.Update

Else
MsgBox "Record not deleted...!!!", vbInformation, "Message"
End If

End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\university Student.mdb;Persist Security Info=False"
cn.Open
rs.Open "select * from function_det1", cn, adOpenDynamic, adLockPessimistic
Adodc1.Visible = False
End Sub
