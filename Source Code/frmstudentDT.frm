VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmstudentDT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   10830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10830
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmstudentDT.frx":0000
      Left            =   2400
      List            =   "frmstudentDT.frx":0025
      TabIndex        =   49
      Text            =   "Combo4"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   48
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox txtcourse 
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
      ItemData        =   "frmstudentDT.frx":008E
      Left            =   2400
      List            =   "frmstudentDT.frx":00B3
      TabIndex        =   47
      Text            =   "Select Your Course"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H0000FF00&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0000FF00&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6000
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7560
      TabIndex        =   44
      Top             =   2280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97910785
      CurrentDate     =   43491
   End
   Begin VB.PictureBox Picture1 
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
      Left            =   4800
      Picture         =   "frmstudentDT.frx":011C
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   4680
      Picture         =   "frmstudentDT.frx":16136
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   42
      Top             =   3120
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
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
      Left            =   4800
      Picture         =   "frmstudentDT.frx":2C150
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   2280
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11400
      Top             =   6240
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
      RecordSource    =   "Stu_regi1"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16800
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      Caption         =   "Frame2"
      Height          =   3735
      Left            =   16080
      TabIndex        =   38
      Top             =   1680
      Width           =   3375
      Begin VB.CommandButton ChangePic 
         BackColor       =   &H0000FF00&
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2415
         Left            =   480
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H0000FF00&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H0000FF00&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   960
      TabIndex        =   34
      Top             =   5640
      Width           =   5055
   End
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
      Left            =   6360
      TabIndex        =   30
      Top             =   6000
      Width           =   855
   End
   Begin VB.Frame Frame 
      BackColor       =   &H80000002&
      Caption         =   "Navigator"
      Height          =   1095
      Left            =   6240
      TabIndex        =   29
      Top             =   5640
      Width           =   3615
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
         Left            =   840
         TabIndex        =   33
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
         Left            =   1680
         TabIndex        =   32
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
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmstudentDT.frx":4216A
      Height          =   3255
      Left            =   360
      TabIndex        =   28
      Top             =   7200
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   5741
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
   Begin VB.TextBox txtadmi 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtphno 
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
      Left            =   12840
      MaxLength       =   10
      TabIndex        =   11
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtmark 
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
      Left            =   7560
      MaxLength       =   4
      TabIndex        =   10
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox txtfees 
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
      Left            =   12840
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtadd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12840
      TabIndex        =   8
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox txtage 
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
      Left            =   7560
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtmname 
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
      Left            =   2400
      TabIndex        =   6
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtfname 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   3840
      Width           =   2895
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
      Height          =   525
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtmail 
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
      Left            =   12840
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
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
      ItemData        =   "frmstudentDT.frx":4217F
      Left            =   7560
      List            =   "frmstudentDT.frx":42189
      TabIndex        =   2
      Text            =   "Select your Gender"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmstudentDT.frx":4219B
      Left            =   7560
      List            =   "frmstudentDT.frx":421AE
      TabIndex        =   1
      Text            =   "Select your Cast"
      Top             =   3840
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "frmstudentDT.frx":421C7
      Left            =   12840
      List            =   "frmstudentDT.frx":421E3
      TabIndex        =   0
      Text            =   "Select your Bloodgroup"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   14
      Left            =   -480
      TabIndex        =   40
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT DETAILS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   6000
      TabIndex        =   27
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Admission NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   -600
      TabIndex        =   26
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mail ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   10560
      TabIndex        =   25
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   2
      Left            =   10560
      TabIndex        =   24
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   10320
      TabIndex        =   23
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   10560
      TabIndex        =   22
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   5
      Left            =   5280
      TabIndex        =   21
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mark"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   5280
      TabIndex        =   20
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Recipt No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   7
      Left            =   10200
      TabIndex        =   19
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   8
      Left            =   5280
      TabIndex        =   18
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   9
      Left            =   5280
      TabIndex        =   17
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "D.O.B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   10
      Left            =   5280
      TabIndex        =   16
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mother Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   11
      Left            =   0
      TabIndex        =   15
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   12
      Left            =   0
      TabIndex        =   14
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   13
      Left            =   -120
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
End
Attribute VB_Name = "frmstudentDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim str As String
Dim cn As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim ssql As String

Private Sub ChangePic_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
str = CommonDialog1.FileName
Image1.Picture = LoadPicture(str)
End Sub

Private Sub cmddelete_Click(Index As Integer)
confirm = MsgBox("Do you want to delete the STUDENT DETAILS", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox " Record has been deleted successfully", vbInformation, "Message"
rs.Update

Else
MsgBox "Record not deleted...!!!", vbInformation, "Message"
End If

End Sub

Private Sub cmdexit_Click(Index As Integer)
cn.Close
Unload Me
End Sub

Private Sub cmdfirst_Click(Index As Integer)
rs.MoveFirst
display
End Sub

Private Sub cmdlast_Click(Index As Integer)
rs.MoveLast
display
End Sub

Private Sub cmdnew_Click()
txtadmi.Text = ""
txtname.Text = ""
txtfname.Text = ""
txtmname.Text = ""
DTPicker1.Value = "01/01/2019"
Combo1.Text = ""
txtage.Text = ""
txtfees.Text = ""
txtmark.Text = ""
txtphno.Text = ""
Combo2.Text = ""
Combo3.Text = ""
txtmail.Text = ""
txtadd.Text = ""
txtcourse.Text = ""
Dim nums As String
Dim numss As Long
Set rs = New ADODB.Recordset
rs.Open "select * from Stu_regi1", cn, adOpenDynamic, adLockPessimistic

With rs
If rs.EOF Then
nums = "0000"
txtadmi.Text = nums
Else
numss = Right(!AdmissionNo, 3) + 1
            nums = Right("0000" & numss, 3)
            End If
            txtadmi.Text = nums
        End With
        Set rs = Nothing

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

Private Sub cmdsave_Click(Index As Integer)
Set rs = New ADODB.Recordset
rs.Open "select * from Stu_regi1", cn, adOpenDynamic, adLockPessimistic
If txtadmi.Text = "" Or txtcourse.Text = "" Or txtname.Text = "" Or txtfname.Text = "" Or _
txtmname.Text = "" Or DTPicker1.Value = "" Or Combo1.Text = "" Or txtage.Text = "" Or _
txtfees.Text = "" Or txtmark.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or txtphno.Text = "" Or _
txtadd.Text = "" Or txtmail.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
End If


With rs
.AddNew

.Fields("Name").Value = txtname.Text
.Fields("FatherName").Value = txtfname.Text
.Fields("MotherName").Value = txtmname.Text
.Fields("DOB").Value = DTPicker1.Value
.Fields("Gender").Value = Combo1.Text
.Fields("Age").Value = txtage.Text
.Fields("Course").Value = txtcourse.Text
.Fields("AdmissionFees").Value = txtfees.Text
.Fields("Mark").Value = txtmark.Text
.Fields("ContactNo").Value = txtphno.Text
.Fields("Cast").Value = Combo2.Text
.Fields("BloodGroup").Value = Combo3.Text
.Fields("MailID").Value = txtmail.Text
.Fields("Address").Value = txtadd.Text
.Fields("PHOTO").Value = str
.Update
End With
MsgBox "SuccessFull..!!!", vbInformation, Admission
Set rs = Nothing
End Sub

Sub display()
txtadmi.Text = rs!AdmissionNo
txtname.Text = rs!Name
txtfname.Text = rs!FatherName
txtmname.Text = rs!MotherName
txtcourse.Text = rs!Course
DTPicker1.Value = rs!DOB
Combo1.Text = rs!Gender
txtage.Text = rs!age
txtfees.Text = rs!AdmissionFees
txtmark.Text = rs!Mark
Combo2.Text = rs!Cast
Combo3.Text = rs!BloodGroup
txtphno.Text = rs!ContactNo
txtmail.Text = rs!MailID
txtadd.Text = rs!address
Image1.Picture = LoadPicture(rs!Photo)
End Sub

Private Sub cmdsave1_Click(Index As Integer)
If txtadmi.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtcourse.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtname.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtfname.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtmname.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf DTPicker1.Value = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf Combo1.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtage.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtfees.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtmark.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf Combo2.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf Combo3.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtphno.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtadd.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
ElseIf txtmail.Text = "" Then
MsgBox "Fill In All  TextBox", vbExclamation
End If


With rs
.AddNew
.Fields("AdmissionNo").Value = txtadmi.Text
.Fields("Name").Value = txtname.Text
.Fields("FatherName").Value = txtfname.Text
.Fields("MotherName").Value = txtmname.Text
.Fields("DOB").Value = DTPicker1.Value
.Fields("Gender").Value = Combo1.Text
.Fields("Age").Value = txtage.Text
.Fields("AdmissionFees").Value = txtfees.Text
.Fields("Mark").Value = txtmark.Text
.Fields("Course").Value = txtcourse.Text
.Fields("ContactNo").Value = txtphno.Text
.Fields("Cast").Value = Combo2.Text
.Fields("BloodGroup").Value = Combo3.Text
.Fields("MailID").Value = txtmail.Text
.Fields("Address").Value = txtadd.Text
.Fields("PHOTO").Value = str
.Update
End With
MsgBox "SuccessFull..!!!", vbInformation, Admission
Set rs = Nothing

End Sub

Private Sub Command1_Click()
ssql = "select * from Stu_regi1 where Course='" & Combo4.Text & "'"
rs1.ActiveConnection = adocon
rs1.Source = ssql
rs1.Open
Set DataReport1.DataSource = rs1
DataReport1.Show
End Sub

Private Sub cmdupdate_Click(Index As Integer)
With rs
.Fields("AdmissionNo").Value = txtadmi.Text
.Fields("Name").Value = txtname.Text
.Fields("FatherName").Value = txtfname.Text
.Fields("MotherName").Value = txtmname.Text

.Fields("DOB").Value = DTPicker1.Value
.Fields("Gender").Value = Combo1.Text
.Fields("Age").Value = txtage.Text
.Fields("Course").Value = txtcourse.Text
.Fields("AdmissionFees").Value = txtfees.Text
.Fields("Mark").Value = txtmark.Text
.Fields("ContactNo").Value = txtphno.Text
.Fields("Cast").Value = Combo2.Text
.Fields("BloodGroup").Value = Combo3.Text
.Fields("MailID").Value = txtmail.Text
.Fields("Address").Value = txtadd.Text
.Fields("PHOTO").Value = str
.Update
End With
MsgBox "SuccessFull..!!!", vbInformation, Admission

Set rs = Nothing

End Sub

Private Sub Command2_Click()
If t = 0 Then
DataGrid1.Visible = True
t = 1
Else
DataGrid1.Visible = False
t = 0
End If

End Sub



Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
cn.Open
rs.Open "select * from Stu_regi1", cn, adOpenDynamic, adLockPessimistic

Adodc1.Visible = False
DataGrid1.Enabled = False
End Sub

Private Sub Picture1_Click()
rs.Close
rs.Open "Select*from Stu_regi1 where AdmissionNo='" + txtadmi.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found..!!!", vbInformation
End If
End Sub

Sub reload()
rs.Close
rs.Open "select*from Stu_regi1", con, adOpenDynamic, adLockPessimistic

End Sub

Private Sub Picture2_Click()
rs.Close
rs.Open "Select*form Stu_regi1 where Name='" + txtname.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found..!!!", vbInformation
End If

End Sub

Private Sub Picture3_Click()
rs.Close
rs.Open "Select*from Stu_regi1 where Course='" + txtcourse.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found..!!!", vbInformation
End If

End Sub


Private Sub txtage_GotFocus()
Dim y As Integer
y = (DateValue(Date) - DateValue(DTPicker1.Value)) / 365
txtage = y
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
Else
KeyAscii = 0
End If

End Sub

Private Sub txtfees_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
Else
KeyAscii = 0
End If

End Sub

Private Sub txtfname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
KeyAscii = 0
Else

End If

End Sub

Private Sub txtmark_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
Else
KeyAscii = 0
End If

End Sub

Private Sub txtmname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
KeyAscii = 0
Else

End If

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
KeyAscii = 0
Else

End If
End Sub

Private Sub txtphno_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
Else
KeyAscii = 0
End If
End Sub
