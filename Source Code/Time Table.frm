VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form14 
   BackColor       =   &H8000000E&
   Caption         =   "Form14"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   Picture         =   "Time Table.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   11085
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7368
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox t1 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   46
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t2 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   45
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t3 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   44
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t4 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   43
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t5 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t6 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   41
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t7 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   40
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox t8 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   39
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t9 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   38
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t10 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   37
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t11 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   36
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t12 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   35
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t13 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   34
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t14 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   33
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox t15 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   32
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t16 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   31
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t17 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t18 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t19 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   28
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t20 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   27
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t21 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox t22 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t23 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t24 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t25 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t26 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t27 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t28 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox t29 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t30 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t31 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t32 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t33 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t34 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t35 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox t36 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox t37 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox t38 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox t39 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox t40 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox t41 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox t42 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8544
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6192
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5016
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   48
      Top             =   1320
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16761087
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Time Table.frx":1CA29
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "t0"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Line1(0)"
      Tab(0).Control(9)=   "Line2"
      Tab(0).Control(10)=   "Line3"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "STUDENT TIME TABLE"
      TabPicture(1)   =   "Time Table.frx":1CA45
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Line4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label13"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label12"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label9"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "DTPicker2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin MSComCtl2.DTPicker t0 
         Height          =   375
         Left            =   -73920
         TabIndex        =   49
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   32768
         Format          =   99418113
         CurrentDate     =   41556
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1080
         TabIndex        =   50
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   32768
         Format          =   99418113
         CurrentDate     =   41556
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   64
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day I"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   63
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day II"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   62
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day III"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   61
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day IV"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   60
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day V"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   59
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day VI"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   -74880
         TabIndex        =   58
         Top             =   3960
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -71280
         X2              =   -71280
         Y1              =   1440
         Y2              =   4440
      End
      Begin VB.Line Line2 
         X1              =   -68520
         X2              =   -68520
         Y1              =   1440
         Y2              =   4440
      End
      Begin VB.Line Line3 
         X1              =   -68280
         X2              =   -68280
         Y1              =   1440
         Y2              =   4440
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   0
         TabIndex        =   57
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day I"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day II"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day III"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day IV"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day V"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day VI"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   3960
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3840
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3840
      End
      Begin VB.Line Line5 
         X1              =   3720
         X2              =   3720
         Y1              =   1440
         Y2              =   4680
      End
      Begin VB.Line Line6 
         X1              =   6480
         X2              =   6480
         Y1              =   1440
         Y2              =   4560
      End
      Begin VB.Line Line7 
         X1              =   6720
         X2              =   6720
         Y1              =   1440
         Y2              =   4680
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdcancel_Click()

rs.MoveFirst
rs.CancelUpdate
MsgBox "Update Cancelled"

cmdclear.Enabled = True
cmdedit.Enabled = True
cmdupdate.Enabled = False
cmdclose.Enabled = True
cmdcancel.Enabled = False

cmdsave.Enabled = False

t1.Enabled = False
t2.Enabled = False
t3.Enabled = False
t4.Enabled = False
t5.Enabled = False
t6.Enabled = False
t7.Enabled = False
t8.Enabled = False
t9.Enabled = False
t10.Enabled = False
t11.Enabled = False
t12.Enabled = False
t13.Enabled = False
t14.Enabled = False
t15.Enabled = False
t16.Enabled = False
t17.Enabled = False
t18.Enabled = False
t19.Enabled = False
t20.Enabled = False
t21.Enabled = False
t22.Enabled = False
t23.Enabled = False
t24.Enabled = False
t25.Enabled = False
t26.Enabled = False
t27.Enabled = False
t28.Enabled = False
t29.Enabled = False
t30.Enabled = False
t31.Enabled = False
t32.Enabled = False
t33.Enabled = False
t34.Enabled = False
t35.Enabled = False
t36.Enabled = False
t37.Enabled = False
t38.Enabled = False
t39.Enabled = False
t40.Enabled = False
t41.Enabled = False
t42.Enabled = False

End Sub

Private Sub cmdclear_Click()

cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdclose.Enabled = False

cmdcancel.Enabled = True
cmdsave.Enabled = True

t1.Enabled = True
t2.Enabled = True
t3.Enabled = True
t4.Enabled = True
t5.Enabled = True
t6.Enabled = True
t7.Enabled = True
t8.Enabled = True
t9.Enabled = True
t10.Enabled = True
t11.Enabled = True
t12.Enabled = True
t13.Enabled = True
t14.Enabled = True
t15.Enabled = True
t16.Enabled = True
t17.Enabled = True
t18.Enabled = True
t19.Enabled = True
t20.Enabled = True
t21.Enabled = True
t22.Enabled = True
t23.Enabled = True
t24.Enabled = True
t25.Enabled = True
t26.Enabled = True
t27.Enabled = True
t28.Enabled = True
t29.Enabled = True
t30.Enabled = True
t31.Enabled = True
t32.Enabled = True
t33.Enabled = True
t34.Enabled = True
t35.Enabled = True
t36.Enabled = True
t37.Enabled = True
t38.Enabled = True
t39.Enabled = True
t40.Enabled = True
t41.Enabled = True
t42.Enabled = True

End Sub

Private Sub cmdclose_Click()
Form1.Show
Me.Hide
End Sub

Private Sub cmdedit_Click()

cmdedit.Enabled = False
cmdclose.Enabled = False
cmdclear.Enabled = False

cmdsave.Enabled = False

cmdupdate.Enabled = True
cmdcancel.Enabled = True

t1.Enabled = True
t2.Enabled = True
t3.Enabled = True
t4.Enabled = True
t5.Enabled = True
t6.Enabled = True
t7.Enabled = True
t8.Enabled = True
t9.Enabled = True
t10.Enabled = True
t11.Enabled = True
t12.Enabled = True
t13.Enabled = True
t14.Enabled = True
t15.Enabled = True
t16.Enabled = True
t17.Enabled = True
t18.Enabled = True
t19.Enabled = True
t20.Enabled = True
t21.Enabled = True
t22.Enabled = True
t23.Enabled = True
t24.Enabled = True
t25.Enabled = True
t26.Enabled = True
t27.Enabled = True
t28.Enabled = True
t29.Enabled = True
t30.Enabled = True
t31.Enabled = True
t32.Enabled = True
t33.Enabled = True
t34.Enabled = True
t35.Enabled = True
t36.Enabled = True
t37.Enabled = True
t38.Enabled = True
t39.Enabled = True
t40.Enabled = True
t41.Enabled = True
t42.Enabled = True

End Sub

Private Sub cmdsave_Click()

cmdclear.Enabled = True
cmdedit.Enabled = False
cmdclose.Enabled = False

cmdsave.Enabled = False
cmdcancel.Enabled = False

t1.Enabled = True
t2.Enabled = True
t3.Enabled = True
t4.Enabled = True
t5.Enabled = True
t6.Enabled = True
t7.Enabled = True
t8.Enabled = True
t9.Enabled = True
t10.Enabled = True
t11.Enabled = True
t12.Enabled = True
t13.Enabled = True
t14.Enabled = True
t15.Enabled = True
t16.Enabled = True
t17.Enabled = True
t18.Enabled = True
t19.Enabled = True
t20.Enabled = True
t21.Enabled = True
t22.Enabled = True
t23.Enabled = True
t24.Enabled = True
t25.Enabled = True
t26.Enabled = True
t27.Enabled = True
t28.Enabled = True
t29.Enabled = True
t30.Enabled = True
t31.Enabled = True
t32.Enabled = True
t33.Enabled = True
t34.Enabled = True
t35.Enabled = True
t36.Enabled = True
t37.Enabled = True
t38.Enabled = True
t39.Enabled = True
t40.Enabled = True
t41.Enabled = True
t42.Enabled = True

rs.AddNew
rs.Fields(0) = t1.Text
rs.Fields(1) = t2.Text
rs.Fields(2) = t3.Text
rs.Fields(3) = t4.Text
rs.Fields(4) = t5.Text
rs.Fields(5) = t6.Text
rs.Fields(6) = t7.Text
rs.Fields(7) = t8.Text
rs.Fields(8) = t9.Text
rs.Fields(9) = t10.Text
rs.Fields(10) = t11.Text
rs.Fields(11) = t12.Text
rs.Fields(12) = t13.Text
rs.Fields(13) = t14.Text
rs.Fields(14) = t15.Text
rs.Fields(15) = t16.Text
rs.Fields(16) = t17.Text
rs.Fields(17) = t18.Text
rs.Fields(18) = t19.Text
rs.Fields(19) = t20.Text
rs.Fields(20) = t21.Text
rs.Fields(21) = t22.Text
rs.Fields(22) = t23.Text
rs.Fields(23) = t24.Text
rs.Fields(24) = t25.Text
rs.Fields(25) = t26.Text
rs.Fields(26) = t27.Text
rs.Fields(27) = t28.Text
rs.Fields(28) = t29.Text
rs.Fields(29) = t30.Text
rs.Fields(30) = t31.Text
rs.Fields(31) = t32.Text
rs.Fields(32) = t33.Text
rs.Fields(33) = t34.Text
rs.Fields(34) = t35.Text
rs.Fields(35) = t36.Text
rs.Fields(36) = t37.Text
rs.Fields(37) = t38.Text
rs.Fields(38) = t39.Text
rs.Fields(39) = t40.Text
rs.Fields(40) = t41.Text
rs.Fields(41) = t42.Text
rs.Fields(42) = t0.Value
rs.Update
MsgBox "Records Added"
rs.Close
rs.Open "select * from Time_Table", cn, adOpenDynamic, adLockPessimistic

cmdclear.Enabled = True
cmdedit.Enabled = True
cmdclose.Enabled = True

rs.Update

't0.Enabled = False
t1.Enabled = False
t2.Enabled = False
t3.Enabled = False
t4.Enabled = False
t5.Enabled = False
t6.Enabled = False
t7.Enabled = False
t8.Enabled = False
t9.Enabled = False
t10.Enabled = False
t11.Enabled = False
t12.Enabled = False
t13.Enabled = False
t14.Enabled = False
t15.Enabled = False
t16.Enabled = False
t17.Enabled = False
t18.Enabled = False
t19.Enabled = False
t20.Enabled = False
t21.Enabled = False
t22.Enabled = False
t23.Enabled = False
t24.Enabled = False
t25.Enabled = False
t26.Enabled = False
t27.Enabled = False
t28.Enabled = False
t29.Enabled = False
t30.Enabled = False
t31.Enabled = False
t32.Enabled = False
t33.Enabled = False
t34.Enabled = False
t35.Enabled = False
t36.Enabled = False
t37.Enabled = False
t38.Enabled = False
t39.Enabled = False
t40.Enabled = False
t41.Enabled = False
t42.Enabled = False

End Sub

Private Sub cmdupdate_Click()

cn.Execute "update Time_Table set tt1='" + t1.Text + "',tt2='" + t2.Text + "',tt3='" + t3.Text + "',ta4='" + t4.Text + "',tt5='" + t5.Text + "',tt6='" + t6.Text + "',tt7='" + t7.Text + "',tt8='" + t8.Text + "',tt9='" + t9.Text + "',tt10='" + t10.Text + "',tt11='" + t11.Text + "',tt12='" + t12.Text + "',tt13='" + t13.Text + "',tt14='" + t14.Text + "',tt15='" + t15.Text + "', tt16='" + t16.Text + "', tt17='" + t17.Text + "', tt18='" + t18.Text + "', tt19='" + t19.Text + "',tt20='" + t20.Text + "',tt21='" + t21.Text + "', tt22='" + t22.Text + "',tt23='" + t23.Text + "',tt24='" + t24.Text + "',tt25='" + t25.Text + "',tt + t26.Text + " ',tt27='" + t27.Text + "',tt28='" + t28.Text + "',tt29='" + t29.Text + "',tt30='" + t30.Text + "',tt31='" + t31.Text + "',tt32='" + t32.Text + "',tt33='" + t33.Text + "',tt34='" + t34.Text + "',tt35='" + t35.Text + "',tt36='" + t36.Text + "',tt37='" + t37.Text + "',tt38='" + t38.Text + "',tt39='" + t39.Text + "',tt40='" + t40.Text + "',tt4='" + t41.Text + "',tt42='" + t42.Text + "'
MsgBox "Updated"
    
rs.Close
rs.Open "select * from Time_Table", cn, adOpenDynamic, adLockPessimistic

cmdedit.Enabled = False
cmdupdate.Enabled = False
cmdclose.Enabled = False
cmdclear.Enabled = False
cmdsave.Enabled = False

cmdcancel.Enabled = True

t1.Enabled = True
t2.Enabled = True
t3.Enabled = True
t4.Enabled = True
t5.Enabled = True
t6.Enabled = True
t7.Enabled = True
t8.Enabled = True
t9.Enabled = True
t10.Enabled = True
t11.Enabled = True
t12.Enabled = True
t13.Enabled = True
t14.Enabled = True
t15.Enabled = True
t16.Enabled = True
t17.Enabled = True
t18.Enabled = True
t19.Enabled = True
t20.Enabled = True
t21.Enabled = True
t22.Enabled = True
t23.Enabled = True
t24.Enabled = True
t25.Enabled = True
t26.Enabled = True
t27.Enabled = True
t28.Enabled = True
t29.Enabled = True
t30.Enabled = True
t31.Enabled = True
t32.Enabled = True
t33.Enabled = True
t34.Enabled = True
t35.Enabled = True
t36.Enabled = True
t37.Enabled = True
t38.Enabled = True
t39.Enabled = True
t40.Enabled = True
t41.Enabled = True
t42.Enabled = True

End Sub

Private Sub Form_Load()

Me.Top = 250
Me.Left = 1200

cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
cn.Open
rs.Open "select * from Time_Table", cn, adOpenDynamic, adLockPessimistic

cmdsave.Enabled = False
cmdupdate.Enabled = False
cmdcancel.Enabled = False

 'rs.MoveFirst
 'If t1.Text = rs(0) Then
    't2.Text = rs(1)
    't3.Text = rs(2)
    't4.Text = rs(3)
    't5.Text = rs(4)
    't6.Text = rs(5)
    't7.Text = rs(6)
    't8.Text = rs(7)
    't9.Text = rs(8)
    't10.Text = rs(9)
    't11.Text = rs(10)
    't12.Text = rs(11)
    't13.Text = rs(12)
    't14.Text = rs(13)
    't15.Text = rs(14)
    't16.Text = rs(15)
    't17.Text = rs(16)
    't18.Text = rs(17)
    't19.Text = rs(18)
    't20.Text = rs(19)
    't21.Text = rs(10)
    't22.Text = rs(21)
    't23.Text = rs(22)
    't24.Text = rs(23)
    't25.Text = rs(24)
    't26.Text = rs(25)
    't27.Text = rs(26)
    't28.Text = rs(27)
    't29.Text = rs(28)
    't30.Text = rs(29)
    't31.Text = rs(30)
    't32.Text = rs(31)
    't33.Text = rs(32)
    't34.Text = rs(33)
    't35.Text = rs(34)
    't36.Text = rs(35)
    't37.Text = rs(36)
    't38.Text = rs(37)
    't39.Text = rs(38)
    't40.Text = rs(39)
    't41.Text = rs(40)
    't42.Text = rs(41)
    't0.Value = rs(42)
    'End If
    'rs.MoveNext

t0.Enabled = False
t1.Enabled = False
t2.Enabled = False
t3.Enabled = False
t4.Enabled = False
t5.Enabled = False
t6.Enabled = False
t7.Enabled = False
t8.Enabled = False
t9.Enabled = False
t10.Enabled = False
t11.Enabled = False
t12.Enabled = False
t13.Enabled = False
t14.Enabled = False
t15.Enabled = False
t16.Enabled = False
t17.Enabled = False
t18.Enabled = False
t19.Enabled = False
t20.Enabled = False
t21.Enabled = False
t22.Enabled = False
t23.Enabled = False
t24.Enabled = False
t25.Enabled = False
t26.Enabled = False
t27.Enabled = False
t28.Enabled = False
t29.Enabled = False
t30.Enabled = False
t31.Enabled = False
t32.Enabled = False
t33.Enabled = False
t34.Enabled = False
t35.Enabled = False
t36.Enabled = False
t37.Enabled = False
t38.Enabled = False
t39.Enabled = False
t40.Enabled = False
t41.Enabled = False
t42.Enabled = False

End Sub
