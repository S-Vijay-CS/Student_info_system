VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H8000000E&
   Caption         =   "Form10"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form10"
   Picture         =   "Hostel Fees.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   5040
      Picture         =   "Hostel Fees.frx":137A7
      ScaleHeight     =   1755
      ScaleWidth      =   5355
      TabIndex        =   22
      Top             =   0
      Width           =   5415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HOSTEL FEES"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   0
      Picture         =   "Hostel Fees.frx":80F50
      ScaleHeight     =   1755
      ScaleWidth      =   10395
      TabIndex        =   21
      Top             =   0
      Width           =   10455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   3960
      Picture         =   "Hostel Fees.frx":887C6
      ScaleHeight     =   2475
      ScaleWidth      =   6435
      TabIndex        =   20
      Top             =   1920
      Width           =   6495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000E&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8880
      Picture         =   "Hostel Fees.frx":B89E5
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7320
      Picture         =   "Hostel Fees.frx":B9062
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5760
      Picture         =   "Hostel Fees.frx":B963F
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      Picture         =   "Hostel Fees.frx":B9C39
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Record Options"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   3960
      TabIndex        =   15
      Top             =   4440
      Width           =   6495
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   13
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   9
      Top             =   2985
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Hostel Fees"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   3735
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees For"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Room No"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         BeginProperty Font 
            Name            =   "Centaur"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Form2.Show
Me.Hide
End Sub

