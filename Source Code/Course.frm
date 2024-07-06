VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H8000000E&
   Caption         =   "Form8"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   Picture         =   "Course.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   6270
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   4440
      Picture         =   "Course.frx":7C4C
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   0
      Picture         =   "Course.frx":840C
      ScaleHeight     =   1395
      ScaleWidth      =   5835
      TabIndex        =   20
      Top             =   0
      Width           =   5895
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "COURSE DETAILS"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00004000&
      Caption         =   "PG Courses"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5055
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   2895
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc Computer Science"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc Bio Technology"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc Micro Biology"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc Chemistry"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc Maths"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Com CA"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "MCA"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "M.A English"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "M.A Tamil"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "UG Courses"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Sc Computer Science"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Sc Bio Technology"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Sc Micro Biology"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Sc Chemistry"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Sc Maths"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Com CA"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "BCA"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "B.A English"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "B.A Tamil"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture3_Click()
MDIForm1.Show
MDIForm1.Admin.Enabled = True
MDIForm1.HOD.Enabled = False
MDIForm1.Lecturer.Enabled = False
MDIForm1.Library.Enabled = False
MDIForm1.Lab.Enabled = False
MDIForm1.Hostel.Enabled = False
MDIForm1.Office.Enabled = False
MDIForm1.Pass.Enabled = False
MDIForm1.Exit.Enabled = True
Me.Hide
End Sub
