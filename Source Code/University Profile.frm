VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   Caption         =   "Form3"
   ClientHeight    =   7350
   ClientLeft      =   2160
   ClientTop       =   1395
   ClientWidth     =   8475
   LinkTopic       =   "Form3"
   Picture         =   "University Profile.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   8475
   Begin VB.CommandButton Command1 
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   30
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5160
      TabIndex        =   29
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5160
      TabIndex        =   28
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&EDIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5160
      TabIndex        =   27
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5160
      TabIndex        =   26
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000080FF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5775
      Left            =   5040
      TabIndex        =   23
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox text11 
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox text10 
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox text9 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox text8 
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox text7 
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Students Count"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5895
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Com"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc Maths"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Sc - Computer Science"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "B.SC Maths"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Sc - Computer Science"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "M.A English"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "M.A Tamil"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "BBA"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "B.Com"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "B.A - English"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "B.A - Tamil"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "M.Com"
      Height          =   495
      Left            =   1560
      TabIndex        =   24
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UNIVERSITY DETAILS"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim str As String
Dim cn As New ADODB.Connection

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
text7.Text = ""
text8.Text = ""
text9.Text = ""
text10.Text = ""
text11.Text = ""
End Sub

Private Sub Command4_Click()
Form2.Show
Me.Hide
End Sub

Private Sub Command5_Click()
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Text6.Text
rs.Fields(6) = text7.Text
rs.Fields(7) = text8.Text
rs.Fields(8) = text9.Text
rs.Fields(9) = text0.Text
rs.Fields(10) = text10.Text
rs.Update
MsgBox "Successfully saved", vbInformation
rs.Close
rs.Open "select * from UniversityTBL", cn, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Command6_Click()
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Text6.Text
rs.Fields(6) = text7.Text
rs.Fields(7) = text8.Text
rs.Fields(8) = text9.Text
rs.Fields(9) = text0.Text
rs.Fields(10) = text10.Text
rs.Update
MsgBox "Successfully saved", vbInformation
rs.Close
rs.Open "select * from UniversityTBL", cn, adOpenDynamic, adLockPessimistic

End Sub

Private Sub Command7_Click()
confirm = MsgBox("Do you want to delete the STUDENT DETAILS", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox " Record has been deleted successfully", vbInformation, "Message"
rs.Update

Else
MsgBox "Record not deleted...!!!", vbInformation, "Message"
End If

End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\University Student.mdb;Persist Security Info=False"
cn.Open
rs.Open "select * from UniversityTBL", cn, adOpenDynamic, adLockPessimistic

End Sub

