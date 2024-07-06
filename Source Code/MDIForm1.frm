VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5700
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnumagement 
      Caption         =   "Management"
      Begin VB.Menu mnuregistration 
         Caption         =   "Student Registration"
      End
      Begin VB.Menu mnufunction 
         Caption         =   "Function Details"
      End
      Begin VB.Menu mnucourse 
         Caption         =   "Course Details"
      End
   End
   Begin VB.Menu mnumark 
      Caption         =   "Mark Entry"
   End
   Begin VB.Menu mnuattendance 
      Caption         =   "Attendance Details"
   End
   Begin VB.Menu mnufees 
      Caption         =   "Fees Details"
   End
   Begin VB.Menu mnutime 
      Caption         =   "Time Table"
      Visible         =   0   'False
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu mnusreport 
         Caption         =   "Student Report"
      End
      Begin VB.Menu Mreport 
         Caption         =   "Mark Report"
      End
      Begin VB.Menu Freport 
         Caption         =   "Fees Report"
      End
      Begin VB.Menu Areport 
         Caption         =   "Attendance Report"
      End
   End
   Begin VB.Menu mnulogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Areport_Click()
DataReport4.Show
End Sub

Private Sub Freport_Click()
DataReport3.Show
End Sub

Private Sub mnuattendance_Click()
frmattendance.Show
End Sub

Private Sub mnucourse_Click()
Form8.Show
End Sub

Private Sub mnufees_Click()
Form11.Show
End Sub

Private Sub mnufunction_Click()
Form12.Show
End Sub

Private Sub mnulogout_Click()
frmlogin.Show
Unload Me
End Sub

Private Sub mnumark_Click()
Form2.Show
End Sub

Private Sub mnuregistration_Click()
frmstudentDT.Show
End Sub

Private Sub mnusreport_Click()
DataReport1.Show
End Sub

Private Sub mnutime_Click()
Form14.Show

End Sub

Private Sub Mreport_Click()
DataReport2.Show
End Sub
