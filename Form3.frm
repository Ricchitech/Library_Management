VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "HOME"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "Form3.frx":2BC78
   ScaleHeight     =   10635
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnustud 
      Caption         =   "Students data"
      Begin VB.Menu mnunew 
         Caption         =   "New Student  Data"
      End
      Begin VB.Menu mnumod 
         Caption         =   "Modify Student Data"
      End
   End
   Begin VB.Menu mnubooks 
      Caption         =   "Books"
      Begin VB.Menu menuaddnew 
         Caption         =   "New Books"
      End
      Begin VB.Menu menuupdate 
         Caption         =   "Update Books"
      End
      Begin VB.Menu menurep 
         Caption         =   "Books Report"
      End
   End
   Begin VB.Menu mnuledger 
      Caption         =   "Ledger"
      Begin VB.Menu mnuissue 
         Caption         =   "Book Issue"
      End
      Begin VB.Menu mnuret 
         Caption         =   "Return"
      End
   End
   Begin VB.Menu mnuclose 
      Caption         =   ">>"
      Begin VB.Menu mnuaddnewuser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu menulg 
         Caption         =   "Logout"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub menuaddnew_Click()
Form6.Show
Unload Me
End Sub

Private Sub menulg_Click()
Form2.Show
Unload Me
End Sub

Private Sub menurep_Click()
Form8.Show
Unload Me
End Sub

Private Sub menuupdate_Click()
Form7.Show
Unload Me
End Sub

Private Sub mexit_Click()
End
End Sub

Private Sub mnuaddnewuser_Click()
Form10.Show
End Sub

Private Sub mnuissue_Click()
Form2.Show
Unload Me
End Sub

Private Sub mnumod_Click()
Form5.Show
Unload Me
End Sub

Private Sub mnunew_Click()
Form4.Show
Unload Me
End Sub

Private Sub mnuret_Click()
Form11.Show
Unload Me
End Sub
