VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Multiplex System"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13830
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAddNewMovie 
      Caption         =   "&Add New Movie"
   End
   Begin VB.Menu mnuCBooking 
      Caption         =   "Current &Booking"
   End
   Begin VB.Menu mnuAdvBooking 
      Caption         =   "Advance B&ooking"
   End
   Begin VB.Menu mnuCollection 
      Caption         =   "&Collection"
      Begin VB.Menu mnuDaily 
         Caption         =   "&Daily"
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Screen"
      End
   End
   Begin VB.Menu mnuupdate 
      Caption         =   "&Update"
      Begin VB.Menu mnurate 
         Caption         =   "&Rate"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAddNewMovie_Click()
  Load frmNewMovie
  frmNewMovie.Show
End Sub

Private Sub mnuAdvBooking_Click()
    Load frmBookDate
    frmBookDate.Show
End Sub

Private Sub mnuCBooking_Click()
    Load frmcurrentbooking
    frmcurrentbooking.Show
End Sub

Private Sub mnuDaily_Click()
    Load frmDailyDt
    frmDailyDt.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnurate_Click()
    Load frmrate
    frmrate.Show
End Sub

Private Sub mnuScreen_Click()
    If isLoad = True Then
        Unload DataReport3
        Unload DataEnvironment1
    Else
        isLoad = True
    End If


    DataEnvironment1.Connection1.ConnectionString = "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    DataEnvironment1.cmdmonthly
    Unload Me
    Load DataReport3
    DataReport3.Show

End Sub
