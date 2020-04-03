VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form Master"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13890
   LinkTopic       =   "Form13"
   Picture         =   "Form13.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu guru 
      Caption         =   "Guru"
   End
   Begin VB.Menu siswa 
      Caption         =   "Siswa"
   End
   Begin VB.Menu jadwal 
      Caption         =   "Jadwal"
   End
   Begin VB.Menu djd 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub djd_Click()
Form1.Show
Unload Me
End Sub

Private Sub guru_Click()
Form4.Show
Unload Me
End Sub



Private Sub jadwal_Click()
Form9.Show
Unload Me
End Sub

Private Sub siswa_Click()
Form5.Show
Unload Me
End Sub
