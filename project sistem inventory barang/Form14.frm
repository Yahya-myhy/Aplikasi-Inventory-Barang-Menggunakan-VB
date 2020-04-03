VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Form Master"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form14"
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu jadwal 
      Caption         =   "Jadwal"
   End
   Begin VB.Menu bh 
      Caption         =   "Hasil Nilai Siswa"
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bh_Click()
Form12.Show
Unload Me
End Sub

Private Sub jadwal_Click()
Form15.Show
Unload Me
End Sub

Private Sub logout_Click()
Form1.Show
Unload Me
End Sub
