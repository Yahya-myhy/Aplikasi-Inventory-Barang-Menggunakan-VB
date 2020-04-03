VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form Master"
   ClientHeight    =   6255
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10350
   LinkTopic       =   "Data Film"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu admin 
      Caption         =   "Admin"
   End
   Begin VB.Menu guru 
      Caption         =   "Guru"
   End
   Begin VB.Menu siswa 
      Caption         =   "Siswa"
   End
   Begin VB.Menu tlogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub admin_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command1_Click()

End Sub





Private Sub guru_Click()
Form19.Show
Unload Me
End Sub

Private Sub siswa_Click()
Form18.Show
Unload Me
End Sub

Private Sub tlogout_Click()
Form1.Show
Unload Me

End Sub

Private Sub tpemesanan_Click()
DataReport1.Show
Unload Me
End Sub
