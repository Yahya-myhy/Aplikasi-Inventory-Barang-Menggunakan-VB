VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Jadwal"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15435
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   9630
   ScaleWidth      =   15435
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Height          =   765
      Left            =   8520
      TabIndex        =   22
      Top             =   9000
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Height          =   765
      Left            =   8520
      TabIndex        =   19
      Top             =   8040
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   765
      Left            =   8520
      TabIndex        =   18
      Top             =   7080
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   765
      Left            =   8520
      TabIndex        =   11
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ke Form Master"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton brnBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton btnCari 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton btnUbah 
      Caption         =   "Ubah"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   8520
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   1005
      Left            =   8520
      TabIndex        =   2
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   8520
      TabIndex        =   1
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   765
      Left            =   8520
      TabIndex        =   0
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas X11 IPS"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   23
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas X1 IPS"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas X11 IPA"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PENJADWALAN"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   7080
      TabIndex        =   17
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas X1 IPA"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Jadwal"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas X"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connOra As ADODB.Connection
Dim rsTab As ADODB.Recordset

Sub bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub brnBack_Click()
Form9.Show
Unload Me
End Sub

Private Sub btnCari_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from jadwal where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data telah ditemukan", vbInformation
        
         Text2.Text = Trim(rsTab!hari_tanggal)
        Text3.Text = Trim(rsTab!jam)
        Text4.Text = Trim(rsTab!kelas_x)
        Text5.Text = Trim(rsTab!kelas_x1ipa)
        Text6.Text = Trim(rsTab!kelas_x1ips)
        Text7.Text = Trim(rsTab!kelas_x11ipa)
        Text8.Text = Trim(rsTab!kelas_x11ips)
    
    Else
        MsgBox "Data tidak ditemukan", vbInformation
        Call bersih
    End If
    rsTab.Close
    Set rsTab = Nothing
End Sub

Private Sub btnHapus_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from jadwal where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
    rsTab.Delete
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan ID = " & Text1.Text & "Telah dihapus", vbInformation
End Sub

Private Sub btnTambah_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from jadwal where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data dengan nomer ID tersebut sudah ada"
        rsTab.Close
        Set rsTab = Nothing
        Exit Sub
    Else
        Set rsTab = New ADODB.Recordset
        rsTab.Open "select * from jadwal where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
        rsTab.AddNew
        rsTab!id = Trim(Text1.Text)
        rsTab!hari_tanggal = Trim(Text2.Text)
        rsTab!jam = Trim(Text3.Text)
        rsTab!kelas_x = Trim(Text4.Text)
        rsTab!kelas_x1ipa = Trim(Text5.Text)
        rsTab!kelas_x1ips = Trim(Text6.Text)
        rsTab!kelas_x11ipa = Trim(Text7.Text)
        rsTab!kelas_x11ips = Trim(Text8.Text)
        rsTab.Update
        connOra.Execute "Commit"
        rsTab.Close
        Set rsTab = Nothing
        MsgBox "Data telah ditambahkan"
        Call bersih
    End If
End Sub

Private Sub btnUbah_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from jadwal where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
      rsTab!id = Trim(Text1.Text)
        rsTab!hari_tanggal = Trim(Text2.Text)
        rsTab!jam = Trim(Text3.Text)
        rsTab!kelas_x = Trim(Text4.Text)
        rsTab!kelas_x1ipa = Trim(Text5.Text)
        rsTab!kelas_x1ips = Trim(Text6.Text)
        rsTab!kelas_x11ipa = Trim(Text7.Text)
        rsTab!kelas_x11ips = Trim(Text8.Text)
    rsTab.Update
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan ID = " & Text1.Text & "Telah Berhasil di Update", vbInformation
    Call bersih
End Sub


Private Sub Command1_Click()
Form13.Show
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Set connOra = New ADODB.Connection
connOra.Open "Provider=OraOLEDB.Oracle.1;Password=system;Persist Security Info=True;User ID=system"
connOra.CursorLocation = adUseClient
Call bersih
End Sub




