VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Siswa"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   10440
      TabIndex        =   11
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   10440
      TabIndex        =   10
      Top             =   5760
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   10440
      TabIndex        =   9
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   10440
      TabIndex        =   8
      Top             =   4560
      Width           =   2535
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
      Left            =   6240
      TabIndex        =   7
      Top             =   6840
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
      Left            =   4800
      TabIndex        =   6
      Top             =   5520
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
      Left            =   6240
      TabIndex        =   5
      Top             =   6000
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
      Left            =   6240
      TabIndex        =   4
      Top             =   5160
      Width           =   975
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
      Left            =   4800
      TabIndex        =   3
      Top             =   6360
      Width           =   975
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
      Left            =   11760
      TabIndex        =   2
      Top             =   7680
      Width           =   1335
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
      Left            =   10080
      TabIndex        =   1
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   10440
      TabIndex        =   0
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No Telepon"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8520
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Siswa"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Siswa"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   6960
      Width           =   1815
   End
End
Attribute VB_Name = "Form8"
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
End Sub

Private Sub brnBack_Click()
Form5.Show
Unload Me
End Sub

Private Sub btnCari_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from siswa where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data telah ditemukan", vbInformation
        
         Text2.Text = Trim(rsTab!nama_siswa)
        Text3.Text = Trim(rsTab!alamat_siswa)
        Text4.Text = Trim(rsTab!telepone)
        Text5.Text = Trim(rsTab!jenis_kelamin)
        
    Else
        MsgBox "Data tidak ditemukan", vbInformation
        Call bersih
    End If
    rsTab.Close
    Set rsTab = Nothing
End Sub

Private Sub btnHapus_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from siswa where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
    rsTab.Delete
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan ID = " & Text1.Text & "Telah dihapus", vbInformation
End Sub

Private Sub btnTambah_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from siswa where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data dengan nomer ID tersebut sudah ada"
        rsTab.Close
        Set rsTab = Nothing
        Exit Sub
    Else
        Set rsTab = New ADODB.Recordset
        rsTab.Open "select * from siswa where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
        rsTab.AddNew
        rsTab!id = Trim(Text1.Text)
        rsTab!nama_siswa = Trim(Text2.Text)
        rsTab!alamat_siswa = Trim(Text3.Text)
        rsTab!telepone = Trim(Text4.Text)
        rsTab!jenis_kelamin = Trim(Text5.Text)
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
    rsTab.Open "select * from siswa where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
     rsTab!id = Trim(Text1.Text)
        rsTab!nama_siswa = Trim(Text2.Text)
        rsTab!alamat_siswa = Trim(Text3.Text)
        rsTab!telepone = Trim(Text4.Text)
        rsTab!jenis_kelamin = Trim(Text5.Text)
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



