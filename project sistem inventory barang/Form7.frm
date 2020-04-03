VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "Pengaturan Admin"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   11640
      TabIndex        =   10
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   11640
      TabIndex        =   9
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   11640
      TabIndex        =   8
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   11640
      TabIndex        =   7
      Top             =   4320
      Width           =   2775
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
      Left            =   7560
      TabIndex        =   6
      Top             =   6120
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
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   5880
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
      Left            =   7560
      TabIndex        =   4
      Top             =   5280
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
      Left            =   7560
      TabIndex        =   3
      Top             =   4440
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
      Height          =   615
      Left            =   6120
      Picture         =   "Form7.frx":52861
      TabIndex        =   2
      Top             =   5040
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
      Height          =   615
      Left            =   13080
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
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
      Height          =   375
      Left            =   11760
      TabIndex        =   0
      Top             =   8400
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6480
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=OraOLEDB.Oracle.1;Password=system;Persist Security Info=True;User ID=system"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Password=system;Persist Security Info=True;User ID=system"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "siswa"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Admin"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10320
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No_Telepone"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10320
      TabIndex        =   13
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat Admin"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10320
      TabIndex        =   12
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Admin"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10320
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "Form7"
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

End Sub

Private Sub brnBack_Click()
Form3.Show
Unload Me
End Sub

Private Sub btnCari_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from admin where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data telah ditemukan", vbInformation
        Text2.Text = Trim(rsTab!nama)
        Text3.Text = Trim(rsTab!alamat)
        Text4.Text = Trim(rsTab!no_hp)
    Else
        MsgBox "Data tidak ditemukan", vbInformation
        Call bersih
    End If
    rsTab.Close
    Set rsTab = Nothing
End Sub

Private Sub btnHapus_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from admin where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
    rsTab.Delete
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan Kode = " & Text1.Text & "Telah dihapus", vbInformation
End Sub

Private Sub btnTambah_Click()
    Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from admin where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data dengan kode tersebut sudah ada"
        rsTab.Close
        Set rsTab = Nothing
        Exit Sub
    Else
        Set rsTab = New ADODB.Recordset
        rsTab.Open "select * from admin where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
        rsTab.AddNew
        rsTab!id = Trim(Text1.Text)
        rsTab!nama = Trim(Text2.Text)
        rsTab!alamat = Trim(Text3.Text)
        rsTab!no_hp = Trim(Text4.Text)
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
    rsTab.Open "select * from admin where id = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
    rsTab!id = Trim(Text1.Text)
    rsTab!nama = Trim(Text2.Text)
    rsTab!alamat = Trim(Text3.Text)
    rsTab!no_hp = Trim(Text4.Text)
    rsTab.Update
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan kode = " & Text1.Text & "Telah Berhasil di Update", vbInformation
    Call bersih
End Sub


Private Sub Command1_Click()
Form2.Show
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



