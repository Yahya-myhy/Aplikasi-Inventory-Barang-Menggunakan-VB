VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14310
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":15162
   ScaleHeight     =   8985
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   450
      Left            =   120
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   1335
      Left            =   9360
      TabIndex        =   5
      Top             =   7680
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":36085
      Left            =   10680
      List            =   "Form1.frx":36092
      TabIndex        =   4
      Top             =   6120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   10680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistem Inventory Barang"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   15495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9240
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTemp, LenTemp, n

Private Sub Command1_Click()
If Combo1.Text = "Admin" And Text2.Text = "admin" Then
Form13.Show
Form1.Hide


ElseIf Combo1.Text = "Kasir" And Text2.Text = "kasir" Then
Form14.Show
Form1.Hide


ElseIf Combo1.Text = "Pimpinan" And Text2.Text = "pimpinan" Then
Form2.Show
Form1.Hide


 ElseIf Text2.Text = "" Then
    MsgBox "Anda belum memasukkan password", vbCritical, "Salah"
    
    Text2.Text = ""

Else
 MsgBox "Login anda salah", vbCritical, "Salah"
 End If
 
Text2.Text = ""
End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_Load()
strTemp = Label3
    n = 1
End Sub

Private Sub Timer1_Timer()
LenTemp = Len(strTemp)
    Dim Form As String
    LenTemp = Len(strTemp)
    Label3 = Left(strTemp, n) + "_"
    n = n + 1
    If n > LenTemp Then
        n = 1
    End If
End Sub
