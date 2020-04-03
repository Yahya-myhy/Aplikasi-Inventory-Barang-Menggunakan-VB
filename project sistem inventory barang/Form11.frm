VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15855
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Nilai Siswa"
      TabPicture(0)   =   "Form11.frx":4775C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data Siswa"
      TabPicture(1)   =   "Form11.frx":47778
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Data Guru"
      TabPicture(2)   =   "Form11.frx":47794
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture4 
         Height          =   6615
         Left            =   600
         Picture         =   "Form11.frx":477B0
         ScaleHeight     =   6555
         ScaleWidth      =   11715
         TabIndex        =   4
         Top             =   420
         Width           =   11775
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Form11.frx":6346A
            Left            =   3720
            List            =   "Form11.frx":6347D
            TabIndex        =   37
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form11.frx":634C4
            Left            =   3720
            List            =   "Form11.frx":634F2
            TabIndex        =   35
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox Text5 
            Height          =   405
            Left            =   8160
            TabIndex        =   34
            Top             =   480
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Hitung Nilai Siswa"
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
            TabIndex        =   28
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            Height          =   645
            Left            =   8160
            TabIndex        =   23
            Top             =   3000
            Width           =   2535
         End
         Begin VB.TextBox Text8 
            Height          =   405
            Left            =   8160
            TabIndex        =   22
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   8160
            TabIndex        =   21
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox Text6 
            Height          =   405
            Left            =   8160
            TabIndex        =   20
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            Height          =   405
            Left            =   3720
            TabIndex        =   13
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   3720
            TabIndex        =   12
            Top             =   480
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
            Left            =   1680
            TabIndex        =   11
            Top             =   5640
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
            Left            =   240
            TabIndex        =   10
            Top             =   4320
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
            Left            =   1680
            TabIndex        =   9
            Top             =   4800
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
            Left            =   1680
            TabIndex        =   8
            Top             =   3960
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
            Left            =   240
            TabIndex        =   7
            Top             =   5160
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
            Left            =   10200
            TabIndex        =   6
            Top             =   5520
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
            Left            =   8640
            TabIndex        =   5
            Top             =   5520
            Width           =   1335
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   360
            Top             =   2880
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
            Caption         =   "Kelas"
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
            Left            =   1800
            TabIndex        =   36
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label15 
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1800
            TabIndex        =   31
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai Keseluruhan siswa"
            BeginProperty Font 
               Name            =   "Rockwell"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   4320
            TabIndex        =   29
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Absensi"
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
            Left            =   6720
            TabIndex        =   27
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai UAS"
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
            Left            =   6840
            TabIndex        =   26
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai UTS"
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
            Left            =   6840
            TabIndex        =   25
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "No"
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
            Left            =   1800
            TabIndex        =   24
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pelajaran"
            BeginProperty Font 
               Name            =   "Rockwell Condensed"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   1800
            TabIndex        =   15
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Guru"
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
            Left            =   6840
            TabIndex        =   14
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   6735
         Left            =   -74400
         Picture         =   "Form11.frx":63583
         ScaleHeight     =   6675
         ScaleWidth      =   11955
         TabIndex        =   3
         Top             =   420
         Width           =   12015
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   330
            Left            =   240
            Top             =   5160
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            RecordSource    =   "guru"
            Caption         =   "Adodc4"
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "Form11.frx":84356
            Height          =   2415
            Left            =   600
            TabIndex        =   32
            Top             =   2520
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   4260
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "ID_GURU"
               Caption         =   "ID_GURU"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "NAMA_GURU"
               Caption         =   "NAMA_GURU"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "ALAMAT_GURU"
               Caption         =   "ALAMAT_GURU"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "TELEPONE"
               Caption         =   "TELEPONE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "JENIS_KELAMIN"
               Caption         =   "JENIS_KELAMIN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Guru"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   1215
            Left            =   720
            TabIndex        =   33
            Top             =   1200
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   6735
         Left            =   -74400
         Picture         =   "Form11.frx":8436B
         ScaleHeight     =   6675
         ScaleWidth      =   11955
         TabIndex        =   1
         Top             =   600
         Width           =   12015
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "Form11.frx":A513E
            Height          =   2415
            Left            =   960
            TabIndex        =   30
            Top             =   3600
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   4260
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "ID"
               Caption         =   "ID"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "NAMA_SISWA"
               Caption         =   "NAMA_SISWA"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "ALAMAT_SISWA"
               Caption         =   "ALAMAT_SISWA"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "TELEPONE"
               Caption         =   "TELEPONE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "JENIS_KELAMIN"
               Caption         =   "JENIS_KELAMIN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   1800
            Top             =   2400
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            Caption         =   "Adodc2"
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Siswa"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   1215
            Left            =   480
            TabIndex        =   2
            Top             =   1200
            Width           =   3855
         End
      End
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Guru"
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
      Left            =   7800
      TabIndex        =   17
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connOra As ADODB.Connection
Dim rsTab As ADODB.Recordset

Sub bersih()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End Sub
Private Sub brnBack_Click()
Form12.Show
Unload Me
End Sub

Private Sub btnCari_Click()
Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from nilai_siswa where no = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data telah ditemukan", vbInformation
        
         Text2.Text = Trim(rsTab!nama_siswa)
        Combo1.Text = Trim(rsTab!kelas)
        Combo2.Text = Trim(rsTab!nama_pelajaran)
        Text5.Text = Trim(rsTab!nama_guru)
        Text6.Text = Trim(rsTab!nilai_uts)
        Text7.Text = Trim(rsTab!nilai_uas)
        Text8.Text = Trim(rsTab!absensi)
        Text9.Text = Trim(rsTab!nilai_siswa)
    Else
        MsgBox "Data tidak ditemukan", vbInformation
        Call bersih
    End If
    rsTab.Close
    Set rsTab = Nothing
End Sub

Private Sub btnHapus_Click()
 Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from nilai_siswa where no = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
    rsTab.Delete
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan ID = " & Text1.Text & "Telah dihapus", vbInformation
End Sub

Private Sub btnTambah_Click()
Set rsTab = New ADODB.Recordset
    rsTab.Open "select * from nilai_siswa where no = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTab.RecordCount <> 0 Then
        MsgBox "Data dengan nomer ID tersebut sudah ada"
        rsTab.Close
        Set rsTab = Nothing
        Exit Sub
    Else
        Set rsTab = New ADODB.Recordset
        rsTab.Open "select * from nilai_siswa where no = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
        rsTab.AddNew
        rsTab!no = Trim(Text1.Text)
        rsTab!nama_siswa = Trim(Text2.Text)
        rsTab!kelas = Trim(Combo2.Text)
        rsTab!nama_pelajaran = Trim(Combo1.Text)
        rsTab!nama_guru = Trim(Text5.Text)
        rsTab!nilai_uts = Trim(Text6.Text)
        rsTab!nilai_uas = Trim(Text7.Text)
        rsTab!absensi = Trim(Text8.Text)
        rsTab!nilai_siswa = Trim(Text9.Text)
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
    rsTab.Open "select * from nilai_siswa where no = '" & Text1.Text & "'", connOra, adOpenKeyset, adLockOptimistic, adCmdText
     rsTab!no = Trim(Text1.Text)
        rsTab!nama_siswa = Trim(Text2.Text)
        rsTab!kelas = Trim(Combo2.Text)
        rsTab!nama_pelajaran = Trim(Combo1.Text)
        rsTab!nama_guru = Trim(Text5.Text)
        rsTab!nilai_uts = Trim(Text6.Text)
        rsTab!nilai_uas = Trim(Text7.Text)
        rsTab!absensi = Trim(Text8.Text)
        rsTab!nilai_siswa = Trim(Text9.Text)
    rsTab.Update
    connOra.Execute "Commit"
    rsTab.Close
    Set rsTab = Nothing
    MsgBox "Data dengan ID = " & Text1.Text & "Telah Berhasil di Update", vbInformation
    Call bersih
End Sub

Private Sub Command1_Click()
Form14.Show
Unload Me

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
Text9.Text = (Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)) / 3



End Sub



Private Sub DataGrid1_Click()

End Sub



Private Sub DataGrid3_Click()
Text5.Text = DataGrid3.Columns(1)
End Sub

Private Sub DataGrid4_Click()
Text2.Text = DataGrid4.Columns(1)
End Sub

Private Sub Form_Load()
Set connOra = New ADODB.Connection
connOra.Open "Provider=OraOLEDB.Oracle.1;Password=system;Persist Security Info=True;User ID=system"
connOra.CursorLocation = adUseClient
Call bersih
End Sub

