VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form hslidentifikasi 
   Caption         =   "vZ3"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   17355
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   15360
      TabIndex        =   144
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   15360
      TabIndex        =   143
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   15360
      TabIndex        =   142
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PERAMALAN"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   141
      Top             =   9600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VALIDASI"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      Picture         =   "hslidentifikasi.frx":0000
      TabIndex        =   140
      Top             =   9000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "hslidentifikasi.frx":112F16
      Height          =   2175
      Left            =   8040
      TabIndex        =   122
      Top             =   6840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8160
      Top             =   7440
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\programku\program\database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\programku\program\database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "jstmse"
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
   Begin VB.OptionButton Option2 
      Caption         =   "RECOVERY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   4320
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SUSCEPTIBLE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "INFORMASI INPUT"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Label7"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Label6"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Label5"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Batas Error"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Maks Iterasi"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Learning Rate"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "hslidentifikasi.frx":112F2B
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   1920
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\programku\program\database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\programku\program\database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "data"
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
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   145
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label wbr31 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   139
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label wbr21 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   138
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label vbr31 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   137
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label vbr21 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   136
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label wbr01 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   134
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label wbr11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   133
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label vbr03 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   132
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label vbr02 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   131
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label vbr01 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   130
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label vbr33 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   129
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label vbr32 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   128
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label vbr23 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   127
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label vbr22 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   126
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label vbr13 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   125
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label vbr12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   124
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label vbr11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   123
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label w01ar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   121
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label w31ar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   120
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label w21ar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   119
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label w11ar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   118
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label v03ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   117
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v02ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   116
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v01ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   115
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v33ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   114
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v32ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   113
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v31ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   112
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v23ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   111
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v22ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   110
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v21ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   109
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v13ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   108
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v12ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   107
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v11ar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   106
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v03as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   105
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v02as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   104
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   103
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label v01as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "v01as"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   102
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v33as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   101
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v32as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   100
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v31as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   99
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v23as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   98
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v22as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   97
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v21as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   96
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v13as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   95
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v12as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   94
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v11as 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   93
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MEAN SQUARE ERROR"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7920
      TabIndex        =   92
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   44
      Left            =   12720
      TabIndex        =   91
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   43
      Left            =   12720
      TabIndex        =   90
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   42
      Left            =   12720
      TabIndex        =   89
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label wbi01 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   88
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label wbi21 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   87
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label wbi11 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   86
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   41
      Left            =   11520
      TabIndex        =   85
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   40
      Left            =   11520
      TabIndex        =   84
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   39
      Left            =   11520
      TabIndex        =   83
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label vbi02 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   82
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label vbi22 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   81
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label vbi12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   80
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label vbi01 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   79
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label vbi21 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   78
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label vbi11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   77
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   38
      Left            =   7800
      TabIndex        =   76
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   37
      Left            =   7800
      TabIndex        =   75
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   36
      Left            =   7800
      TabIndex        =   74
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   35
      Left            =   10200
      TabIndex        =   73
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   33
      Left            =   7800
      TabIndex        =   71
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label wbs01 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   70
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label wbs31 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   69
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label wbs21 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   68
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label wbs11 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   67
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   29
      Left            =   12720
      TabIndex        =   66
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label vbs03 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   65
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label vbs33 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   64
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label vbs23 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   63
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label vbs13 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   62
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label vbs02 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   61
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label vbs32 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   60
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label vbs22 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   59
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label vbs12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   58
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label vbs31 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   57
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label vbs21 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   56
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label vbs11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   55
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   28
      Left            =   7800
      TabIndex        =   54
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   27
      Left            =   7800
      TabIndex        =   53
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   26
      Left            =   7800
      TabIndex        =   52
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   25
      Left            =   7800
      TabIndex        =   51
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   24
      Left            =   11400
      TabIndex        =   50
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   23
      Left            =   10200
      TabIndex        =   49
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   22
      Left            =   9000
      TabIndex        =   48
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   21
      Left            =   7800
      TabIndex        =   47
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BOBOT BIAS OPTIMAL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7920
      TabIndex        =   46
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   7800
      X2              =   7800
      Y1              =   1080
      Y2              =   10320
   End
   Begin VB.Label w01ai 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   45
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label w21ai 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   44
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label w11ai 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   43
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   20
      Left            =   3840
      TabIndex        =   42
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   19
      Left            =   3840
      TabIndex        =   41
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   18
      Left            =   3840
      TabIndex        =   40
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label v02ai 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   39
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label v22ai 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   38
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label v12ai 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   37
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label v01ai 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   36
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label v21ai 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   35
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label v11ai 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   34
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   17
      Left            =   5040
      TabIndex        =   33
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   15
      Left            =   120
      TabIndex        =   31
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   14
      Left            =   5040
      TabIndex        =   30
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   13
      Left            =   5040
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   12
      Left            =   120
      TabIndex        =   28
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "BOBOT BIAS AWAL POPULASI INFECTED"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   7200
      Width           =   5295
   End
   Begin VB.Label w01as 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label w31as 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label w21as 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   24
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label w11as 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   23
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   11
      Left            =   5040
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   10
      Left            =   2520
      TabIndex        =   21
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   9
      Left            =   1320
      TabIndex        =   20
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   4
      Left            =   3720
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BOBOT BIAS AWAL POPULASI"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PENYEBARAN PENYAKIT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -240
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label vbs01 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   135
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Z1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   34
      Left            =   9000
      TabIndex        =   72
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   10320
      Left            =   0
      Picture         =   "hslidentifikasi.frx":112F40
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "hslidentifikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim indeks As Integer, sum1 As Double, mmrea As Double
'input data acces ke vb
Form1.Adodc2.Recordset.MoveFirst
For i = 0 To 83
    dataasli(i, 0) = Form1.Adodc2.Recordset.Fields("No").Value
    dataasli(i, 1) = Form1.Adodc2.Recordset.Fields("S").Value
    dataasli(i, 2) = Form1.Adodc2.Recordset.Fields("I").Value
    dataasli(i, 3) = Form1.Adodc2.Recordset.Fields("R").Value
    dataasli(i, 4) = Form1.Adodc2.Recordset.Fields("N").Value
    Form1.Adodc2.Recordset.MoveNext
Next i

'normalisasi data
'untuk M

For i = 1 To 3 '(S/I/R/N)
    maks(i) = dataasli(0, i)
    min(i) = dataasli(0, i)
Next i

For i = 1 To 3 '(S/I/R/N/M)
    For j = 0 To 83 '84 data
        If maks(i) < dataasli(j, i) Then
            maks(i) = dataasli(j, i)
        End If
        If min(i) > dataasli(j, i) Then
            min(i) = dataasli(j, i)
        End If
    Next j
    'Form1.List4.AddItem (min(i))
    'Form1.List5.AddItem (maks(i))
Next i


'Denormalisasi
For i = 0 To 83
    Ddatain(i, 0) = i + 1
Next i
    
For i = 1 To 3
    For j = 0 To 83 'jumlah data
        Ddatain(j, i) = Round(((outin(i - 1, j) + 1) * ((maks(i) - min(i)) / 2) + min(i)), 6)
        
       'Form1.List1.AddItem (outin(1, j))
       'Form1.List2.AddItem (outin(2, j))
       'Form1.List3.AddItem (outin(0, j))
    Next j
    
Next i
'validasi model (mmre)
indeks = 1
For j = 0 To 83
    sum1 = 0
        For k = 1 To 3 '3 populasi
        errorval(j, k) = mmre(dataasli(j, k), Ddatain(j, k))
        sum1 = sum1 + errorval(j, k)
        Next k
        errorval(j, k) = sum1 / 3
Next j
sum1 = 0
For j = 0 To 83
    sum1 = sum1 + errorval(j, 3)
Next j
mmrea = Round(sum1 / 84, 6)
Form1.Label2.Caption = mmrea

For i = 0 To 83
    Form1.Adodc1.Recordset.AddNew
    Form1.Adodc1.Recordset.Fields("Bulan ke") = Ddatain(i, 0)
    Form1.Adodc1.Recordset.Fields("S") = Ddatain(i, 1)
    Form1.Adodc1.Recordset.Fields("I") = Ddatain(i, 2)
    Form1.Adodc1.Recordset.Fields("R") = Ddatain(i, 3)
    Form1.Adodc1.Recordset.Update
    Form1.DataGrid1.Refresh
Next i


Form1.Show
hslidentifikasi.Hide

        
            

    
End Sub

Private Sub Command2_Click()
Dim pum1 As Double, pum2 As Double, pum3 As Double, pum4 As Double, indeks1 As Integer, sum1 As Double, ervup As Double
peramalan.Label2.Caption = learning
peramalan.Label3.Caption = maksiter
peramalan.Label4.Caption = makseror
'poladata peramalan
For i = 0 To 2
  For j = 0 To 57
    poladatar(i, j, 0) = j
    Next j
    If i = 0 Then
        For j = 0 To 57
            poladatar(i, j, 1) = normalisasi(j, 1)
            poladatar(i, j, 2) = normalisasi(j + 1, 1)
            poladatar(i, j, 3) = normalisasi(j + 2, 1)
            poladatar(i, j, 4) = normalisasi(j + 3, 1)
        Next j
    Else
    If i = 1 Then
        For j = 0 To 57
            poladatar(i, j, 1) = normalisasi(j, 3)
            poladatar(i, j, 2) = normalisasi(j + 1, 3)
            poladatar(i, j, 3) = normalisasi(j + 2, 3)
            poladatar(i, j, 4) = normalisasi(j + 3, 3)
        Next j
    Else
        For j = 0 To 57
            poladatar(i, j, 1) = normalisasi(j, 2)
            poladatar(i, j, 2) = normalisasi(j + 1, 2)
            poladatar(i, j, 3) = normalisasi(j + 2, 2)
        Next j
    End If
    End If
Next i
'bobot bias awal peramalan
For i = 0 To 2 Step 1
    If i = 0 Then
        'populasi S
        bbpr(i, 1) = bbin(i, 1)
        bbpr(i, 2) = bbin(i, 2)
        bbpr(i, 3) = bbin(i, 3)
        bbpr(i, 4) = bbin(i, 4)
        bbpr(i, 5) = bbin(i, 5)
        bbpr(i, 6) = bbin(i, 6)
        bbpr(i, 7) = bbin(i, 7)
        bbpr(i, 8) = bbin(i, 8)
        bbpr(i, 9) = bbin(i, 9)
        BPpr(i, 10) = bbin(i, 10)
        bbpr(i, 11) = bbin(i, 11)
        bbpr(i, 12) = bbin(i, 12)
        bbpr(i, 13) = bbin(i, 13)
        bbpr(i, 14) = bbin(i, 14)
        bbpr(i, 15) = bbin(i, 15)
        bbpr(i, 16) = bbin(i, 16)
    Else
    If i = 1 Then
       'populasi R
        bbpr(i, 1) = bbin(i, 1)
        bbpr(i, 2) = bbin(i, 2)
        bbpr(i, 3) = bbin(i, 3)
        bbpr(i, 4) = bbin(i, 4)
        bbpr(i, 5) = bbin(i, 5)
        bbpr(i, 6) = bbin(i, 6)
        bbpr(i, 7) = bbin(i, 7)
        bbpr(i, 8) = bbin(i, 8)
        bbpr(i, 9) = bbin(i, 9)
        BPpr(i, 10) = bbin(i, 10)
        bbpr(i, 11) = bbin(i, 11)
        bbpr(i, 12) = bbin(i, 12)
        bbpr(i, 13) = bbin(i, 13)
        bbpr(i, 14) = bbin(i, 14)
        bbpr(i, 15) = bbin(i, 15)
        bbpr(i, 16) = bbin(i, 16)
    Else
        'populasi I
        bbpr(i, 1) = bbin(i, 1)
        bbpr(i, 2) = bbin(i, 2)
        bbpr(i, 4) = bbin(i, 4)
        bbpr(i, 5) = bbin(i, 5)
        bbpr(i, 10) = bbin(i, 10)
        bbpr(i, 11) = bbin(i, 11)
        bbpr(i, 13) = bbin(i, 13)
        bbpr(i, 14) = bbin(i, 14)
        bbpr(i, 16) = bbin(i, 16)
    End If
    End If
Next i
poh = 0
Do
'proses feedforward dan backpropagation error, hitung mse
For i = 0 To 1
    For j = 0 To 57
    'feedforward
    pum1 = 0
    pum2 = 0
    pum3 = 0
    For k = 1 To 3 '3 inputan
        pum1 = pum1 + bbpr(i, k) * poladatar(i, j, k)
        pum2 = pum2 + bbpr(i, k + 3) * poladatar(i, j, k)
        pum3 = pum3 + bbpr(i, k + 6) * poladatar(i, j, k)
    Next k
    FFpr(i, 0) = bbpr(i, 10) + pum1
    FFpr(i, 1) = bbpr(i, 11) + pum2
    FFpr(i, 2) = bbpr(i, 12) + pum3
    FFpr(i, 3) = FsAktivasi(FFpr(i, 0))
    FFpr(i, 4) = FsAktivasi(FFpr(i, 1))
    FFpr(i, 5) = FsAktivasi(FFpr(i, 2))
    pum4 = 0
    For k = 1 To 3 'hidden layer ada 3
        pum4 = pum4 + bbpr(i, k + 12) * FFpr(i, k + 1)
    Next k
    FFpr(i, 6) = bbpr(i, 16) + pum4
    outinpr(i, j) = FsAktivasi(FFpr(i, 6))
    ermsep(i, j, 0) = j
    ermsep(i, j, 1) = MSE(poladatar(i, j, 4), outinpr(i, j))
    'Backpropagation Error
    BPpr(i, 0) = (poladatar(i, j, 4) - outinpr(i, j)) * FsAktivasi(FFpr(i, 6))
    BPpr(i, 1) = learning * BPpr(i, 0) * FFpr(i, 3)
    BPpr(i, 2) = learning * BPpr(i, 0) * FFpr(i, 4)
    BPpr(i, 3) = learning * BPpr(i, 0) * FFpr(i, 5)
    BPpr(i, 4) = learning * BPpr(i, 0)
    BPpr(i, 5) = BPpr(i, 0) * bbpr(i, 13)
    BPpr(i, 6) = BPpr(i, 0) * bbpr(i, 14)
    BPpr(i, 7) = BPpr(i, 0) * bbpr(i, 15)
    BPpr(i, 8) = BPpr(i, 5) * Fsaksaktivasi(FFpr(i, 0))
    BPpr(i, 9) = BPpr(i, 6) * Fsaksaktivasi(FFpr(i, 1))
    BPpr(i, 10) = BPpr(i, 7) * Fsaksaktivasi(FFpr(i, 2))
    BPpr(i, 11) = learning * BPpr(i, 8) * poladatar(i, j, 1)
    BPpr(i, 12) = learning * BPpr(i, 8) * poladatar(i, j, 2)
    BPpr(i, 13) = learning * BPpr(i, 8) * poladatar(i, j, 3)
    BPpr(i, 14) = learning * BPpr(i, 9) * poladatar(i, j, 1)
    BPpr(i, 15) = learning * BPpr(i, 9) * poladatar(i, j, 2)
    BPpr(i, 16) = learning * BPpr(i, 9) * poladatar(i, j, 3)
    BPpr(i, 17) = learning * BPpr(i, 10) * poladatar(i, j, 1)
    BPpr(i, 18) = learning * BPpr(i, 10) * poladatar(i, j, 2)
    BPpr(i, 19) = learning * BPpr(i, 10) * poladatar(i, j, 3)
    BPpr(i, 20) = learning * BPpr(i, 8)
    BPpr(i, 21) = learning * BPpr(i, 9)
    BPpr(i, 22) = learning * BPpr(i, 10)
    'update bobot dan bias
    For k = 1 To 12
    bbpr(i, k) = bbpr(i, k) + BPpr(i, k + 10)
    Next k
    For k = 1 To 4
    bbpr(i, k + 12) = bbpr(i, k + 12) + BPpr(i, k)
    Next k
    Next j
Next i
    
'feedforward, backpropagation error, update bobot dan bias untuk populasi I
For i = 2 To 2
    For j = 0 To 57
    'feedforward
    pum1 = 0
    pum2 = 0
    For k = 1 To 2 'inputan ada 2
        pum1 = pum1 + bbpr(i, k) * poladatar(i, j, k)
        pum2 = pum2 + bbpr(i, k + 3) * poladatar(i, j, k)
    Next k
    FFpr(i, 0) = bbpr(i, 10) + pum1
    FFpr(i, 1) = bbpr(i, 11) + pum2
    FFpr(i, 3) = FsAktivasi(FFpr(i, 0))
    FFpr(i, 4) = FsAktivasi(FFpr(i, 1))
    pum3 = 0
    For k = 1 To 2 'hidden layer ada 2
        pum3 = pum3 + bbpr(i, k + 12) * FFpr(i, k + 1)
    Next k
    FFpr(i, 6) = bbpr(i, 16) + pum3
    outinpr(i, j) = FsAktivasi(FFpr(i, 6))
    ermsep(i, j, 0) = j
    ermsep(i, j, 1) = MSE(poladatar(i, j, 3), outinpr(i, j))
    'backpro
    BPpr(i, 0) = (poladatar(i, j, 3) - outinpr(i, j)) * Fsaksaktivasi(FFpr(i, 6))
    BPpr(i, 2) = learning * BPpr(i, 0) * FFpr(i, 1)
    BPpr(i, 4) = learning * BPpr(i, 0)
    BPpr(i, 5) = BPpr(i, 0) * bbpr(i, 13)
    BPpr(i, 6) = BPpr(i, 0) * bbpr(i, 14)
    BPpr(i, 8) = BPpr(i, 5) * Fsaksaktivasi(FFpr(i, 0))
    BPpr(i, 9) = BPpr(i, 6) * Fsaksaktivasi(FFpr(i, 1))
    BPpr(i, 11) = learning * BPpr(i, 8) * poladatar(i, j, 1)
    BPpr(i, 12) = learning * BPpr(i, 8) * poladatar(i, j, 2)
    BPpr(i, 14) = learning * BPpr(i, 9) * poladatar(i, j, 1)
    BPpr(i, 15) = learning * BPpr(i, 9) * poladatar(i, j, 2)
    BPpr(i, 20) = learning * BPpr(i, 8)
    BPpr(i, 21) = learning * BPpr(i, 9)
    'upadate bobot dan bias
    For k = 1 To 2
    bbpr(i, k) = bbpr(i, k) + BPpr(i, k + 10)
    bbpr(i, k + 3) = bbpr(i, k + 3) + BPpr(i, k + 13)
    bbpr(i, k + 9) = bbpr(i, k + 9) + BPpr(i, k + 19)
    bbpr(i, k + 12) = bbpr(i, k + 12) + BPpr(i, k + 1)
    Next k
    bbpr(i, 16) = bbpr(i, 16) + BPpr(i, 4)
    Next j
Next i

For i = 0 To 2
    For j = 0 To 57
        For k = 1 To 16
            If bbpr(i, k) > 1 Or bbpr(i, k) < -1 Then
                bbpr(i, k) = FsAktivasi(bbpr(i, k))
            End If
        Next k
    Next j
Next i
'MSE akhir
For i = 0 To 2
    pum1 = 0
    For j = 0 To 57
        pum1 = pum1 + ermsep(i, j, 1)
    Next j
    ermsep(i, 0, 2) = pum1 / 57
Next i
pum1 = 0
For i = 0 To 2
    pum1 = pum1 + ermsep(i, 0, 2)
Next i

err = Round(pum1 / 3, 6)
poh = poh + 1
        
    
        
Loop Until poh = maksiter Or err <= makseror

peramalan.Label62.Caption = err

        


'input bobot bias identifikasi
peramalan.v11bs.Caption = Round(bbin(0, 1), 6)
peramalan.v12bs.Caption = Round(bbin(0, 4), 6)
peramalan.v13bs.Caption = Round(bbin(0, 7), 6)
peramalan.v21bs.Caption = Round(bbin(0, 2), 6)
peramalan.v22bs.Caption = Round(bbin(0, 5), 6)
peramalan.v23bs.Caption = Round(bbin(0, 8), 6)
peramalan.v31bs.Caption = Round(bbin(0, 3), 6)
peramalan.v32bs.Caption = Round(bbin(0, 6), 6)
peramalan.v33bs.Caption = Round(bbin(0, 9), 6)
peramalan.b11bs.Caption = Round(bbin(0, 10), 6)
peramalan.b12bs.Caption = Round(bbin(0, 11), 6)
peramalan.b13bs.Caption = Round(bbin(0, 12), 6)
peramalan.w11bs.Caption = Round(bbin(0, 13), 6)
peramalan.w21bs.Caption = Round(bbin(0, 14), 6)
peramalan.w31bs.Caption = Round(bbin(0, 15), 6)
peramalan.b01bs.Caption = Round(bbin(0, 16), 6)
peramalan.v11br.Caption = Round(bbin(1, 1), 6)
peramalan.v12br.Caption = Round(bbin(1, 4), 6)
peramalan.v13br.Caption = Round(bbin(1, 7), 6)
peramalan.v21br.Caption = Round(bbin(1, 2), 6)
peramalan.v22br.Caption = Round(bbin(1, 5), 6)
peramalan.v23br.Caption = Round(bbin(1, 8), 6)
peramalan.v31br.Caption = Round(bbin(1, 3), 6)
peramalan.v32br.Caption = Round(bbin(1, 6), 6)
peramalan.v33br.Caption = Round(bbin(1, 9), 6)
peramalan.b11br.Caption = Round(bbin(1, 10), 6)
peramalan.b12br.Caption = Round(bbin(1, 11), 6)
peramalan.b13br.Caption = Round(bbin(1, 12), 6)
peramalan.w11br.Caption = Round(bbin(1, 13), 6)
peramalan.w21br.Caption = Round(bbin(1, 14), 6)
peramalan.w31br.Caption = Round(bbin(1, 15), 6)
peramalan.b01br.Caption = Round(bbin(1, 16), 6)
peramalan.v11bi.Caption = Round(bbin(2, 1), 6)
peramalan.v12bi.Caption = Round(bbin(2, 4), 6)
peramalan.v21bi.Caption = Round(bbin(2, 2), 6)
peramalan.v22bi.Caption = Round(bbin(2, 5), 6)
peramalan.b11bi.Caption = Round(bbin(2, 10), 6)
peramalan.b12bi.Caption = Round(bbin(2, 11), 6)
peramalan.w11bi.Caption = Round(bbin(2, 13), 6)
peramalan.w21bi.Caption = Round(bbin(2, 14), 6)
peramalan.b01bi.Caption = Round(bbin(2, 16), 6)

'input bobot dan bias optimal peramalan
peramalan.v11os.Caption = Round(bbpr(0, 1), 6)
peramalan.v12os.Caption = Round(bbpr(0, 4), 6)
peramalan.v13os.Caption = Round(bbpr(0, 7), 6)
peramalan.v21os.Caption = Round(bbpr(0, 2), 6)
peramalan.v220s.Caption = Round(bbpr(0, 5), 6)
peramalan.v23os.Caption = Round(bbpr(0, 8), 6)
peramalan.v31os.Caption = Round(bbpr(0, 3), 6)
peramalan.v32os.Caption = Round(bbpr(0, 6), 6)
peramalan.v33os.Caption = Round(bbpr(0, 9), 6)
peramalan.b11os.Caption = Round(bbpr(0, 10), 6)
peramalan.b12os.Caption = Round(bbpr(0, 11), 6)
peramalan.b13os.Caption = Round(bbpr(0, 12), 6)
peramalan.w11os.Caption = Round(bbpr(0, 13), 6)
peramalan.w21os.Caption = Round(bbpr(0, 14), 6)
peramalan.w31os.Caption = Round(bbpr(0, 15), 6)
peramalan.b01os.Caption = Round(bbpr(0, 16), 6)
peramalan.v11or.Caption = Round(bbpr(1, 1), 6)
peramalan.v12or.Caption = Round(bbpr(1, 4), 6)
peramalan.v13or.Caption = Round(bbpr(1, 7), 6)
peramalan.v21or.Caption = Round(bbpr(1, 2), 6)
peramalan.v22or.Caption = Round(bbpr(1, 5), 6)
peramalan.v23or.Caption = Round(bbpr(1, 8), 6)
peramalan.v31or.Caption = Round(bbpr(1, 3), 6)
peramalan.v32or.Caption = Round(bbpr(1, 6), 6)
peramalan.v33or.Caption = Round(bbpr(1, 9), 6)
peramalan.b12or.Caption = Round(bbpr(1, 11), 6)
peramalan.b11or.Caption = Round(bbpr(1, 10), 6)
peramalan.b13or.Caption = Round(bbpr(1, 12), 6)
peramalan.w11or.Caption = Round(bbpr(1, 13), 6)
peramalan.w21or.Caption = Round(bbpr(1, 14), 6)
peramalan.w31or.Caption = Round(bbpr(1, 15), 6)
peramalan.b01or.Caption = Round(bbpr(1, 16), 6)
peramalan.v11oi.Caption = Round(bbpr(2, 1), 6)
peramalan.v12oi.Caption = Round(bbpr(2, 4), 6)
peramalan.v21oi.Caption = Round(bbpr(2, 2), 6)
peramalan.v22oi.Caption = Round(bbpr(2, 5), 6)
peramalan.b11oi.Caption = Round(bbpr(2, 10), 6)
peramalan.b12oi.Caption = Round(bbpr(2, 11), 6)
peramalan.w11oi.Caption = Round(bbpr(2, 13), 6)
peramalan.w21oi.Caption = Round(bbpr(2, 14), 6)
peramalan.b01oi.Caption = Round(bbpr(2, 16), 6)



peramalan.Option1.Value = True
peramalan.v11bi.Visible = False
peramalan.v12bi.Visible = False
peramalan.v21bi.Visible = False
peramalan.v22bi.Visible = False
peramalan.b12bi.Visible = False
peramalan.b11bi.Visible = False
peramalan.Label34.Visible = False
peramalan.Label35.Visible = False
peramalan.Label36.Visible = False
peramalan.Label27.Visible = False
peramalan.Label30.Visible = False
peramalan.Label33.Visible = False
peramalan.Label40.Visible = False
peramalan.w11bi.Visible = False
peramalan.w21bi.Visible = False
peramalan.b01bi.Visible = False
peramalan.v11br.Visible = False
peramalan.v12br.Visible = False
peramalan.v13br.Visible = False
peramalan.v21br.Visible = False
peramalan.v22br.Visible = False
peramalan.v23br.Visible = False
peramalan.v33br.Visible = False
peramalan.b11br.Visible = False
peramalan.b12br.Visible = False
peramalan.b13br.Visible = False
peramalan.w11br.Visible = False
peramalan.w21br.Visible = False
peramalan.w31br.Visible = False
peramalan.b01br.Visible = False
peramalan.v11oi.Visible = False
peramalan.v12oi.Visible = False
peramalan.v21oi.Visible = False
peramalan.v22oi.Visible = False
peramalan.b11oi.Visible = False
peramalan.b12oi.Visible = False
peramalan.w11oi.Visible = False
peramalan.w21oi.Visible = False
peramalan.b01oi.Visible = False
peramalan.Label70.Visible = False
peramalan.Label73.Visible = False
peramalan.Label77.Visible = False
peramalan.Label78.Visible = False
peramalan.Label79.Visible = False
peramalan.Label83.Visible = False
peramalan.b13oi.Visible = False
peramalan.v11or.Visible = False
peramalan.v12or.Visible = False
peramalan.v13or.Visible = False
peramalan.v21or.Visible = False
peramalan.v22or.Visible = False
peramalan.v23or.Visible = False
peramalan.v31or.Visible = False
peramalan.v32or.Visible = False
peramalan.v33or.Visible = False
peramalan.b11or.Visible = False
peramalan.b12or.Visible = False
peramalan.b13or.Visible = False
peramalan.w11or.Visible = False
peramalan.w21or.Visible = False
peramalan.w31or.Visible = False
peramalan.b01or.Visible = False


'UJI PERAMALAN
'pola data peramalan

For i = 0 To 2
  For j = 60 To 81
    poladataru(i, j, 0) = j
    Next j
    If i = 0 Then
        For j = 60 To 81
            poladataru(i, j, 1) = normalisasi(j, 1)
            poladataru(i, j, 2) = normalisasi(j + 1, 1)
            poladataru(i, j, 3) = normalisasi(j + 2, 1)
            'poladataru(i, j, 4) = normalisasi(j + 3, 1)
        Next j
    Else
    If i = 1 Then
        For j = 60 To 81
            poladataru(i, j, 1) = normalisasi(j, 3)
            poladataru(i, j, 2) = normalisasi(j + 1, 3)
            poladataru(i, j, 3) = normalisasi(j + 2, 3)
            'poladataru(i, j, 4) = normalisasi(j + 3, 3)
        Next j
    Else
        For j = 60 To 81
            poladataru(i, j, 1) = normalisasi(j, 2)
            poladataru(i, j, 2) = normalisasi(j + 1, 2)
            poladataru(i, j, 3) = normalisasi(j + 2, 2)
        Next j
    End If
    End If
Next i

'proses feedforwar untuk S dan R
For i = 0 To 1
    For j = 60 To 81
    'feedforward
    pum1 = 0
    pum2 = 0
    pum3 = 0
    For k = 1 To 3 '3 inputan
        pum1 = pum1 + bbpr(i, k) * poladataru(i, j, k)
        pum2 = pum2 + bbpr(i, k + 3) * poladataru(i, j, k)
        pum3 = pum3 + bbpr(i, k + 6) * poladataru(i, j, k)
    Next k
    FFpr(i, 0) = bbpr(i, 10) + pum1
    FFpr(i, 1) = bbpr(i, 11) + pum2
    FFpr(i, 2) = bbpr(i, 12) + pum3
    FFpr(i, 3) = FsAktivasi(FFpr(i, 0))
    FFpr(i, 4) = FsAktivasi(FFpr(i, 1))
    FFpr(i, 5) = FsAktivasi(FFpr(i, 2))
    pum4 = 0
    For k = 1 To 3 'hidden layer ada 3
        pum4 = pum4 + bbpr(i, k + 12) * FFpr(i, k + 1)
    Next k
    FFpr(i, 6) = bbpr(i, 16) + pum4
    outinpr(i, j) = FsAktivasi(FFpr(i, 6))
    'ermsep(i, j, 0) = j
    'ermsep(i, j, 1) = MSE(poladataru(i, j, 4), outinpr(i, j))
    Next j
Next i
'proses feedforward untuk I
For i = 2 To 2
    For j = 60 To 81
    'feedforward
    pum1 = 0
    pum2 = 0
    For k = 1 To 2 'inputan ada 2
        pum1 = pum1 + bbpr(i, k) * poladataru(i, j, k)
        pum2 = pum2 + bbpr(i, k + 3) * poladataru(i, j, k)
    Next k
    FFpr(i, 0) = bbpr(i, 10) + pum1
    FFpr(i, 1) = bbpr(i, 11) + pum2
    FFpr(i, 3) = FsAktivasi(FFpr(i, 0))
    FFpr(i, 4) = FsAktivasi(FFpr(i, 1))
    pum3 = 0
    For k = 1 To 2 'hidden layer ada 2
        pum3 = pum3 + bbpr(i, k + 12) * FFpr(i, k + 1)
    Next k
    FFpr(i, 6) = bbpr(i, 16) + pum3
    outinpr(i, j) = FsAktivasi(FFpr(i, 6))
    ermsep(i, j, 0) = j
    ermsep(i, j, 1) = MSE(poladataru(i, j, 3), outinpr(i, j))
    'peramalan.List1.AddItem (ermsep(i, j, 4))
    Next j
Next i
'MSE akhir UP
'For i = 0 To 2
 '   pum1 = 0
  '  For j = 58 To 68
   '     pum1 = pum1 + ermsep(i, j, 1)
   ' Next j
   ' ermsep(i, 0, 2) = pum1 / 11
    'peramalan.List1.AddItem (ermsep(i, 0, 2))
'Next i
'pum1 = 0
'For i = 0 To 2
 '   pum1 = pum1 + ermsep(i, 0, 2)
    'peramalan.List1.AddItem (pum1)
'Next i

'errup = Round(pum1 / 3, 6)
'peramalan.Label6.Caption = errup

'denormalisai uji peramalan
For i = 60 To 81
    Ddataup(i, 0) = i + 1
Next i
For i = 1 To 3
    For j = 60 To 81
        Ddataup(j, i) = Round(((outinpr(i - 1, j) + 1) * (maks(i) - min(i)) / 2) + min(i), 6)
    Next j
Next i
'validasi model
indeks1 = 1
For j = 60 To 81
    sum1 = 0
    For k = 0 To 2
        errvup(0, j, k) = mmre(dataasli(j + 2, k + 1), Ddataup(j, k + 1))
        sum1 = sum1 + errvup(0, j, k)
    Next k
    errvup(0, j, 3) = sum1 / 3
Next j
sum1 = 0
For j = 60 To 81
    sum1 = sum1 + errvup(0, j, 3)
Next j
ervup = Round(sum1 / 66, 6)

For i = 60 To 81
    Form3.Adodc2.Recordset.AddNew
    Form3.Adodc2.Recordset.Fields("Bulan ke") = Ddataup(i, 0)
    Form3.Adodc2.Recordset.Fields("S") = Ddataup(i, 1)
    Form3.Adodc2.Recordset.Fields("I") = Ddataup(i, 2)
    Form3.Adodc2.Recordset.Fields("R") = Ddataup(i, 3)
    Form3.Adodc2.Recordset.Update
    Form3.DataGrid2.Refresh
Next i
Form3.Label4.Caption = ervup
    
    


peramalan.Show
hslidentifikasi.Hide
inidentify.Hide
End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
hslidentifikasi.v11as.Visible = True
hslidentifikasi.v12as.Visible = True
hslidentifikasi.v13as.Visible = True
hslidentifikasi.v21as.Visible = True
hslidentifikasi.v22as.Visible = True
hslidentifikasi.v23as.Visible = True
hslidentifikasi.v31as.Visible = True
hslidentifikasi.v32as.Visible = True
hslidentifikasi.v33as.Visible = True
hslidentifikasi.v01as.Visible = True
hslidentifikasi.v02as.Visible = True
hslidentifikasi.v03as.Visible = True
hslidentifikasi.w11as.Visible = True
hslidentifikasi.w21as.Visible = True
hslidentifikasi.w31as.Visible = True
hslidentifikasi.w01as.Visible = True

hslidentifikasi.vbs11.Visible = True
hslidentifikasi.vbs12.Visible = True
hslidentifikasi.vbs13.Visible = True
hslidentifikasi.vbs21.Visible = True
hslidentifikasi.vbs22.Visible = True
hslidentifikasi.vbs23.Visible = True
hslidentifikasi.vbs31.Visible = True
hslidentifikasi.vbs32.Visible = True
hslidentifikasi.vbs33.Visible = True
hslidentifikasi.vbs01.Visible = True
hslidentifikasi.vbs02.Visible = True
hslidentifikasi.vbs03.Visible = True
hslidentifikasi.wbs11.Visible = True
hslidentifikasi.wbs21.Visible = True
hslidentifikasi.wbs31.Visible = True
hslidentifikasi.wbs01.Visible = True

hslidentifikasi.v11ar.Visible = False
hslidentifikasi.v12ar.Visible = False
hslidentifikasi.v13ar.Visible = False
hslidentifikasi.v21ar.Visible = False
hslidentifikasi.v22ar.Visible = False
hslidentifikasi.v23ar.Visible = False
hslidentifikasi.v31ar.Visible = False
hslidentifikasi.v32ar.Visible = False
hslidentifikasi.v33ar.Visible = False
hslidentifikasi.v01ar.Visible = False
hslidentifikasi.v02ar.Visible = False
hslidentifikasi.v03ar.Visible = False
hslidentifikasi.w11ar.Visible = False
hslidentifikasi.w21ar.Visible = False
hslidentifikasi.w31ar.Visible = False
hslidentifikasi.w01ar.Visible = False
hslidentifikasi.vbr11.Visible = False
hslidentifikasi.vbr12.Visible = False
hslidentifikasi.vbr13.Visible = False
hslidentifikasi.vbr21.Visible = False
hslidentifikasi.vbr22.Visible = False
hslidentifikasi.vbr23.Visible = False
hslidentifikasi.vbr31.Visible = False
hslidentifikasi.vbr32.Visible = False
hslidentifikasi.vbr33.Visible = False
hslidentifikasi.vbr01.Visible = False
hslidentifikasi.vbr02.Visible = False
hslidentifikasi.vbr03.Visible = False
hslidentifikasi.wbr11.Visible = False
hslidentifikasi.wbr21.Visible = False
hslidentifikasi.wbr31.Visible = False
hslidentifikasi.wbr01.Visible = False
End Sub

Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
hslidentifikasi.v11as.Visible = False
hslidentifikasi.v12as.Visible = False
hslidentifikasi.v13as.Visible = False
hslidentifikasi.v21as.Visible = False
hslidentifikasi.v22as.Visible = False
hslidentifikasi.v23as.Visible = False
hslidentifikasi.v31as.Visible = False
hslidentifikasi.v32as.Visible = False
hslidentifikasi.v33as.Visible = False
hslidentifikasi.v01as.Visible = False
hslidentifikasi.v02as.Visible = False
hslidentifikasi.v03as.Visible = False
hslidentifikasi.w11as.Visible = False
hslidentifikasi.w21as.Visible = False
hslidentifikasi.w31as.Visible = False
hslidentifikasi.w01as.Visible = False
hslidentifikasi.vbs11.Visible = False
hslidentifikasi.vbs12.Visible = False
hslidentifikasi.vbs13.Visible = False
hslidentifikasi.vbs21.Visible = False
hslidentifikasi.vbs22.Visible = False
hslidentifikasi.vbs23.Visible = False
hslidentifikasi.vbs31.Visible = False
hslidentifikasi.vbs32.Visible = False
hslidentifikasi.vbs33.Visible = False
hslidentifikasi.vbs01.Visible = False
hslidentifikasi.vbs02.Visible = False
hslidentifikasi.vbs03.Visible = False
hslidentifikasi.wbs11.Visible = False
hslidentifikasi.wbs21.Visible = False
hslidentifikasi.wbs31.Visible = False
hslidentifikasi.wbs01.Visible = False



hslidentifikasi.v11ar.Visible = True
hslidentifikasi.v12ar.Visible = True
hslidentifikasi.v13ar.Visible = True
hslidentifikasi.v21ar.Visible = True
hslidentifikasi.v22ar.Visible = True
hslidentifikasi.v23ar.Visible = True
hslidentifikasi.v31ar.Visible = True
hslidentifikasi.v32ar.Visible = True
hslidentifikasi.v33ar.Visible = True
hslidentifikasi.v01ar.Visible = True
hslidentifikasi.v02ar.Visible = True
hslidentifikasi.v03ar.Visible = True
hslidentifikasi.w11ar.Visible = True
hslidentifikasi.w21ar.Visible = True
hslidentifikasi.w31ar.Visible = True
hslidentifikasi.w01ar.Visible = True

hslidentifikasi.vbr11.Visible = True
hslidentifikasi.vbr12.Visible = True
hslidentifikasi.vbr13.Visible = True
hslidentifikasi.vbr21.Visible = True
hslidentifikasi.vbr22.Visible = True
hslidentifikasi.vbr23.Visible = True
hslidentifikasi.vbr31.Visible = True
hslidentifikasi.vbr32.Visible = True
hslidentifikasi.vbr33.Visible = True
hslidentifikasi.vbr01.Visible = True
hslidentifikasi.vbr02.Visible = True
hslidentifikasi.vbr03.Visible = True
hslidentifikasi.wbr11.Visible = True
hslidentifikasi.wbr21.Visible = True
hslidentifikasi.wbr31.Visible = True
hslidentifikasi.wbr01.Visible = True


End Sub
