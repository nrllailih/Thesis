VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
   LinkTopic       =   "Form3"
   ScaleHeight     =   9285
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   240
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5295
      Left            =   4920
      OleObjectBlob   =   "Form3.frx":2D24
      TabIndex        =   5
      Top             =   2760
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAMPILKAN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form3.frx":507A
      Height          =   5535
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9763
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":508F
      Height          =   2415
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4260
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
      Height          =   375
      Left            =   960
      Top             =   4440
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
      RecordSource    =   "denorup"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   960
      Top             =   3720
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "Form3.frx":50A4
      Left            =   480
      List            =   "Form3.frx":50B1
      TabIndex        =   0
      Text            =   "Pilih Populasi"
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000011&
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
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "error"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   1800
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hasil Peramalan"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   9285
      Left            =   0
      Picture         =   "Form3.frx":50F7
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim min As Double, xmax As Double, xstep As Single, numx As Integer, x As Double

If Combo1.Text = "Suspectible Population" Then
    Label2.Caption = "Graphic Identification of Suspectible Population"
    'grafik
    xmin = 0
    xmax = 17
    xstep = 1
    numx = (xmax - xmin) / xstep
    ReDim Values(1 To numx, 1)
    'memasukkan nilai data
    x = xmin
    'Values(1, 0) = dataasli(0, 1)
    'Values(1, 1) = dataasli(0, 1)
    For i = 2 To numx
        Values(i, 0) = dataasli(i + 63, 1)
        Values(i, 1) = Ddataup(i + 63, 1)
        x = x + xstep
    Next i
    Form3.MSChart1.RowCount = 2
    Form3.MSChart1.ColumnCount = numx
    Form3.MSChart1.ChartData = Values
    Form3.MSChart1.chartType = VtChChartType2dLine
    'keterangan
    With Form3.MSChart1.Legend
    .Location.Visible = True
    .Location.LocationType = VtChLocationTypeRight
    .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
    End With
    
    Form3.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
    Form3.MSChart1.Plot.SeriesCollection(2).LegendText = "Hasil Peramalan"
    Else
        If Combo1.Text = "Infected Population" Then
            Label2.Caption = "Graphic Identification of Infected Population"
            'grafik
            xmin = 0
            xmax = 17
            xstep = 1
            numx = (xmax - xmin) / xstep
            ReDim Values(1 To numx, 1)
            'memasukkan nilai data
            x = xmin
            'Values(1, 0) = dataasli(0, 2)
            'Values(1, 1) = dataasli(0, 2)
                For i = 2 To numx
                    Values(i, 0) = dataasli(i + 63, 2)
                    Values(i, 1) = Ddataup(i + 63, 2)
                    x = x + xstep
                Next i
                Form3.MSChart1.RowCount = 2
                Form3.MSChart1.ColumnCount = numx
                Form3.MSChart1.ChartData = Values
                Form3.MSChart1.chartType = VtChChartType2dLine
                'keterangan
                With Form3.MSChart1.Legend
                .Location.Visible = True
                .Location.LocationType = VtChLocationTypeRight
                .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
                End With
                Form3.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
                Form3.MSChart1.Plot.SeriesCollection(2).LegendText = "Hasil Peramalan"
            Else
                If Combo1.Text = "Recovery Population" Then
                    Label2.Caption = "Graphic Identification of Recovery Population"
                    'grafik
                    xmin = 0
                    xmax = 17
                    xstep = 1
                    numx = (xmax - xmin) / xstep
                    ReDim Values(1 To numx, 1)
                    'memasukkan nilai data
                    x = xmin
                    'Values(1, 0) = dataasli(0, 3)
                    'Values(1, 1) = dataasli(0, 3)
                        For i = 2 To numx
                            Values(i, 0) = dataasli(i + 63, 3)
                            Values(i, 1) = Ddataup(i + 63, 3)
                            x = x + xstep
                        Next i
                        Form3.MSChart1.RowCount = 2
                        Form3.MSChart1.ColumnCount = numx
                        Form3.MSChart1.ChartData = Values
                        Form3.MSChart1.chartType = VtChChartType2dLine
                        'keterangan
                        With Form3.MSChart1.Legend
                        .Location.Visible = True
                        .Location.LocationType = VtChLocationTypeRight
                        .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
                        End With
    
                         Form3.MSChart1.Plot.SeriesCollection(1).LegendText = "data asli"
                        Form3.MSChart1.Plot.SeriesCollection(2).LegendText = "hasil uji peramalan"
            
            End If
        End If
    End If
    
    
End Sub

Private Sub Picture1_Click()
peramalan.Show

End Sub
