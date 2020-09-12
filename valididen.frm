VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      Picture         =   "valididen.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5535
      Left            =   5400
      OleObjectBlob   =   "valididen.frx":2D24
      TabIndex        =   0
      Top             =   2760
      Width           =   8535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tampilkan"
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
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "valididen.frx":4EC5
      Height          =   5775
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   10186
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
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
         Name            =   "Palatino Linotype"
         Size            =   9.75
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
      Height          =   615
      Left            =   480
      Top             =   2760
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      RecordSource    =   "denoriden"
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
      ItemData        =   "valididen.frx":4EDA
      Left            =   600
      List            =   "valididen.frx":4EE7
      TabIndex        =   1
      Text            =   "Pilih Populasi"
      Top             =   1560
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "valididen.frx":4F2D
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
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
      Left            =   600
      Top             =   5040
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Error :"
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
      Left            =   480
      TabIndex        =   3
      Top             =   8760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   10320
      Left            =   0
      Picture         =   "valididen.frx":4F42
      Top             =   -120
      Width           =   15000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim min As Double, xmax As Double, xstep As Single, numx As Integer, x As Double

If Combo1.Text = "Suspectible Population" Then
    Label3.Caption = "Graphic Identification of Suspectible Population"
    'grafik
    xmin = 0
    xmax = 84
    xstep = 1
    numx = (xmax - xmin) / xstep
    ReDim Values(1 To numx, 1)
    'memasukkan nilai data
    x = xmin
    Values(1, 0) = dataasli(0, 1)
    Values(1, 1) = dataasli(0, 1)
    For i = 2 To numx
        Values(i, 0) = dataasli(i - 1, 1)
        Values(i, 1) = Ddatain(i - 1, 1)
        x = x + xstep
    Next i
    Form1.MSChart1.RowCount = 2
    Form1.MSChart1.ColumnCount = numx
    Form1.MSChart1.ChartData = Values
    Form1.MSChart1.chartType = VtChChartType2dLine
    'keterangan
    With Form1.MSChart1.Legend
    .Location.Visible = True
    .Location.LocationType = VtChLocationTypeRight
    .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
    End With
    
    Form1.MSChart1.Plot.SeriesCollection(1).LegendText = "data asli"
    Form1.MSChart1.Plot.SeriesCollection(2).LegendText = "hasil identifikasi"
    Else
        If Combo1.Text = "Infected Population" Then
            Label3.Caption = "Graphic Identification of Infected Population"
            'grafik
            xmin = 0
            xmax = 84
            xstep = 1
            numx = (xmax - xmin) / xstep
            ReDim Values(1 To numx, 1)
            'memasukkan nilai data
            x = xmin
            Values(1, 0) = dataasli(0, 2)
            Values(1, 1) = dataasli(0, 2)
                For i = 2 To numx
                    Values(i, 0) = dataasli(i - 1, 2)
                    Values(i, 1) = Ddatain(i - 1, 2)
                    x = x + xstep
                Next i
                Form1.MSChart1.RowCount = 2
                Form1.MSChart1.ColumnCount = numx
                Form1.MSChart1.ChartData = Values
                Form1.MSChart1.chartType = VtChChartType2dLine
                'keterangan
                With Form1.MSChart1.Legend
                .Location.Visible = True
                .Location.LocationType = VtChLocationTypeRight
                .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
                End With
                Form1.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
                Form1.MSChart1.Plot.SeriesCollection(2).LegendText = "Hasil Identifikasi"
            Else
                If Combo1.Text = "Recovery Population" Then
                    Label3.Caption = "Graphic Identification of Recovery Population"
                    'grafik
                    xmin = 0
                    xmax = 84
                    xstep = 1
                    numx = (xmax - xmin) / xstep
                    ReDim Values(1 To numx, 1)
                    'memasukkan nilai data
                    x = xmin
                    Values(1, 0) = dataasli(0, 3)
                    Values(1, 1) = dataasli(0, 3)
                        For i = 2 To numx
                            Values(i, 0) = dataasli(i - 1, 3)
                            Values(i, 1) = Ddatain(i - 1, 3)
                            x = x + xstep
                        Next i
                        Form1.MSChart1.RowCount = 2
                        Form1.MSChart1.ColumnCount = numx
                        Form1.MSChart1.ChartData = Values
                        Form1.MSChart1.chartType = VtChChartType2dLine
                        'keterangan
                        With Form1.MSChart1.Legend
                        .Location.Visible = True
                        .Location.LocationType = VtChLocationTypeRight
                        .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
                        End With
    
                         Form1.MSChart1.Plot.SeriesCollection(1).LegendText = "data asli"
                        Form1.MSChart1.Plot.SeriesCollection(2).LegendText = "hasil identifikasi"
            
            End If
        End If
    End If
    
    
    
End Sub

Private Sub Picture1_Click()
hslidentifikasi.Show
End Sub
