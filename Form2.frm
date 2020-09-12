VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   360
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAMPILKAN"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
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
      ItemData        =   "Form2.frx":2D24
      Left            =   360
      List            =   "Form2.frx":2D31
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1680
      Width           =   2775
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4215
      Left            =   4560
      OleObjectBlob   =   "Form2.frx":2D77
      TabIndex        =   2
      Top             =   2880
      Width           =   6975
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2.frx":5213
      Height          =   5055
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8916
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":5228
      Height          =   1935
      Left            =   720
      TabIndex        =   0
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3413
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
      Height          =   330
      Left            =   720
      Top             =   3360
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
      RecordSource    =   "normalisasi"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   720
      Top             =   3840
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
      RecordSource    =   "snum"
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   8250
      Left            =   0
      Picture         =   "Form2.frx":523D
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim min As Double, xmax As Double, xstep As Single, numx As Integer, x As Double

If Combo1.Text = "Suspectible Population" Then
    Label1.Caption = "Graphic Identification of Suspectible Population"
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
        Values(i, 1) = solnum(i - 1, 1)
        x = x + xstep
    Next i
    Form2.MSChart1.RowCount = 2
    Form2.MSChart1.ColumnCount = numx
    Form2.MSChart1.ChartData = Values
    Form2.MSChart1.chartType = VtChChartType2dLine
    'keterangan
    With Form2.MSChart1.Legend
    .Location.Visible = True
    .Location.LocationType = VtChLocationTypeRight
    .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
    End With
    
    Form2.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
    Form2.MSChart1.Plot.SeriesCollection(2).LegendText = "hasil simulasi parameter"
    Else
        If Combo1.Text = "Infected Population" Then
            Label1.Caption = "Graphic Identification of Infected Population"
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
                    Values(i, 1) = solnum(i - 1, 2)
                    x = x + xstep
                Next i
                Form2.MSChart1.RowCount = 2
                Form2.MSChart1.ColumnCount = numx
                Form2.MSChart1.ChartData = Values
                Form2.MSChart1.chartType = VtChChartType2dLine
                'keterangan
                With Form2.MSChart1.Legend
                .Location.Visible = True
                .Location.LocationType = VtChLocationTypeRight
                .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
                End With
                Form2.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
                Form2.MSChart1.Plot.SeriesCollection(2).LegendText = "Hasil Simulasi Parameter"
            Else
                If Combo1.Text = "Recovery Population" Then
                    Label1.Caption = "Graphic Identification of Recovery Population"
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
                            Values(i, 1) = solnum(i - 1, 3)
                            x = x + xstep
                        Next i
                        Form2.MSChart1.RowCount = 2
                        Form2.MSChart1.ColumnCount = numx
                        Form2.MSChart1.ChartData = Values
                        Form2.MSChart1.chartType = VtChChartType2dLine
                        'keterangan
                        With Form2.MSChart1.Legend
                        .Location.Visible = True
                        .Location.LocationType = VtChLocationTypeRight
                        .TextLayout.HorzAlignment = VtHorizontalAlignmentRight
                        End With
    
                         Form2.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
                        Form2.MSChart1.Plot.SeriesCollection(2).LegendText = "Hasil Simulasi Parameter"
            
            End If
        End If
    End If
    
    
    
End Sub

Private Sub Picture1_Click()
hslep.Show
Form2.Hide
End Sub
