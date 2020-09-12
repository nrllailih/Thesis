VERSION 5.00
Begin VB.Form menu 
   Caption         =   "MENU UTAMA"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   735
      Left            =   10800
      Picture         =   "menu.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   7200
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   5160
      Picture         =   "menu.frx":2D24
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   5280
      Picture         =   "menu.frx":82B6
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   5040
      Picture         =   "menu.frx":AFC8
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "PERAMALAN"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   5
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "IDENTIFIKASI"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "ESTIMASI PARAMETER"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   3120
      X2              =   8760
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   8250
      Left            =   0
      Picture         =   "menu.frx":10A6C
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Picture1_Click()
epba.Show
menu.Hide
End Sub
