VERSION 5.00
Begin VB.Form peramalan 
   Caption         =   "PERAMALAN"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form2"
   Picture         =   "peramalan.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Hasil Peramalan"
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
      Left            =   7560
      TabIndex        =   138
      Top             =   8280
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      Picture         =   "peramalan.frx":48516
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   137
      Top             =   0
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Informasi Input"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5280
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PILIH POPULASI"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   139
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label b01or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   135
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label w31or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   134
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label w21or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   133
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label w11or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   132
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label b13or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   131
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label b12or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   130
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label b11or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   129
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label v33or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   128
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v32or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   127
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v31or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   126
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v23or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   125
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v22or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   124
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v21or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   123
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v13or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   122
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v12or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   121
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v11or 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   120
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label b01br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   119
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label w31br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   118
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label w21br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   117
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label w11br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   116
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label b13br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   115
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label b12br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   114
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label b11br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   113
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v33br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   112
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v32br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   111
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v31br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   110
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label v23br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   109
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v22br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   108
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v21br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   107
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v13br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   106
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label v12br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   105
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label83 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   11760
      TabIndex        =   104
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label b01oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   103
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label w21oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   102
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label w11oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   101
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   9240
      TabIndex        =   100
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label78 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   8040
      TabIndex        =   99
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   6840
      TabIndex        =   98
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label b13oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   9240
      TabIndex        =   97
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label b12oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   96
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label b11oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   95
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   9240
      TabIndex        =   94
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v22oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   93
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v21oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   92
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      Left            =   9240
      TabIndex        =   91
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v12oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   90
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v11oi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   89
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label67 
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
      Left            =   10560
      TabIndex        =   88
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X3/b"
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
      Left            =   10560
      TabIndex        =   87
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label65 
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
      Left            =   9240
      TabIndex        =   86
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label64 
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
      Left            =   5640
      TabIndex        =   85
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X3/b"
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
      Left            =   5640
      TabIndex        =   84
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   83
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label b01os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   82
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label w31os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   81
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label w21os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   80
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label w11os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   11760
      TabIndex        =   79
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label b13os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   78
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label b12os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   77
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label b11os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   76
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label v33os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   75
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v32os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   74
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v31os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   73
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label v23os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   72
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v220s 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   71
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v21os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   70
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v13os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   9240
      TabIndex        =   69
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v12os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   8040
      TabIndex        =   68
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v11os 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Left            =   6840
      TabIndex        =   67
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label45 
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
      Left            =   120
      TabIndex        =   66
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X3/b"
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
      Left            =   120
      TabIndex        =   65
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label43 
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
      Left            =   3720
      TabIndex        =   64
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "X3/b"
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
      Left            =   120
      TabIndex        =   63
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label v11br 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label6"
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
      TabIndex        =   62
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   61
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label b01bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   60
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label w21bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   59
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label w11bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   58
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   57
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   56
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   55
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   54
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label b12bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   53
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label b11bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   52
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   51
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v22bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   50
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v21bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   49
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
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
      TabIndex        =   48
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label v12bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   47
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label v11bi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
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
      TabIndex        =   46
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label b01bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   45
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label w31bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   44
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label w21bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   43
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label21 
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
      Left            =   120
      TabIndex        =   42
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label w11bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   41
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label b12bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   40
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label b11bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   39
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label v33bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   38
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v32bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v31bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label v23bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   35
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v22bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   34
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v21bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   33
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label v13bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   32
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label v12bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   31
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label v11bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "MSE "
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
      Left            =   10560
      TabIndex        =   29
      Top             =   7200
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
      Index           =   21
      Left            =   10560
      TabIndex        =   28
      Top             =   5760
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
      Index           =   20
      Left            =   10560
      TabIndex        =   27
      Top             =   5280
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
      Left            =   5640
      TabIndex        =   26
      Top             =   6240
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
      Left            =   5640
      TabIndex        =   25
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
      Index           =   13
      Left            =   8040
      TabIndex        =   24
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
      Index           =   12
      Left            =   6840
      TabIndex        =   23
      Top             =   5280
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
      Index           =   11
      Left            =   5640
      TabIndex        =   22
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BOBOT BIAS PTIMAL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5760
      TabIndex        =   21
      Top             =   4680
      Width           =   4335
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
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   7440
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
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   6960
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
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   5280
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
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   4800
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
      TabIndex        =   16
      Top             =   4320
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
      Index           =   0
      Left            =   1320
      TabIndex        =   15
      Top             =   4320
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
      Index           =   17
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Recovery"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Infected"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Susceptible"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BOBOT BIAS AWAL PERAMALAN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label b13bs 
      Alignment       =   2  'Center
      Caption         =   "Label6"
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
      TabIndex        =   136
      Top             =   6240
      Width           =   1215
   End
End
Attribute VB_Name = "peramalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command2_Click()
Form3.Show

End Sub

Private Sub DataGrid2_Click()

End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Option3.Value = False

peramalan.v11bs.Visible = True
peramalan.v12bs.Visible = True
peramalan.v13bs.Visible = True
peramalan.v21bs.Visible = True
peramalan.v22bs.Visible = True
peramalan.v23bs.Visible = True
peramalan.v31bs.Visible = True
peramalan.v32bs.Visible = True
peramalan.v33bs.Visible = True
peramalan.b11bs.Visible = True
peramalan.b12bs.Visible = True
peramalan.b13bs.Visible = True
peramalan.w11bs.Visible = True
peramalan.w21bs.Visible = True
peramalan.w31bs.Visible = True
peramalan.b01bs.Visible = True
peramalan.v11os.Visible = True
peramalan.v12os.Visible = True
peramalan.v13os.Visible = True
peramalan.v21os.Visible = True
peramalan.v220s.Visible = True
peramalan.v23os.Visible = True
peramalan.v31os.Visible = True
peramalan.v32os.Visible = True
peramalan.v33os.Visible = True
peramalan.b11os.Visible = True
peramalan.b12os.Visible = True
peramalan.b13os.Visible = True
peramalan.w11os.Visible = True
peramalan.w21os.Visible = True
peramalan.w31os.Visible = True
peramalan.b01os.Visible = True
peramalan.v11br.Visible = False
peramalan.v12br.Visible = False
peramalan.v13br.Visible = False
peramalan.v21br.Visible = False
peramalan.v22br.Visible = False
peramalan.v23br.Visible = False
peramalan.v31br.Visible = False
peramalan.v32br.Visible = False
peramalan.v33br.Visible = False
peramalan.b11br.Visible = False
peramalan.b12br.Visible = False
peramalan.b13br.Visible = False
peramalan.w11br.Visible = False
peramalan.w21br.Visible = False
peramalan.w31br.Visible = False
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
peramalan.b01br.Visible = False
peramalan.v11bi.Visible = False
peramalan.v12bi.Visible = False
peramalan.v21bi.Visible = False
peramalan.v22bi.Visible = False
peramalan.b11bi.Visible = False
peramalan.b12bi.Visible = False
peramalan.w11bi.Visible = False
peramalan.w21bi.Visible = False
peramalan.b01bi.Visible = False
peramalan.v11oi.Visible = False
peramalan.v12oi.Visible = False
peramalan.v21oi.Visible = False
peramalan.v22oi.Visible = False
peramalan.b11oi.Visible = False
peramalan.b12oi.Visible = False
peramalan.w11oi.Visible = False
peramalan.w21oi.Visible = False
peramalan.b01oi.Visible = False




End Sub

Private Sub Option2_Click()
Option2.Value = True
Option1.Value = False
Option3.Value = False

peramalan.v11bs.Visible = False
peramalan.v12bs.Visible = False
peramalan.v13bs.Visible = False
peramalan.v21bs.Visible = False
peramalan.v22bs.Visible = False
peramalan.v23bs.Visible = False
peramalan.v31bs.Visible = False
peramalan.v32bs.Visible = False
peramalan.v33bs.Visible = False
peramalan.b11bs.Visible = False
peramalan.b12bs.Visible = False
peramalan.b13bs.Visible = False
peramalan.w11bs.Visible = False
peramalan.w21bs.Visible = False
peramalan.w31bs.Visible = False
peramalan.b01bs.Visible = False
peramalan.v11os.Visible = False
peramalan.v12os.Visible = False
peramalan.v13os.Visible = False
peramalan.v21os.Visible = False
peramalan.v220s.Visible = False
peramalan.v23os.Visible = False
peramalan.v31os.Visible = False
peramalan.v32os.Visible = False
peramalan.v33os.Visible = False
peramalan.b11os.Visible = False
peramalan.b12os.Visible = False
peramalan.b13os.Visible = False
peramalan.w11os.Visible = False
peramalan.w21os.Visible = False
peramalan.w31os.Visible = False
peramalan.b01os.Visible = False
peramalan.v11bi.Visible = True
peramalan.v12bi.Visible = True
peramalan.v21bi.Visible = True
peramalan.v22bi.Visible = True
peramalan.b11bi.Visible = True
peramalan.b12bi.Visible = True
peramalan.w11bi.Visible = True
peramalan.w21bi.Visible = True
peramalan.b01bi.Visible = True
peramalan.v11oi.Visible = True
peramalan.v12oi.Visible = True
peramalan.v21oi.Visible = True
peramalan.v22oi.Visible = True
peramalan.b11oi.Visible = True
peramalan.b12oi.Visible = True
peramalan.w11oi.Visible = True
peramalan.w21oi.Visible = True
peramalan.b01oi.Visible = True
peramalan.v11br.Visible = False
peramalan.v12br.Visible = False
peramalan.v13br.Visible = False
peramalan.v21br.Visible = False
peramalan.v22br.Visible = False
peramalan.v23br.Visible = False
peramalan.v31br.Visible = False
peramalan.v32br.Visible = False
peramalan.v33br.Visible = False
peramalan.b11br.Visible = False
peramalan.b12br.Visible = False
peramalan.b13br.Visible = False
peramalan.w11br.Visible = False
peramalan.w21br.Visible = False
peramalan.w31br.Visible = False
peramalan.b01br.Visible = False
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




End Sub

Private Sub Option3_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = True


peramalan.v11br.Visible = True
peramalan.v12br.Visible = True
peramalan.v13br.Visible = True
peramalan.v21br.Visible = True
peramalan.v22br.Visible = True
peramalan.v23br.Visible = True
peramalan.v31br.Visible = True
peramalan.v32br.Visible = True
peramalan.v33br.Visible = True
peramalan.b11br.Visible = True
peramalan.b12br.Visible = True
peramalan.b13br.Visible = True
peramalan.w11br.Visible = True
peramalan.w21br.Visible = True
peramalan.w31br.Visible = True
peramalan.b01br.Visible = True
peramalan.v11or.Visible = True
peramalan.v12or.Visible = True
peramalan.v13or.Visible = True
peramalan.v21or.Visible = True
peramalan.v22or.Visible = True
peramalan.v23or.Visible = True
peramalan.v31or.Visible = True
peramalan.v32or.Visible = True
peramalan.v33or.Visible = True
peramalan.b11or.Visible = True
peramalan.b12or.Visible = True
peramalan.b13or.Visible = True
peramalan.w11or.Visible = True
peramalan.w21or.Visible = True
peramalan.w31or.Visible = True
peramalan.b01or.Visible = True
peramalan.v11bs.Visible = False
peramalan.v12bs.Visible = False
peramalan.v13bs.Visible = False
peramalan.v21bs.Visible = False
peramalan.v22bs.Visible = False
peramalan.v23bs.Visible = False
peramalan.v31bs.Visible = False
peramalan.v32bs.Visible = False
peramalan.v33bs.Visible = False
peramalan.b11bs.Visible = False
peramalan.b12bs.Visible = False
peramalan.b13bs.Visible = False
peramalan.w11bs.Visible = False
peramalan.w21bs.Visible = False
peramalan.w31bs.Visible = False
peramalan.b01bs.Visible = False
peramalan.v11os.Visible = False
peramalan.v12os.Visible = False
peramalan.v13os.Visible = False
peramalan.v21os.Visible = False
peramalan.v220s.Visible = False
peramalan.v23os.Visible = False
peramalan.v31os.Visible = False
peramalan.v32os.Visible = False
peramalan.v33os.Visible = False
peramalan.b11os.Visible = False
peramalan.b12os.Visible = False
peramalan.b13os.Visible = False
peramalan.w11os.Visible = False
peramalan.w21os.Visible = False
peramalan.w31os.Visible = False
peramalan.b01os.Visible = False
peramalan.v11bi.Visible = False
peramalan.v12bi.Visible = False
peramalan.v21bi.Visible = False
peramalan.v22bi.Visible = False
peramalan.b11bi.Visible = False
peramalan.b12bi.Visible = False
peramalan.w11bi.Visible = False
peramalan.w21bi.Visible = False
peramalan.b01bi.Visible = False
peramalan.v11oi.Visible = False
peramalan.v12oi.Visible = False
peramalan.v21oi.Visible = False
peramalan.v22oi.Visible = False
peramalan.b11oi.Visible = False
peramalan.b12oi.Visible = False
peramalan.w11oi.Visible = False
peramalan.w21oi.Visible = False
peramalan.b01oi.Visible = False

End Sub

Private Sub Picture1_Click()
hslidentifikasi.Show

End Sub

