VERSION 5.00
Begin VB.Form WELCOME 
   Caption         =   "WELCOME"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   10680
      Picture         =   "welcome.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   7200
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   7320
      Picture         =   "welcome.frx":2D24
      Top             =   5280
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   8250
      Left            =   0
      Picture         =   "welcome.frx":3BA1A
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "WELCOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Dim con As New ADODB.Connection
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\programku\program\database.mdb;Persist Security Info=False")
con.Execute ("DELETE FROM popawalbat;")
con.Execute ("DELETE FROM jstmse;")
con.Execute ("DELETE FROM denoriden;")
con.Execute ("DELETE FROM denorup;")
con.Execute ("DELETE FROM snum;")




epba.Show
WELCOME.Hide
End Sub

Private Sub Picture1_Click()
End
End Sub
