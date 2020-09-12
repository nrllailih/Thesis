VERSION 5.00
Begin VB.Form inidentify 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
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
      Left            =   3840
      TabIndex        =   18
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "INISIALISASI BACKPRO"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1200
      TabIndex        =   11
      Top             =   4440
      Width           =   4455
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2040
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000B&
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
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000B&
         Caption         =   "Batas Iterasi"
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000B&
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
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Hasil Estimasi Parameter"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Label10"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Label9"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Label8"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Label7"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Label6"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "MAE"
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
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "MIU"
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
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "ALFA"
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
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "BETA"
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
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "PROP"
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   8250
      Left            =   -3000
      Picture         =   "identify.frx":0000
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "inidentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim indeks As Integer
Dim cum1 As Double, cum2 As Double, cum3 As Double, cum4 As Double, cum5 As Double



learning = Val(Text1.Text)
maksiter = Val(Text2.Text)
makseror = Val(Text3.Text)

If (Text1.Text = "") Then
MsgBox "Silahkan masukkan nilai learning rate [0,1]", vbOKOnly, "Nilai Parameter!"
Else
If (learning <= 0 Or learning >= 1 Or IsNumeric(Text1.Text) = False) Then
MsgBox "Silahkan isikan lagi learning rate bernilai (0,1)", vbOKOnly, "Parameter tidak tepat!"
Text1.SetFocus
Else
If (Text2.Text = "") Then
MsgBox "silahkan masukkan jumlah iterasi", vbOKOnly, "Parameter tidak tepat"
Else
If Text2.Text < 1 Or IsNumeric(Text2.Text) = False Then
MsgBox "ULANGI! jumlah iterasi bernilai positif", vbOKOnly, "Parameter tidak tepat"
Text2.SetFocus
Else

If Text2.Text = "" Then
MsgBox "silahkan masukkan batas error", vbOKOnly, "Parameter tidak tepat"
Else
If (makserror < 0 Or IsNumeric(Text3.Text) = False) Then
MsgBox "ULANGI! batas error bernilai positif", vbOKOnly, "Parameter tidak tepat"
Text3.SetFocus
Else

'input data acces ke vb
hslidentifikasi.Adodc1.Recordset.MoveFirst
For i = 0 To 83
    dataasli(i, 0) = hslidentifikasi.Adodc1.Recordset.Fields("No").Value
    dataasli(i, 1) = hslidentifikasi.Adodc1.Recordset.Fields("S").Value
    dataasli(i, 2) = hslidentifikasi.Adodc1.Recordset.Fields("I").Value
    dataasli(i, 3) = hslidentifikasi.Adodc1.Recordset.Fields("R").Value
    dataasli(i, 4) = hslidentifikasi.Adodc1.Recordset.Fields("N").Value
    hslidentifikasi.Adodc1.Recordset.MoveNext
Next i

'normalisasi data
'untuk M
For i = 0 To 83
    dataasli(i, 5) = dataasli(i, 1) * dataasli(i, 2) / (dataasli(i, 1) + dataasli(i, 2) + dataasli(i, 3))
Next i

For i = 1 To 5 '(S/I/R/N/M)
    maks(i) = dataasli(0, i)
    min(i) = dataasli(0, i)
Next i

For i = 1 To 5 '(S/I/R/N/M)
    For j = 0 To 71 '72 data
        If maks(i) < dataasli(j, i) Then
            maks(i) = dataasli(j, i)
        End If
        If min(i) > dataasli(j, i) Then
            min(i) = dataasli(j, i)
        End If
    Next j
    'hslidentifikasi.List1.AddItem (min(i))
Next i

For i = 1 To 5
    For j = 0 To 83
        normalisasi(j, i) = Round(-1 + ((2 * (dataasli(j, i) - min(i))) / (maks(i) - min(i))), 6)
       
    Next j
Next i

'membuat pola identifikasi
For i = 0 To 2 '3 populasi (S/R)
    For j = 0 To 83
        poladatain(i, j, 0) = j
    Next j
    If i = 0 Then
        For j = 0 To 71 'pola s
            poladatain(i, j, 1) = normalisasi(j, 1)
            poladatain(i, j, 2) = normalisasi(j, 4)
            poladatain(i, j, 3) = normalisasi(j, 5)
            poladatain(i, j, 4) = normalisasi(j, 1)
           
        Next j
        
        Else
             If i = 1 Then
                For j = 0 To 83 'pola r
                    poladatain(i, j, 1) = normalisasi(j, 3)
                    poladatain(i, j, 2) = normalisasi(j, 2)
                    poladatain(i, j, 3) = normalisasi(j, 4)
                    poladatain(i, j, 4) = normalisasi(j, 3)
                
                Next j
            Else
                For j = 0 To 83 'pola i
                    poladatain(i, j, 1) = normalisasi(j, 2)
                    poladatain(i, j, 2) = normalisasi(j, 5)
                    poladatain(i, j, 3) = normalisasi(j, 2)

                Next j
            End If
    End If
Next i

'bobot bias awal
For i = 0 To 2 Step 1
    If i = 0 Then
        'populasi S
        bbin(i, 1) = 1 - miufix
        bbin(i, 2) = (1 - propfix) * miufix
        bbin(i, 3) = betafix
        bbin(i, 4) = 1 - miufix
        bbin(i, 5) = (1 - propfix) * miufix
        bbin(i, 6) = betafix
        bbin(i, 7) = 1 - miufix
        bbin(i, 8) = (1 - propfix) * miufix
        bbin(i, 9) = betafix
            For k = 10 To 16
                bbin(i, k) = RandomAntara(-1, 1)
            Next k
    Else
        If i = 1 Then
        'populasi R
            bbin(i, 1) = 1 - miufix
            bbin(i, 2) = alfafix
            bbin(i, 3) = propfix * miufix
            bbin(i, 4) = 1 - miufix
            bbin(i, 5) = alfafix
            bbin(i, 6) = propfix * miufix
            bbin(i, 7) = 1 - miufix
            bbin(i, 8) = alfafix
            bbin(i, 9) = propfix * miufix
                For k = 10 To 16
                    bbin(i, k) = RandomAntara(-1, 1)
                Next k
        Else
            'populasi I
            cum3 = 1 - miufix
            bbin(i, 1) = 1 - cum3
             If bbin(i, 1) > 1 Or bbin(i, 1) < -1 Then
            bbin(i, 1) = FsAktivasi(bbin(i, 1))
            End If
            bbin(i, 2) = betafix
            bbin(i, 4) = 1 - cum3
            If bbin(i, 4) > 1 Or bbin(i, 1) < -1 Then
            bbin(i, 1) = FsAktivasi(bbin(i, 1))
            End If
            bbin(i, 5) = betafix
            For k = 10 To 11
                bbin(i, k) = RandomAntara(-1, 1)
            Next k
            For k = 13 To 16
                bbin(i, k) = RandomAntara(-1, 1)
            Next k
        End If
    End If
Next i

hslidentifikasi.v11as.Caption = Round(bbin(0, 1), 6)
hslidentifikasi.v12as.Caption = Round(bbin(0, 4), 6)
hslidentifikasi.v13as.Caption = Round(bbin(0, 7), 6)
hslidentifikasi.v21as.Caption = Round(bbin(0, 2), 6)
hslidentifikasi.v22as.Caption = Round(bbin(0, 5), 6)
hslidentifikasi.v23as.Caption = Round(bbin(0, 8), 6)
hslidentifikasi.v31as.Caption = Round(bbin(0, 3), 6)
hslidentifikasi.v32as.Caption = Round(bbin(0, 6), 6)
hslidentifikasi.v33as.Caption = Round(bbin(0, 9), 6)
hslidentifikasi.v01as.Caption = Round(bbin(0, 10), 6)
hslidentifikasi.v02as.Caption = Round(bbin(0, 11), 6)
hslidentifikasi.v03as.Caption = Round(bbin(0, 12), 6)
hslidentifikasi.w11as.Caption = Round(bbin(0, 13), 6)
hslidentifikasi.w21as.Caption = Round(bbin(0, 14), 6)
hslidentifikasi.w31as.Caption = Round(bbin(0, 15), 6)
hslidentifikasi.w01as.Caption = Round(bbin(0, 16), 6)
hslidentifikasi.v11ai.Caption = Round(bbin(2, 1), 6)
hslidentifikasi.v12ai.Caption = Round(bbin(2, 4), 6)
hslidentifikasi.v21ai.Caption = Round(bbin(2, 2), 6)
hslidentifikasi.v22ai.Caption = Round(bbin(2, 5), 6)
hslidentifikasi.v01ai.Caption = Round(bbin(2, 10), 6)
hslidentifikasi.v02ai.Caption = Round(bbin(2, 11), 6)
hslidentifikasi.w11ai.Caption = Round(bbin(2, 13), 6)
hslidentifikasi.w21ai.Caption = Round(bbin(2, 14), 6)
hslidentifikasi.w01ai.Caption = Round(bbin(2, 16), 6)
hslidentifikasi.v11ar.Caption = Round(bbin(1, 1), 6)
hslidentifikasi.v12ar.Caption = Round(bbin(1, 4), 6)
hslidentifikasi.v13ar.Caption = Round(bbin(1, 7), 6)
hslidentifikasi.v21ar.Caption = Round(bbin(1, 2), 6)
hslidentifikasi.v22ar.Caption = Round(bbin(1, 5), 6)
hslidentifikasi.v23ar.Caption = Round(bbin(1, 8), 6)
hslidentifikasi.v31ar.Caption = Round(bbin(1, 3), 6)
hslidentifikasi.v32ar.Caption = Round(bbin(1, 6), 6)
hslidentifikasi.v33ar.Caption = Round(bbin(1, 9), 6)
hslidentifikasi.v01ar.Caption = Round(bbin(1, 10), 6)
hslidentifikasi.v02ar.Caption = Round(bbin(1, 11), 6)
hslidentifikasi.v03ar.Caption = Round(bbin(1, 12), 6)
hslidentifikasi.w11ar.Caption = Round(bbin(1, 13), 6)
hslidentifikasi.w21ar.Caption = Round(bbin(1, 14), 6)
hslidentifikasi.w31ar.Caption = Round(bbin(1, 15), 6)
hslidentifikasi.w01ar.Caption = Round(bbin(1, 16), 6)
poh = 0
Do
'feedforward, bckpro, update poulasi s dan r
For i = 0 To 1 'populasi S dan R
    For j = 0 To 83
        'feedforward
        cum1 = 0
        cum2 = 0
        cum4 = 0
        For k = 1 To 3 '3inputan
            cum1 = cum1 + bbin(i, k) * poladatain(i, j, k)
            cum2 = cum2 + bbin(i, k + 3) * poladatain(i, j, k)
            cum3 = cum3 + bbin(i, k + 6) * poladatain(i, j, k)
        Next k
        FFin(i, 0) = bbin(i, 10) + cum1
        FFin(i, 1) = bbin(i, 11) + cum2
        FFin(i, 2) = bbin(i, 12) + cum4
        FFin(i, 3) = FsAktivasi(FFin(i, 0))
        FFin(i, 4) = FsAktivasi(FFin(i, 1))
        FFin(i, 5) = FsAktivasi(FFin(i, 2))
        cum5 = 0
        For k = 1 To 3 'hidden layer ada 3
            cum5 = cum5 + bbin(i, k + 12) * FFin(i, k + 1)
        Next k
        FFin(i, 6) = bbin(i, 16) + cum5
        outin(i, j) = FsAktivasi(FFin(i, 6))
        ermse(i, j, 0) = j
        ermse(i, j, 1) = MSE(poladatain(i, j, 4), outin(i, j))
        hslidentifikasi.List2.AddItem (ermse(0, j, 1))
        hslidentifikasi.List1.AddItem (ermse(1, j, 1))
        'Backpropagation
        BPin(i, 0) = (poladatain(i, j, 4) - outin(i, j)) * Fsaksaktivasi(FFin(i, 6))
        BPin(i, 1) = learning * BPin(i, 0) * FFin(i, 3)
        BPin(i, 2) = learning * BPin(i, 0) * FFin(i, 4)
        BPin(i, 3) = learning * BPin(i, 0) * FFin(i, 5)
        BPin(i, 4) = learning * BPin(i, 0)
        BPin(i, 5) = BPin(i, 0) * bbin(i, 13)
        BPin(i, 6) = BPin(i, 0) * bbin(i, 14)
        BPin(i, 7) = BPin(i, 0) * bbin(i, 15)
        BPin(i, 8) = BPin(i, 5) * Fsaksaktivasi(FFin(i, 0))
        BPin(i, 9) = BPin(i, 6) * Fsaksaktivasi(FFin(i, 1))
        BPin(i, 10) = BPin(i, 7) * Fsaksaktivasi(FFin(i, 2))
        BPin(i, 11) = learning * BPin(i, 8) * poladatain(i, j, 1)
        BPin(i, 12) = learning * BPin(i, 8) * poladatain(i, j, 2)
        BPin(i, 13) = learning * BPin(i, 8) * poladatain(i, j, 3)
        BPin(i, 14) = learning * BPin(i, 9) * poladatain(i, j, 1)
        BPin(i, 15) = learning * BPin(i, 9) * poladatain(i, j, 2)
        BPin(i, 16) = learning * BPin(i, 9) * poladatain(i, j, 3)
        BPin(i, 17) = learning * BPin(i, 10) * poladatain(i, j, 1)
        BPin(i, 18) = learning * BPin(i, 10) * poladatain(i, j, 2)
        BPin(i, 19) = learning * BPin(i, 10) * poladatain(i, j, 3)
        BPin(i, 20) = learning * BPin(i, 8)
        BPin(i, 21) = learning * BPin(i, 9)
        BPin(i, 22) = learning * BPin(i, 10)
        'update bobot dan bias
        For k = 1 To 12
        bbin(i, k) = bbin(i, k) + BPin(i, k + 10)
        Next k
        For k = 1 To 4
        bbin(i, k + 12) = bbin(i, k + 12) + BPin(i, k)
        Next k
        
    Next j
'hslidentifikasi.List1.AddItem (outin(1, i))
Next i

'feedforward, backpropagatio error, update bobot dan bias
For i = 2 To 2
    For j = 0 To 83
    'feedforward
    cum1 = 0
    cum2 = 0
    For k = 1 To 2 'inputan ada 2
        cum1 = cum1 + bbin(i, k) * poladatain(i, j, k)
        cum2 = cum2 + bbin(i, k + 3) * poladatain(i, j, k)
    Next k
    FFin(i, 0) = bbin(i, 10) + cum1
    FFin(i, 1) = bbin(i, 11) + cum2
    FFin(i, 3) = Fsaksaktivasi(FFin(i, 0))
    FFin(i, 4) = Fsaksaktivasi(FFin(i, 1))
    cum4 = 0
    For k = 1 To 2 'hidden layer ada 2
        cum4 = cum4 + bbin(i, k + 12) * FFin(i, k + 1)
    Next k
    FFin(i, 6) = bbin(i, 16) + cum4
    outin(i, j) = FsAktivasi(FFin(i, 6))
    ermse(i, j, 0) = j
    ermse(i, j, 1) = MSE(poladatain(i, j, 3), outin(i, j))
    hslidentifikasi.List3.AddItem (ermse(2, j, 1))
    'Backpropagation
    BPin(i, 0) = (poladatain(i, j, 3) - outin(i, j)) * Fsaksaktivasi(FFin(i, 6))
    BPin(i, 2) = learning * BPin(i, 0) * FFin(i, 1)
    BPin(i, 4) = learning * BPin(i, 0)
    BPin(i, 5) = BPin(i, 0) * bbin(i, 13)
    BPin(i, 6) = BPin(i, 0) * bbin(i, 14)
    BPin(i, 8) = BPin(i, 5) * Fsaksaktivasi(FFin(i, 0))
    BPin(i, 9) = BPin(i, 6) * Fsaksaktivasi(FFin(i, 1))
    BPin(i, 11) = learning * BPin(i, 8) * poladatain(i, j, 1)
    BPin(i, 12) = learning * BPin(i, 8) * poladatain(i, j, 2)
    BPin(i, 14) = learning * BPin(i, 9) * poladatain(i, j, 1)
    BPin(i, 15) = learning * BPin(i, 9) * poladatain(i, j, 2)
    BPin(i, 20) = learning * BPin(i, 8)
    BPin(i, 21) = learning * BPin(i, 9)
    'update bobot dan bias
    For k = 1 To 2
    bbin(i, k) = bbin(i, k) + BPin(i, k + 10)
    bbin(i, k + 3) = bbin(i, k + 3) + BPin(i, k + 13)
    bbin(i, k + 9) = bbin(i, k + 9) + BPin(i, k + 19)
    bbin(i, k + 12) = bbin(i, k + 12) + BPin(i, k + 1)
    Next k
    bbin(i, 16) = bbin(i, 16) + BPin(i, 4)
    Next j
Next i
For i = 0 To 2
For j = 0 To 83
    For k = 1 To 16
        If bbin(i, k) > 1 Or bbin(i, k) < -1 Then
            bbin(i, k) = FsAktivasi(bbin(i, k))
        End If
    Next k
Next j
Next i


'MSE akhir
For i = 0 To 2
    cum1 = 0
    For j = 0 To 83
        cum1 = cum1 + ermse(i, j, 1)
    Next j
    ermse(i, 0, 2) = cum1 / 84
Next i

cum1 = 0
For i = 0 To 2
    cum1 = cum1 + ermse(i, 0, 2)
Next i


err = Round(cum1 / 3, 6)
poh = poh + 1

hslidentifikasi.Adodc2.Recordset.AddNew
hslidentifikasi.Adodc2.Recordset.Fields("iterasi") = poh
hslidentifikasi.Adodc2.Recordset.Fields("MSE") = err
hslidentifikasi.Adodc2.Recordset.Update
hslidentifikasi.DataGrid2.Refresh

Loop Until poh = maksiter Or err <= makseror
    
hslidentifikasi.vbs11.Caption = Round(bbin(0, 1), 6)
hslidentifikasi.vbs12.Caption = Round(bbin(0, 4), 6)
hslidentifikasi.vbs13.Caption = Round(bbin(0, 7), 6)
hslidentifikasi.vbs21.Caption = Round(bbin(0, 2), 6)
hslidentifikasi.vbs22.Caption = Round(bbin(0, 5), 6)
hslidentifikasi.vbs23.Caption = Round(bbin(0, 8), 6)
hslidentifikasi.vbs31.Caption = Round(bbin(0, 3), 6)
hslidentifikasi.vbs32.Caption = Round(bbin(0, 6), 6)
hslidentifikasi.vbs33.Caption = Round(bbin(0, 9), 6)
hslidentifikasi.vbs01.Caption = Round(bbin(0, 10), 6)
hslidentifikasi.vbs02.Caption = Round(bbin(0, 11), 6)
hslidentifikasi.vbs03.Caption = Round(bbin(0, 12), 6)
hslidentifikasi.wbs11.Caption = Round(bbin(0, 13), 6)
hslidentifikasi.wbs21.Caption = Round(bbin(0, 14), 6)
hslidentifikasi.wbs31.Caption = Round(bbin(0, 15), 6)
hslidentifikasi.wbs01.Caption = Round(bbin(0, 16), 6)
hslidentifikasi.vbr11.Caption = Round(bbin(1, 1), 6)
hslidentifikasi.vbr12.Caption = Round(bbin(1, 4), 6)
hslidentifikasi.vbr13.Caption = Round(bbin(1, 7), 6)
hslidentifikasi.vbr21.Caption = Round(bbin(1, 2), 6)
hslidentifikasi.vbr22.Caption = Round(bbin(1, 5), 6)
hslidentifikasi.vbr23.Caption = Round(bbin(1, 8), 6)
hslidentifikasi.vbr31.Caption = Round(bbin(1, 3), 6)
hslidentifikasi.vbr32.Caption = Round(bbin(1, 6), 6)
hslidentifikasi.vbr33.Caption = Round(bbin(1, 9), 6)
hslidentifikasi.vbr01.Caption = Round(bbin(1, 10), 6)
hslidentifikasi.vbr02.Caption = Round(bbin(1, 11), 6)
hslidentifikasi.vbr03.Caption = Round(bbin(1, 12), 6)
hslidentifikasi.wbr11.Caption = Round(bbin(1, 13), 6)
hslidentifikasi.wbr21.Caption = Round(bbin(1, 14), 6)
hslidentifikasi.wbr31.Caption = Round(bbin(1, 15), 6)
hslidentifikasi.wbr01.Caption = Round(bbin(1, 16), 6)
hslidentifikasi.vbi11.Caption = Round(bbin(2, 1), 6)
hslidentifikasi.vbi12.Caption = Round(bbin(2, 4), 6)
hslidentifikasi.vbi21.Caption = Round(bbin(2, 2), 6)
hslidentifikasi.vbi22.Caption = Round(bbin(2, 5), 6)
hslidentifikasi.vbi01.Caption = Round(bbin(2, 10), 6)
hslidentifikasi.vbi02.Caption = Round(bbin(2, 11), 6)
hslidentifikasi.wbi11.Caption = Round(bbin(2, 13), 6)
hslidentifikasi.wbi21.Caption = Round(bbin(2, 14), 6)
hslidentifikasi.wbi01.Caption = Round(bbin(2, 16), 6)



hslidentifikasi.Option1.Value = True
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

hslidentifikasi.Show


        
        
        


    
            
            
            








End If
End If
End If
End If
End If
End If





End Sub

