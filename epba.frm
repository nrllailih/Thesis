VERSION 5.00
Begin VB.Form epba 
   Caption         =   "ESTIMASI PARAMETER"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   Picture         =   "epba.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   6360
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   14
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   13
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   9
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Left            =   -2400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Frekuensi max"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   7
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Frekuensi min"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   6360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Koefisien Peningkan  Loudness"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   2280
      TabIndex        =   5
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Koefisien Peningkan  Pulse Rate"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pulse Rate Awal"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loudness Awal"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Iterasi"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Jumlah Kelelawar"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Width           =   3495
   End
End
Attribute VB_Name = "epba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jb As Integer, it As Integer, Al As Double, rl As Double, alf As Double, gamma As Double, fmin As Double, fmax As Double
Dim indeks As Integer, dar As Single, indekss As Integer, dar1 As Single, indeks1 As Double, dar2 As Single, indeks3 As Integer, dar3 As Single
Dim ss As Double, ii As Double, rr As Double
Dim miu As Double, prop As Double, beta As Double, alfa As Double, sum1 As Double
Dim huruf(), solusi() As Double
Dim i, j, p, k, a, c, l, d As Integer, iter As Integer
Dim k1, k2, k3, k4, l1, l2, l3, l4, m1, m2, m3, m4 As Double
Dim batls(100)  As Double, batnls(100) As Double, nilaiacak2 As Double, pbatgst(1000, 3) 'posisi sementara global step
Dim vbatgst(1000, 3) As Double 'kecepatan sgbiner
Dim fitbest As Double, gbest As Double, jbest As Integer, jnonbest As Integer, batfix As Double, gbestind As Integer, bzt(0, 4) As Double, gbest1 As Double, gbestind1 As Integer
Dim miu1 As Double, prop1 As Double, beta1 As Double, alfa1 As Double


Private Sub Command1_Click()

jb = Val(Text9.Text)
it = Val(Text2.Text)
Al = Val(Text3.Text)
rl = Val(Text4.Text)
alf = Val(Text5.Text)
gamma = Val(Text6.Text)
fmin = Val(Text7.Text)
fmax = Val(Text8.Text)

If Text9.Text = "" Then
MsgBox "Silahkan inputkan jumlah kelelawar!", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text9.SetFocus
Else
If Text9.Text < 2 Then
MsgBox "jumlah kelelawar harus lebih dari 2!", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text9.SetFocus
Else
If Text9.Text > 500 Then
MsgBox "Jumlah kelelawar tidak bisa lebih dari 500", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text9.SetFocus
Else

If Text2.Text = "" Then
MsgBox "Silahkan inputkan jumlah iterasi yang diinginkan.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text2.SetFocus
Else
If Text2.Text < 1 Then
MsgBox "Jumlah iterasi harus lebih dari satu.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text2.SetFocus
Else

If Text3.Text = "" Then
MsgBox "Silahkan inputkan nilai loudness awal.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text3.SetFocus
Else
If (Al < 0 Or Al > 1) Then
MsgBox "nilai loudness awal interval [0,1].", vbOKOnly, "PARAMETER TIDAK TEPAT"
Text3.SetFocus
Else

If Text4.Text = "" Then
MsgBox "Silahkan inputkan nilai pulserate awal.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else
If (rl < 0 Or rl > 1) Then
MsgBox "nilai pulserate awal interval [0,1].", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else

If Text5.Text = "" Then
MsgBox "Silahkan inputkan nilai koefisien peningkan pulse rate.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else
If (alfa < 0 Or alfa > 1) Then
MsgBox "nilai koefisien peningkatan pulse rate interval [0,1].", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else

If Text6.Text = "" Then
MsgBox "Silahkan inputkan nilai koefisien peningkatan loudness.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else
If (gamma < 0 Or gamma > 1) Then
MsgBox "nilai koefisien peningkatan loudness interval [0,1].", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else

If Text7.Text = "" Then
MsgBox "Silahkan inputkan frekuensi minimal.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else
If Text7.Text < 0 Then
MsgBox "nilai frekuensi minimal tidak boleh kurang dari 0", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else

If Text8.Text = "" Then
MsgBox "Silahkan inputkan frekuensi maksimal.", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else
If Text7.Text > 1000 Then
MsgBox "nilai frekuensi maksimal boleh lebih dari 1000", vbOKOnly, "PARAMETER TIDAK TEPAT"
Else
'input acces ke vb
hslep.Adodc1.Recordset.MoveFirst
For i = 0 To 83 '84 data
    dataasli(i, 0) = hslep.Adodc1.Recordset.Fields("No").Value
    dataasli(i, 1) = hslep.Adodc1.Recordset.Fields("S").Value
    dataasli(i, 2) = hslep.Adodc1.Recordset.Fields("I").Value
    dataasli(i, 3) = hslep.Adodc1.Recordset.Fields("R").Value
    dataasli(i, 4) = hslep.Adodc1.Recordset.Fields("N").Value
    hslep.Adodc1.Recordset.MoveNext
Next i
'normalisasi ba
hslep.Adodc3.Recordset.MoveFirst
For i = 0 To 83 '84 data
    normbat(i, 0) = hslep.Adodc3.Recordset.Fields("no").Value
    normbat(i, 1) = hslep.Adodc3.Recordset.Fields("s").Value
    normbat(i, 2) = hslep.Adodc3.Recordset.Fields("i").Value
    normbat(i, 3) = hslep.Adodc3.Recordset.Fields("r").Value
    hslep.Adodc3.Recordset.MoveNext
Next i
    

'bangkit populasi posisi kelelawar
indeks = 1
dar = 0
For i = 0 To jb - 1
    For j = 0 To 3 '4parameter
    Randomize
    dar = Rnd
    pbat(i, j) = Round(Rnd, 6)
    Next j
    pbat(i, 4) = indeks
    indeks = indeks + 1
    dar = 0
Next i

ss = normbat(0, 1)
ii = normbat(0, 2)
rr = normbat(0, 3)

'hitung solusi numerik
For i = 0 To jb - 1
    miu = pbat(i, 0)
    prop = pbat(i, 1)
    beta = pbat(i, 2)
    alfa = pbat(i, 3)
    For j = 0 To 71
   
        solusi = rungkutta(ss, ii, rr, miu, prop, beta, alfa)
        solusirk(i, j, 0) = solusi(0)
        solusirk(i, j, 1) = solusi(1)
        solusirk(i, j, 2) = solusi(2)
        ss = solusi(0)
        ii = solusi(1)
        rr = solusi(2)
    Next j
    ss = normbat(0, 1)
    ii = normbat(0, 2)
    rr = normbat(0, 3)
Next i

'MAE
indeks = 1
For i = 0 To jb - 1
    For j = 0 To 83
    sum1 = 0
        For p = 0 To 2
            eror(i, j, p) = errormae(normbat(i, p + 1), solusirk(i, j, p))
            sum1 = sum1 + eror(i, j, p)
        Next p
    eror(i, j, p) = sum1 / 3
    Next j
    eror(i, j, 4) = indeks
    indeks = indeks + 1
Next i

For i = 0 To jb - 1
    sum1 = 0
    For j = 0 To 83 '84 data
        sum1 = sum1 + eror(i, j, 3)
    Next j
    pbat(i, 5) = Round(sum1 / 84, 6) 'dibagi 72 data
Next i


For i = 0 To jb - 1
hslep.Adodc2.Recordset.AddNew
hslep.Adodc2.Recordset.Fields("ID") = pbat(i, 4)
hslep.Adodc2.Recordset.Fields("miu") = Round(pbat(i, 0), 6)
hslep.Adodc2.Recordset.Fields("prop") = Round(pbat(i, 1), 6)
hslep.Adodc2.Recordset.Fields("beta") = Round(pbat(i, 2), 6)
hslep.Adodc2.Recordset.Fields("alfa") = Round(pbat(i, 3), 6)
hslep.Adodc2.Recordset.Fields("fx") = pbat(i, 5)
hslep.Adodc2.Recordset.Update
hslep.DataGrid2.Refresh
Next i
 
'hslep.Adodc1.Recordset.MoveFirst
'hslep.Adodc2.Recordset.MoveFirst

'membuat matriks pulserate, loudness untuk pbat
For i = 0 To jb - 1
    pbat(i, 6) = rl
    pbat(i, 7) = Al
    pbat(i, 8) = alfa
    pbat(1, 9) = gamma
Next i


'iter = 0
'Do

'proses bat
'menghitung frekuensi
'For i = 0 To jb - 1
 '   frekuensi = rmsfrek(fmin, fmax)
'Next i
'membangkitkan kecepatan bat
indeks = 1
dar = 0
For i = 0 To jb - 1
    For j = 0 To 3    '4parameter
    Randomize
    dar = Rnd
    vbat(i, j) = Round(Rnd, 6)
    Next j
    vbat(i, 4) = indeks
    indeks = indeks + 1
    dar = 0
Next i
'membandingkan fungsi tujuan
For i = 0 To jb - 1
    For j = i + 1 To jb - 1
        If pbat(i, 5) < pbat(j, 5) Then
            'For k = 0 To 5
               gbest = pbat(i, 5)
               pbat(i, 5) = pbat(j, 5)
               pbat(j, 5) = gbest
               
               gbestind = pbat(i, 4)
               pbat(i, 4) = pbat(j, 4)
               pbat(j, 4) = gbestind
               
               
        End If

    Next j
Next i
For k = 0 To 3
    bzt(0, k) = pbat(gbestind - 1, k)
    'hslep.List2.AddItem (bzt(0, k))
Next k
'hslep.List1.AddItem (gbestind)
        
 

'membuat matriks posisi best(m)
For i = 0 To jb - 1
    For j = 0 To 3
    pbests(i, j) = bzt(0, k)

Next j
   
Next i

'menghitung frekuensi
'For i = 0 To jb - 1
 '   fminn(i, 0) = fmin
  '  fmaxx(i, 0) = fmax
'Next i
'membangkitkan nilai acak
indekss = 0
dar1 = 0
For i = 0 To jb - 1
    For j = 0 To 0
    Randomize
    dar1 = Rnd
    nilaiacak(i, 0) = Round(Rnd, 6)
    Next j
    nilaiacak(i, 1) = indekss
    indekss = indekss + 1
    dar1 = 0
    'hslep.List1.AddItem (nilaiacak(i, 0))
Next i
'membangkitkan nilai acak untuk update pulse rate dan loudness
indeks3 = 0
dar3 = 0
For i = 0 To jb - 1
    For j = 0 To 0
    Randomize
    dar3 = Rnd
    nilaiacak3(i, 0) = Round(Rnd, 6)
    Next j
    nilaiacak3(i, 1) = indeks3
    indeks3 = indeks3 + 1
    dar3 = 0
    'hslep.List1.AddItem (nilaiacak(i, 0))
Next i

'dar = 0
'indeks = 1
'For i = 0 To jb - 1
 '   Randomize
  '  dar = Rnd
   'nilaiacak2(i) = Round(Rnd, 6)
    'dar = 0
    'indeks = indeks + 1
'Next i
For i = 0 To jb - 1
    For j = 0 To 0
     fminn(i, 0) = fmin
     fmaxx(i, 0) = fmax
     frekuensi(i, 0) = fminn(i, 0) + ((fmaxx(i, 0) - fminn(i, 0)) * nilaiacak(i, 0))
     Next j
     'hslep.List1.AddItem (frekuensi(i, 0))
Next i
'menghitung kecepatan baru kelelawar
For i = 0 To jb - 1
    For j = 0 To 3
        vbatgst(i, j) = vbat(i, j) + (pbests(i, j) - pbat(i, j)) * frekuensi(i, 0)
        vbbat(i, j) = Sgbiner(vbatgst(i, j))
        'hslep.List1.AddItem (vbbat(i, j))
    Next j
    
Next i
iter = 0
Do
iter = iter + 1
'menghitung posisi baru
For i = 0 To jb - 1
    For j = 0 To 3
        pbbatg(i, j) = pbat(i, j) + vbbat(i, j)
        'pbbatg(i, j) = Sgbiner(pbbatg(i, j))
        If pbbatg(i, j) > 0 Or pbbatg(i, j) < 0 Then
            pbbatg(i, j) = Sgbiner(pbbatg(i, j))
        End If
        'hslep.List1.AddItem (Round((pbatgst(i, j)), 6))
    Next j
Next i

'menghitung solusi numerik
For i = 0 To jb - 1
    miu1 = pbbatg(i, 0)
    prop1 = pbbatg(i, 1)
    beta1 = pbbatg(i, 2)
    alfa1 = pbbatg(i, 3)
    For j = 0 To 71
        solusi = rungkutta(ss, ii, rr, miu1, prop1, beta1, alfa1)
        solusirk1(i, j, 0) = solusi(0)
        solusirk1(i, j, 1) = solusi(1)
        solusirk1(i, j, 2) = solusi(2)
        ss = solusi(0)
        ii = solusi(1)
        rr = solusi(2)
    Next j
    ss = normbat(0, 1)
    ii = normbat(0, 2)
    rr = normbat(0, 3)
Next i

'MAE
indeks = 1
For i = 0 To jb - 1
    For j = 0 To 71
    sum1 = 0
        For p = 0 To 2
            eror(i, j, p) = errormae(normbat(i, p + 1), solusirk1(i, j, p))
            sum1 = sum1 + eror(i, j, p)
        Next p
    eror(i, j, p) = sum1 / 3
    Next j
    eror(i, j, 4) = indeks
    indeks = indeks + 1
Next i

For i = 0 To jb - 1
    sum1 = 0
    For j = 0 To 83 '84 data
        sum1 = sum1 + eror(i, j, 3)
    Next j
    pbbatg(i, 5) = Round(sum1 / 84, 6) 'dibagi 72 data
    'hslep.List1.AddItem (pbbatg(i, 5))
Next i
'membuat matriks pulserate, loudness untuk pbbatg
For i = 0 To jb - 1
    pbbatg(i, 6) = rl
    pbbatg(i, 7) = Al
    pbbatg(i, 8) = alfa
    pbbatg(1, 9) = gamma
Next i


'proses localsearch
'jumlah kelelawar yang melakukan localsearch dan non localsearch
jbest = 0
jnonbest = 0
For i = 0 To jb - 1
    If pbbatg(i, 6) < nilaiacak(i, 0) Then
        jbest = jbest + 1
        jnonbest = jb - jbest
        End If
    
'hslep.List1.AddItem (pbbatg(i, 6))
'hslep.List2.AddItem (jnonbest)
'hslep.List3.AddItem (nilaiacak(i, 0))
Next i


'membangkitkan matriks epsilon
indeks1 = 1
dar2 = 0
For i = 0 To jbest
    For j = 0 To 3
        Randomize
        dar = Rnd
        epsilon(i, j) = Round(Rnd, 6)
    Next j
        dar2 = 0
Next i


'kelelawar yang melakukan localsearch
a = 0
c = 0

For i = 0 To jb - 1
        If pbbatg(i, 6) < nilaiacak(i, 0) Then
         batls(a) = nilaiacak(i, 1)
         hslep.List4.AddItem (batls(a))
            For j = 0 To 3
                If pbat(batls(a), 5) < pbbatg(batls(a), 5) Then
                    pbbatls(batls(a), j) = pbat(batls(a), j) + epsilon(batls(a), j) * pbbatg(batls(a), j)
                         If pbat(batls(a), 5) > pbbatg(batls(a), 5) Then
                             pbbatls(batls(a), j) = pbbatg(batls(a), j) + epsilon(batls(a), j) * pbbatg(batls(a), j)
                If pbbatls(batls(a), j) > 1 Or pbbatls(batls(a), j) < 0 Then
                    pbbatls(batls(a), j) = Sgbiner(pbbatls(batls(a), j))
                End If
                End If
                End If
            Next j
        
        
        

        
        'End If
        

'Next i
'hitung solusi numerik
ss = normbat(0, 1)
ii = normbat(0, 2)
rr = normbat(0, 3)
miu = pbbatls(batls(a), 0)
prop = pbbatls(batls(a), 1)
beta = pbbatls(batls(a), 2)
alfa = pbbatls(batls(a), 3)

For k = 0 To 83
    solusi = rungkutta(ss, ii, rr, miu, prop, beta, alfa)
    solusirk(batls(a), k, 0) = solusi(0)
    solusirk(batls(a), k, 1) = solusi(1)
    solusirk(batls(a), k, 2) = solusi(2)
    ss = solusi(0)
    ii = solusi(1)
    rr = solusi(2)
Next k
'hitung fungsi tujuan MAE
For j = 0 To 83
sum1 = 0
    For k = 0 To 2
        eror(batls(a), j, k) = errormae(normbat(j, k + 1), solusirk(batls(a), j, k))
        sum1 = sum1 + eror(batls(a), j, k)
    Next k
eror(batls(a), j, 3) = sum1 / 3
Next j
sum1 = 0
For j = 0 To 83
    sum1 = sum1 + eror(batls(a), j, k)
Next j
pbbatls(batls(a), 5) = sum1 / 84


'update pulse rate dan loudness untuk localsearch

    For j = 0 To 3
    If pbbatls(batls(a), 5) < pbbatg(batls(a), 5) And nilaiacak(batls(a), 0) < pbbatg(batls(a), 7) Then
        pbf(batls(a), j) = pbbatls(batls(a), j)
        pbf(batls(a), 5) = pbbatls(batls(a), 5)
        pbf(batls(a), 7) = pbbatg(batls(a), 8) * pbbatg(batls(a), 7)
        pbf(batls(a), 6) = pbbatg(batls(a), 6) * (1 - (Exp(-pbbatg(batls(a), 9) * iter)))
        Else
            pbf(batls(a), j) = pbbatg(batls(a), j)
            pbf(batls(a), 5) = pbbatg(batls(a), 5)
            pbf(batls(a), 7) = pbbatg(batls(a), 7)
            pbf(batls(a), 6) = pbbatg(batls(a), 6)
    End If
    Next j
    hslep.List1.AddItem (pbf(batls(a), 5))
    
    
End If
    
        
        
Next i

'update pulserate dan loudness nls
For i = 0 To jb - 1
If pbbatg(i, 6) > nilaiacak(i, 0) Then
batnls(c) = nilaiacak(i, 1)
'hslep.List3.AddItem (batnls(c))

 For p = 0 To 3
    If pbbatg((batnls(c)), 5) < pbat((batnls(c)), 5) And nilaiacak(batls(a), 0) < pbbatg(batls(a), 7) Then
        
            pbf(batnls(c), 7) = pbbatg(batnls(c), 8) * pbbatg((batnls(c)), 7)
            pbf(batnls(c), 6) = pbbatg(i, 6) * (1 - (Exp(-pbbatg(batnls(c), 9) * iter)))
            pbf(batnls(c), p) = pbbatg((batnls(c)), p)
            pbf(batnls(c), 5) = pbbatg(batnls(c), 5)
        Else
        pbf(batnls(c), p) = pbat(batnls(c), p)
        pbf(batnls(c), 7) = pbat(batnls(c), 7)
        pbf(batnls(c), 6) = pbat(batnls(c), 6)
        pbf(batnls(c), 5) = pbat(batnls(c), 5)
    End If
  Next p
hslep.List2.AddItem (pbf(batnls(c), 5))
End If

'hslep.List1.AddItem (pbbatls(batls(a), 5))
'hslep.List2.AddItem (pbbatls(batnls(c), 5))
'hslep.List3.AddItem (jbest)
 
Next i

    
'***************************************************
'membanding fungsi tujuan
For i = 0 To jb - 1
    For j = 0 To jb - 1
        If pbf(i, 5) <= pbf(j, 5) Then
            'For k = 0 To 5
               gbest1 = pbf(i, 5)
               pbf(i, 5) = pbf(j, 5)
               pbf(j, 5) = gbest1
               
               gbestind1 = pbf(i, 4)
               pbf(i, 4) = pbf(j, 4)
               pbf(j, 4) = gbestind1
               
         
        End If

    Next j
       'hslep.List2.AddItem (gbest1)
       'hslep.List1.AddItem (gbestind1)
       hslep.List3.AddItem (gbest1)
Next i

For k = 0 To 7
    bzt1(0, k) = pbf(gbestind1, k)
 
Next k

Loop Until iter = it




miufix = Round(bzt1(0, 0), 6)
propfix = Round(bzt1(0, 1), 6)
betafix = Round(bzt1(0, 2), 6)
alfafix = Round(bzt1(0, 3), 6)
fxfix = Round(bzt1(0, 5), 6)
alfix = Round(bzt1(0, 7), 6)
rlfix = Round(bzt1(0, 6), 6)
batke = Round(bzt1(0, 4), 1)


hslep.Label12.Caption = jb
hslep.Label6.Caption = miufix
hslep.Label15.Caption = propfix
hslep.Label16.Caption = betafix
hslep.Label17.Caption = alfafix
hslep.Label18.Caption = rlfix
hslep.Label19.Caption = alfix
hslep.Label20.Caption = fmin
hslep.Label21.Caption = fmax
hslep.Label22.Caption = fxfix





hslep.Adodc1.Recordset.MoveFirst
hslep.Adodc3.Recordset.MoveFirst
hslep.Adodc2.Recordset.MoveFirst
'hslep.Adodc4.Recordset.MoveFirst

hslep.Show
epba.Hide
    
    

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


End Sub

Function rungkutta(ss As Double, ii As Double, rr As Double, miu As Double, prop As Double, beta As Double, alfa As Double) As Double()
Dim k1 As Double, k2 As Double, k3 As Double
Dim l1 As Double, l2 As Double, l3 As Double
Dim m1 As Double, m2 As Double, m3 As Double
Dim k As Double, l As Double, m As Double, outp(3) As Double

'h=1
k1 = Susceptible(ss, ii, prop, miu, beta)
l1 = Infected(ss, ii, beta, miu, alfa)
m1 = Recovery(ii, rr, prop, miu, alfa)

k2 = Susceptible(ss + (k1 / 2), ii + (l1 / 2), prop, miu, beta)
l2 = Infected(ss + (k1 / 2), ii + (l1 / 2), beta, miu, alfa)
m2 = Recovery(ii + (l1 / 2), rr + (m1 / 2), prop, miu, alfa)

k3 = Susceptible(ss + (k2 / 2), ii + (l2 / 2), prop, miu, beta)
l3 = Infected(ss + (k2 / 2), ii + (l2 / 2), beta, miu, alfa)
m3 = Recovery(ii + (l2 / 2), rr + (m2 / 2), prop, miu, alfa)

k1 = Susceptible(ss + k3, ii + l3, prop, miu, beta)
l1 = Infected(ss + k3, ii + l3, beta, miu, alfa)
m1 = Recovery(ii + l3, rr + m3, prop, miu, alfa)

k = ss + (k1 + (2 * k2) + (2 * k3) + k4) / 6
l = ss + (l1 + (2 * l2) + (2 * l3) + l4) / 6
m = ss + (k1 + (2 * k2) + (2 * k3) + k4) / 6

outp(0) = k
outp(1) = l
outp(2) = m

rungkutta = outp

End Function

Function Susceptible(ss As Double, ii As Double, prop As Double, miu As Double, beta As Double) As Double
Susceptible = (1 - prop) * miu - miu * ss - beta * ss * ii

End Function

Function Infected(ss As Double, ii As Double, beta As Double, miu As Double, alfa As Double)
Infected = beta * ss - miu * ii - alfa * ii
End Function

Function Recovery(ii As Double, rr As Double, alfa As Double, prop As Double, miu As Double) As Double
Recovery = alfa * ii + prop * miu - miu * rr

End Function

