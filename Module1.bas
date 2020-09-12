Attribute VB_Name = "Module1"
Option Explicit
Global con As ADODB.Connection
Global rs As ADODB.Recordset
Global dataasli(100, 10) As Double
Global normbat(100, 3) As Double
Global nomrl(100, 3) As Double
Global pbat(1000, 9) As Double 'batke, (prop+miu+beta+alfa+indeks+fx+pulserate+loudness+alfa+gamma)
Global vbat(1000, 5) As Double 'bat ke,(prop+miu+beta+alfa)
Global pbbat(1000, 9) As Double 'batke, (prop+miu+beta+alfa+indeks+fx+pulserate+loudness+alfa+gamma)
Global vbbat(1000, 4) As Double 'bat ke,(prop+miu+beta+alfa+indeks)
Global pbests(1000, 9) As Double '(prop+miu+beta+alfa+indeks+fx+pulserate+loudness+alfa+gamma)
Global pbbatg(1000, 9) As Double 'batke, (prop+miu+beta+alfa+indeks+fx+pulserate+loudness+alfa+gamma)
Global fx(1000, 1) As Double
Global solusirk(1000, 83, 3) As Double 'bat, data, kompartemen
Global solusirk1(1000, 83, 3) As Double 'bat, data, kompartemen
Global eror(1000, 84, 4) As Double
Global nilaiacak(1000, 2) As Double 'nilai acak untuk local search+indeks+indeks nilaiacak ls terpilih
Global nilaiacak1(1000, 2) As Double 'nilai acak untuk  perubahan loudness
Global nilaiacak2(1000) As Double 'nilai acak untuk frekuensi
Global nilaiacak3(1000, 1) As Double 'nilai acak untuk update pulserate dan loudenss
Global epsilon(1000, 3) As Double '(prop+miu+beta+alfa)
Global frek(1000, 1) As Double, frekuensi(1000, 2) As Double
Global fminn(1000, 2) As Double, fmaxx(1000, 2) As Double
Global miufix As Double, propfix As Double, betafix As Double, alfafix As Double, fxfix As Double, alfix As Double, rlfix As Double, batke As Integer
Global pbbatls(1000, 9) As Double 'batke, (prop+miu+beta+alfa+indeks+fx+pulserate+loudness+alfa+gamma)
Global bzt1(1000, 8) As Double
Global pbf(1000, 9) As Double '(prop+miu+beta+alfa+indeks+fx+pulserate+loudness+alfa+gamma)
Global maks(5) As Double, min(5) As Double
Global normalisasi(84, 6) As Double 'data, (S+I+R+M)
Global poladatain(4, 84, 7) As Double 'S/I/R, data, indeks+x1+x2+x3+target
Global poladatar(4, 84, 7) As Double
Global poladataru(4, 84, 7) As Double
Global bbin(5, 17) As Double '(S/R), pola ke, indeks, bobot dan bias(v11,v21,v31,v12,v22,v32,v13,v23,v33,b11,b12,b13,w1,w2,w3,b3)
Global bbinr(2, 10) As Double 'I,pola ke, indeks, bobot dan bias (v11,v12,v21,v22, b11,b12,w1,w2,b2)
Global FFin(2, 7) As Double '(S/R), zin1, zin2,zin3, z1,z2,z3,yin1, y1)
Global FFinr(1, 5) As Double 'R, zin1,zin2,z1,z2, yin1,y1
Global BPin(3, 22) As Double '(S/R), (d1, Dw11, Dw21, Dw31, Dw01, din1,din2,din3,d1,d2,d3, Dv11,Dv21,Dv31,Dv21,Dv22,Dv32,Dv31,Dv32,Dv33,Dv01,Dv02,Dv03)
Global BPinr(1, 13) As Double 'I, (d1, Dw11,Dw21, din1,din2, d1,d2, Dv11,D21,Dv12,Dv22,Dv01,Dv02)
Global outin(2, 84) As Double '(S/R), pola ke
Global outinr(1, 84) As Double 'I, pola ke
Global ermse(2, 84, 2) As Double '(S/I/R), pola ke, (indeks, mse, mse total)
Global Ddatain(84, 4) As Double 'tahun, (bulan, S, I, R)
Global errorval(84, 4) As Double
Global learning As Double, maksiter As Integer, poh As Integer, makseror As Double, err As Double
Global bbpr(5, 17) As Double 'bobot bias peramalan
Global FFpr(2, 7) As Double
Global BPpr(3, 22) As Double
Global outinpr(2, 84) As Double
Global ermsep(2, 84, 2) As Double
Global errvup(200, 84, 4) As Double
Global Ddataup(84, 4) As Double
Global solnum(84, 3) As Double





Function errormae(asli As Double, estimasi As Double) As Double
    errormae = Abs(asli - estimasi)
End Function

Function rmsfrek(fmin As Integer, fmax As Integer) As Integer()
rmsfrek = fmin + (fmax - fmin) * (Rnd)
End Function

Function Sgbiner(x As Double) As Double
Sgbiner = 1 / (1 + Exp(-x))
End Function

Function RandomAntara(Lowerbound As Double, Upperbound As Double) As Double
    RandomAntara = ((Upperbound - Lowerbound) * Rnd + Lowerbound)
End Function

Function FsAktivasi(x As Double) As Double
FsAktivasi = (1 - Exp(-x)) / (1 + Exp(-x))
End Function

Function Fsaksaktivasi(x As Double) As Double
Fsaksaktivasi = 0.5 * (1 + FsAktivasi(x)) * (1 - FsAktivasi(x))
End Function

Function MSE(x As Double, Y As Double) As Double
MSE = (x - Y) ^ 2
End Function

Function mmre(asli As Double, estimasi As Double) As Double
mmre = Abs(asli - estimasi) / asli

End Function


