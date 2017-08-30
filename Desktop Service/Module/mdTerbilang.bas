Attribute VB_Name = "mdTerbilang"
Public Function TERBILANG(x As Double) As String
Dim tampung As Double
Dim teks As String
Dim bagian As String
Dim i As Integer
Dim tanda As Boolean
 
Dim letak(5)
letak(1) = "RIBU "
letak(2) = "JUTA "
letak(3) = "MILYAR "
letak(4) = "TRILYUN "
 
If (x < 0) Then
    TERBILANG = ""
Exit Function
End If
 
If (x = 0) Then
    TERBILANG = "NOL"
Exit Function
End If
 
If (x < 2000) Then
    tanda = True
End If
teks = ""
 
If (x >= 1E+15) Then
    TERBILANG = "NILAI TERLALU BESAR"
Exit Function
End If
 
For i = 4 To 1 Step -1
    tampung = Int(x / (10 ^ (3 * i)))
    If (tampung > 0) Then
        bagian = ratusan(tampung, tanda)
        teks = teks & bagian & letak(i)
    End If
    x = x - tampung * (10 ^ (3 * i))
Next
 
teks = teks & ratusan(x, False)
TERBILANG = teks & " RUPIAH"
End Function
 
Function ratusan(ByVal y As Double, ByVal flag As Boolean) As String
Dim tmp As Double
Dim bilang As String
Dim bag As String
Dim j As Integer
 
Dim angka(9)
angka(1) = "SE"
angka(2) = "DUA "
angka(3) = "TIGA "
angka(4) = "EMPAT "
angka(5) = "LIMA "
angka(6) = "ENAM "
angka(7) = "TUJUH "
angka(8) = "DELAPAN "
angka(9) = "SEMBILAN "
 
Dim posisi(2)
posisi(1) = "PULUH "
posisi(2) = "RATUS "
 
bilang = ""
For j = 2 To 1 Step -1
    tmp = Int(y / (10 ^ j))
    If (tmp > 0) Then
        bag = angka(tmp)
        If (j = 1 And tmp = 1) Then
            y = y - tmp * 10 ^ j
            If (y >= 1) Then
                posisi(j) = "BELAS "
            Else
                angka(y) = "SE"
            End If
            bilang = bilang & angka(y) & posisi(j)
            ratusan = bilang
            Exit Function
        Else
            bilang = bilang & bag & posisi(j)
    End If
End If
y = y - tmp * 10 ^ j
Next
 
If (flag = False) Then
    angka(1) = "SATU "
End If
bilang = bilang & angka(y)
ratusan = bilang
End Function


Public Function hitungUmur(dateOfBird As Date, fromData As Date) As String
    Dim dateNow As Date
    Dim tgl As Date
    Dim tgl1 As Date
 
    Dim years As Long
    Dim months As Long
    Dim days As Long
 
    Dim yearWord As String
    Dim monthWord As String
    Dim dayWord As String
 
    dateNow = fromData
    tgl = dateOfBird
 
    ' menghitung tahun
    years = DateDiff("yyyy", tgl, dateNow)
    If Month(tgl) > Month(dateNow) Then
        years = years - 1
    ElseIf Month(tgl) = Month(dateNow) Then
        years = 0
    ElseIf Month(tgl) = Month(dateNow) And Day(tgl) > Day(dateNow) Then
        years = years - 1
    ElseIf Month(tgl) = Month(dateNow) And Day(tgl) = Day(dateNow) Then
        GoTo finally ' jika bulan dan tanggal sama maka perhitungan selesai
    End If
 
    ' menghitung bulan
    tgl = DateAdd("yyyy", years, tgl)
    months = DateDiff("m", tgl, dateNow)
    If Day(tgl) > Day(dateNow) Then
        months = months - 1
    
    ElseIf Month(tgl) = Month(dateNow) And Day(tgl) > Day(dateNow) Then
        months = months - 1
    ElseIf Month(tgl) = Month(dateNow) And Day(tgl) = Day(dateNow) Then
        months = 0
    End If
 
    tgl = DateAdd("m", months, tgl)
 
    ' menghitung hari
    days = DateDiff("d", tgl, dateNow)
 
finally:
    yearWord = IIf(years = 0, "", years & " Th ")
    monthWord = IIf(months = 0, "", months & " Bl ")
    dayWord = days & " Hr "
 
    hitungUmur = yearWord & monthWord & dayWord
    hitungUmur = Trim(hitungUmur)
End Function

