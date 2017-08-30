VERSION 5.00
Begin VB.Form frmPendaftaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendaftaran"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Pendaftaran(ByVal QueryText As String) As Byte()
    On Error Resume Next
    Dim Root As JNode
    Dim Param1() As String
    Dim Param2() As String
    Dim Param3() As String
    Dim Param4() As String
    Dim arrItem() As String
   
    If CN.State = adStateClosed Then Call openConnection
        
    
    If Len(QueryText) > 0 Then
        arrItem = Split(QueryText, "&")
        Param1 = Split(arrItem(0), "=")
        Param2 = Split(arrItem(1), "=")
        Param3 = Split(arrItem(2), "=")
        Param4 = Split(arrItem(3), "=")
        Select Case Param1(0)
            Case "cek-konek"
                lblStatus.Caption = "Cek"
                Set Root = New JNode
                Root("Status") = "Ok!!"
            
            Case "cetak-kartupasien"
                lblStatus.Caption = "Cetak Kartu Pasien"
'                Me.Show
                ReadRs "SELECT ps.namapasien,ps.nocm,jk.reportdisplay as jk , ps.tgllahir  " & _
                        " from pasien_m ps INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
                        " where ps.id='" & Param2(1) & "' "
                 If Not RS.EOF Then
                    
                   
                    '(strNocm, strNamaPasien, strTglLahir, strJk, view)
                    Call frmCetakPendaftaran.cetakKartuPasien(RS("nocm"), RS("namapasien"), Format(RS("tgllahir"), "dd/mm/yyyy"), RS("jk"), Param3(1))
                
                End If
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-kartupasien=1&id=1231=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                
            Case "cetak-buktipendaftaran"
                lblStatus.Caption = "Cetak Bukti Pendaftaran"
                Call frmCetakPendaftaran.cetakBuktiPendaftaran(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-buktipendaftaran=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
             Case "cetak-tracer"
                lblStatus.Caption = "Cetak Tracer"
                Call frmCetakPendaftaran.cetakTracer(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-tracer=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
            Case "cetak-sep"
                lblStatus.Caption = "Cetak SEP"
                Call frmCetakPendaftaran.cetakSep(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-sep=1&norec=40288c835ba4c322015ba816f5d0002c&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
            Case "cetak-buktilayanan"
                lblStatus.Caption = "Cetak Bukti Layanan"
                Call frmCetakPendaftaran.cetakBuktiLayanan(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-buktilayanan=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
            Case "cetak-labelpasien"
                lblStatus.Caption = "Cetak Label Pasien"
                Call frmCetakPendaftaran.cetakLabelPasien(Param2(1), Param3(1), Param4(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-labelpasien=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false&qty=2
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
            
            Case "cetak-summarylist"
                lblStatus.Caption = "Cetak Summary list Pasien Rawat Jalan"
                Call frmCetakPendaftaran.cetakSummaryList(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-summarylist=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
            
            Case "cetak-lembarmasukkeluar"
                lblStatus.Caption = "Cetak Lembar Masuk Keluar Pasien Rawat Inap"
                Call frmCetakPendaftaran.cetakLembarMasuk(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-summarylist=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
             
             Case "cetak-lembarpersetujuan"
                lblStatus.Caption = "Cetak Lembar Persetjuan Rawat Inap"
                Call frmCetakPendaftaran.cetakPersetujuan(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-summarylist=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
           
            Case Else
                Set Root = New JNode
                Root("Status") = "Error"
        End Select
        
    End If
    
    
    With GossRESTMain.STM
        .Open
        .Type = adTypeText
        .CharSet = "utf-8"
        .WriteText Root.JSON, adWriteChar
        .Position = 0
        .Type = adTypeBinary
        Pendaftaran = .Read(adReadAll)
        .Close
    End With
    If CN.State = adStateOpen Then CN.Close
    Unload Me
End Function

Private Sub CETAK_Billing(strNoregistrasi As String, jumlahCetak As Integer)
On Error Resume Next
    Dim prn As Printer
    Dim strPrinter As String
  
    ReadRs "select pp.norec,pp.tglpelayanan,pr.id as prid, pr.namaproduk, pp.jumlah,kl.id as klid, kl.namakelas, " & _
           "ru.id as ruid,ru.namaruangan,pp.produkfk,pp.hargajual,pg.id as pgid,pg.namalengkap,sp.nostruk, " & _
           "jpp.id as jppid,jpp.jenispetugaspe from " & _
           "pasiendaftar_t As pd " & _
           "inner  join antrianpasiendiperiksa_t as apd on apd.noregistrasifk= pd.norec " & _
           "inner join pelayananpasien_t as pp on pp.noregistrasifk= apd.norec " & _
           "inner join produk_m as pr ON pr.id= pp.produkfk " & _
           "inner JOIN  kelas_m as kl ON kl.id= apd.objectkelasfk " & _
           "inner join ruangan_m as ru ON ru.id= apd.objectruanganfk " & _
           "inner join pelayananpasienpetugas_t as ptu ON ptu.pelayananpasien= pp.norec " & _
           "inner join jenispetugaspelaksana_m as jpp ON jpp.id= ptu.objectjenispetugaspefk " & _
           "inner join pegawai_m as pg ON pg.id= ptu.objectpegawaifk " & _
           "left join strukpelayanan_t as sp ON sp.norec= pp.strukfk " & _
           "Where pd.tglpulang Is Not Null " & _
           "and pd.noregistrasi='" & strNoregistrasi & "'"
    
    
    strPrinter = GetTxt("Setting.ini", "Printer", "BuktiPendaftaran")
    If Printers.count > 0 Then
        For Each prn In Printers
            If prn.DeviceName = strPrinter Then
                Set Printer = prn
                Exit For
            End If
        Next prn
    End If
    
    For i = 0 To RS.RecordCount - 1
        'MsgBox "CETAK"
        Printer.FontSize = 10
        Printer.Print "     RUMAH SAKIT ANAK DAN BUNDA"
        Printer.FontSize = 18
        Printer.FontBold = True
        Printer.Print "      HARAPAN KITA"
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print "   Jl. S. Parman Kav.87, Slipi, Jakarta Barat"
        Printer.Print "      Telp. 021-5668286, 021-5668284"
        Printer.Print "      Fax.  021-5601816, 021-5673832"
        Printer.Print "___________________________________"
        Printer.Print ""
        Printer.Print "Tanggal :" & Format(Now(), "yyyy MM dd hh:mm")
        Printer.Print ""
        Printer.FontSize = 12
          '1,3,,4,6,8,10,12,13,15
        Printer.Print RS(1) & " " & RS(1) & " " & RS(3) & " " & RS(4) & " " & RS(6) & " " & RS(8) & " " & RS(10) & " " & RS(12) & " "
        
        Printer.EndDoc
    Next
End Sub

