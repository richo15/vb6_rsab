VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRCetakBilling 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCetakBilling.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   5820
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOption 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmCRCetakBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crBilling
'Dim bolSuppresDetailSection10 As Boolean
'Dim ii As Integer
'Dim tempPrint1 As String
'Dim p As Printer
'Dim p2 As Printer
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String

Private Sub cmdCetak_Click()
    Report.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
    'PrinterNama = cboPrinter.Text
    Report.PrintOut False
End Sub

Private Sub CmdOption_Click()
    Report.PrinterSetup Me.hwnd
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    Dim p As Printer
    cboPrinter.Clear
    For Each p In Printers
        cboPrinter.AddItem p.DeviceName
    Next
    cboPrinter.Text = GetTxt("Setting.ini", "Printer", "Billing")
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCRCetakBilling = Nothing
End Sub

Public Sub CetakBilling(strNoregistrasi As String, jumlahCetak As Integer, view As String)
'On Error GoTo errLoad
On Error Resume Next

Set frmCRCetakBilling = Nothing
Dim adocmd As New ADODB.Command

Set Report = New crBilling
    strSQL = "SELECT pd.noregistrasi,ps.nocm,(ps.namapasien || ' ( ' || jk.reportdisplay || ' )' ) as namapasienjk ,ru.namaruangan,kl.namakelas, " & _
                " pg.namalengkap,pd.tglregistrasi,pd.tglpulang,rk.namarekanan,pp.tglpelayanan, ru2.namaruangan as ruanganTindakan, " & _
                " pr.namaproduk,jp.jenisproduk, pg2.namalengkap as dokter,pp.jumlah,pp.hargajual, " & _
                " case when pp.hargadiscount is null then 0 else pp.hargadiscount end as diskon, " & _
                " (pp.jumlah*(pp.hargajual-case when pp.hargadiscount is null then 0 else pp.hargadiscount end)) as total, kmr.namakamar " & _
                " from pasiendaftar_t as pd INNER join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
                " INNER join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
                " INNER join produk_m as pr on pr.id=pp.produkfk " & _
                " INNER join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
                " INNER join jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
                " INNER join pasien_m as ps on ps.id=pd.nocmfk " & _
                " INNER join jeniskelamin_m as jk on jk.id=ps.objectjeniskelaminfk " & _
                " INNER join ruangan_m  as ru on ru.id=pd.objectruanganlastfk " & _
                " INNER join ruangan_m  as ru2 on ru2.id=apd.objectruanganfk " & _
                " left join kelas_m  as kl on kl.id=pd.objectkelasfk " & _
                " INNER join pegawai_m  as pg on pg.id=pd.objectpegawaifk " & _
                " INNER join pegawai_m  as pg2 on pg2.id=apd.objectpegawaifk " & _
                " left join rekanan_m  as rk on rk.id=pd.objectrekananfk " & _
                " left join kamar_m  as kmr on kmr.id=apd.objectkamarfk " & _
                " where pd.noregistrasi='" & strNoregistrasi & "' "
    
    ReadRs2 "SELECT " & _
            "sum((pp.jumlah*(pp.hargajual-case when pp.hargadiscount is null then 0 else pp.hargadiscount end))) as total " & _
            "from pasiendaftar_t as pd " & _
            "inner join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
            "inner join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
            "inner join produk_m as pr on pr.id=pp.produkfk " & _
            "inner join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
            "inner join jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
            "inner join pasien_m as ps on ps.id=pd.nocmfk " & _
            "inner join jeniskelamin_m as jk on jk.id=ps.objectjeniskelaminfk " & _
            "inner join ruangan_m  as ru on ru.id=pd.objectruanganlastfk " & _
            "inner join ruangan_m  as ru2 on ru2.id=apd.objectruanganfk " & _
            "LEFT join kelas_m  as kl on kl.id=pd.objectkelasfk " & _
            "inner join pegawai_m  as pg on pg.id=pd.objectpegawaifk " & _
            "inner join pegawai_m  as pg2 on pg2.id=apd.objectpegawaifk " & _
            "left join rekanan_m  as rk on rk.id=pd.objectrekananfk " & _
            "where pd.noregistrasi='" & strNoregistrasi & "' "

   
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
        'If Not RS.EOF Then
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasienjk}")
            .usRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usKamar.SetUnboundFieldSource IIf(IsNull("{ado.namakamar}") = True, "-", ("{ado.namakamar}"))
            .usKelasH.SetUnboundFieldSource ("{ado.namakelas}")
            .usDokterPJawab.SetUnboundFieldSource ("{ado.namalengkap}")
            .udTglMasuk.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .udTglPlng.SetUnboundFieldSource IIf(IsNull("{ado.tglpulang}") = True, "-", ("{ado.tglpulang}"))
            .usPenjamin.SetUnboundFieldSource IIf(IsNull("{ado.namarekanan}") = True, ("-"), ("{ado.namarekanan}"))
        
            
            .txtTerbilang.SetText TERBILANG(RS2!total)
            
            .usJenisProduk.SetUnboundFieldSource ("{ado.jenisproduk}")
            .udTanggal.SetUnboundFieldSource ("{ado.tglpelayanan}")
            .usLayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usDokter.SetUnboundFieldSource ("{ado.dokter}")
            .unQty.SetUnboundFieldSource ("{ado.jumlah}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .ucTotal.SetUnboundFieldSource ("{ado.total}")
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "Billing")
                .SelectPrinter "winspool", strPrinter, "Ne00:"
                .PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ReportSource = Report
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
            End If
        'End If
    End With
Exit Sub
'errLoad:
End Sub
