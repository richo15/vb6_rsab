VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakFarmasiApotik 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmCetakFarmasiApotik.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9075
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
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
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
Attribute VB_Name = "frmCetakFarmasiApotik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportResep As New cr_RincianBiayaResep_2

Dim ii As Integer
Dim tempPrint1 As String
Dim p As Printer
Dim p2 As Printer
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String

Dim bolStrukResep As Boolean


Dim strPrinter As String
Dim strPrinter1 As String
Dim PrinterNama As String

Dim adoReport As New ADODB.Command

Private Sub cmdCetak_Click()
  If cboPrinter.Text = "" Then MsgBox "Printer belum dipilih", vbInformation, ".: Information": Exit Sub
    If bolStrukResep = True Then
        ReportResep.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        ReportResep.PrintOut False
    
    End If
End Sub

Private Sub CmdOption_Click()
    
    If bolStrukResep = True Then
        ReportResep.PrinterSetup Me.hwnd
    End If
    
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    
    Dim p As Printer
    cboPrinter.Clear
    For Each p In Printers
        cboPrinter.AddItem p.DeviceName
    Next
    strPrinter = strPrinter1
    
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCetakFarmasiApotik = Nothing

End Sub

Public Sub cetakStrukResep(strNores As String, view As String, strUser As String)
On Error GoTo errLoad
Set frmCetakFarmasiApotik = Nothing
Dim strSQL As String

bolStrukResep = True


    With ReportResep
                
                Set adoReport = New ADODB.Command
                 adoReport.ActiveConnection = CN_String
                
                strSQL = "SELECT pd.noregistrasi, ps.nocm, " & _
                          " ps.namapasien || ' ( ' || jk.reportdisplay || ' )' as namapasienjk , kpp.kelompokpasien, " & _
                          " ps.tgllahir, pd.tglregistrasi, ru.namaruangan AS ruanganpasien, " & _
                          " sr.noresep, pp.rke, pr.namaproduk || ' / ' || sstd.satuanstandar as namaprodukstandar, " & _
                          " pp.jumlah , pp.hargasatuan,(pp.jumlah || ' x ' || pp.hargasatuan) as qtyhrg,(pp.jumlah * pp.hargasatuan ) as totalharga ,jnskem.jeniskemasan, pgw.namalengkap, " & _
                          " CASE when pp.hargadiscount isnull then 0 ELSE  pp.hargadiscount * pp.jumlah end as totaldiscound, " & _
                          " ((pp.jumlah * pp.hargasatuan ) - (CASE when pp.hargadiscount isnull then 0 ELSE  pp.hargadiscount * pp.jumlah end)) as totalbiaya FROM pelayananpasien_t AS pp " & _
                          " INNER JOIN antrianpasiendiperiksa_t AS apdp ON pp.noregistrasifk = apdp.norec " & _
                          " INNER JOIN pasiendaftar_t AS pd ON apdp.noregistrasifk = pd.norec " & _
                          " INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                          " INNER JOIN produk_m AS pr ON pp.produkfk = pr.id " & _
                          " INNER JOIN ruangan_m AS ru ON apdp.objectruanganfk = ru.id " & _
                          " INNER JOIN strukresep_t AS sr ON pp.strukresepfk = sr.norec " & _
                          " INNER JOIN jeniskemasan_m AS jnskem ON pp.jeniskemasanfk = jnskem.id " & _
                          " INNER JOIN pegawai_m AS pgw ON sr.penulisresepfk = pgw.id " & _
                          " INNER JOIN satuanstandar_m AS sstd ON pp.satuanviewfk = sstd.id " & _
                          " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                          " INNER JOIN kelompokpasien_m AS kpp ON pd.objectkelompokpasienlastfk = kpp.id" & _
                          " WHERE sr.norec='" & strNores & "'"
            
                ReadRs strSQL & " limit 1 "
                
                adoReport.CommandText = strSQL
                adoReport.CommandType = adCmdUnknown
               .database.AddADOCommand CN_String, adoReport
               
              
                .txtNoPendaftaran.SetText RS("noregistrasi")
                .txtnocm.SetText RS("nocm")
                .txtnmpasien.SetText RS("namapasienjk")
                .txtklpkpasien.SetText RS("kelompokpasien")
                '.txtPenjamin.SetText IIf(IsNull(RS("NamaPenjamin")), "Sendiri", RS("NamaPenjamin"))
                .txtNamaRuangan.SetText RS("ruanganpasien")
                .txtumur.SetText hitungUmur(Format(RS("tgllahir"), "dd/mm/yyyy"), Format(RS("tglregistrasi"), "dd/mm/yyyy"))
                .txtNamaDokter.SetText RS("namalengkap")
                .txtuser.SetText strUser
                
                
              '  .usSatuan.SetUnboundFieldSource ("{ado.SatuanJmlK}")
             '   .udtanggal.SetUnboundFieldSource ("{Ado.tglpelayanan}")
                .usNoResep.SetUnboundFieldSource ("{Ado.noresep}")
                .ucbiayasatuan.SetUnboundFieldSource ("{Ado.totalharga}")
                .ucHrgSatuan.SetUnboundFieldSource ("{Ado.hargasatuan}")
                .ustindakan.SetUnboundFieldSource ("{Ado.namaprodukstandar}")
                .usQtyHrg.SetUnboundFieldSource ("{Ado.qtyhrg}")
                .unQtyTotal.SetUnboundFieldSource ("{Ado.jumlah}")
                .ucGrandTotal.SetUnboundFieldSource ("{Ado.totalharga}")
                .undis.SetUnboundFieldSource ("{Ado.totaldiscound}")
                .unTotal.SetUnboundFieldSource ("{Ado.totalbiaya}")
                
                .unRacikanKe.SetUnboundFieldSource ("{ado.rke}")
                .usJenisObat.SetUnboundFieldSource ("{ado.jeniskemasan}")
                
                
                
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "CetakResep")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = ReportResep
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

End Sub

