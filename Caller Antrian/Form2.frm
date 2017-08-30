VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panggil Antrian"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   ".."
      Height          =   270
      Left            =   4200
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Text            =   "172.16.16.14"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Text            =   "2001"
      Top             =   840
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   7680
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Text            =   "C"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Text            =   "1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Panggil Ulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Panggil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "No Antri Panggil Ulang"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "ip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Jenis"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Loket"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If ws1.State <> sckClosed Then ws1.Close
    ws1.Connect Text4.Text, Text3.Text

    strSQL = "select norec, noantrian from antrianpasienregistrasi_t where  " & _
            "statuspanggil ='0' and " & _
            "jenis ='" & Text2.Text & "' and " & _
            "tanggalreservasi between '" & Format(Now(), "yyyy-mm-dd 00:00") & "' and '" & Format(Now(), "yyyy-mm-dd 23:59") & "' " & _
            "order by tanggalreservasi"
    Call ReadRs(strSQL)
    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        
        strSQL2 = "update antrianpasienregistrasi_t set statuspanggil ='2' where statuspanggil='1' and tempatlahir ='" & Text1.Text & "'"
        Call ReadRs2(strSQL2)
        strSQL3 = "update antrianpasienregistrasi_t set statuspanggil ='1',tempatlahir='" & Text1.Text & "' where norec='" & RS!norec & "'"
        Call ReadRs3(strSQL3)
    Else
        MsgBox "Antrian habis!", vbInformation, "..:."
    End If
End Sub

Private Sub Command2_Click()
    If ws1.State <> sckClosed Then ws1.Close
    ws1.Connect Text4.Text, Text3.Text

    strSQL = "select norec, noantrian from antrianpasienregistrasi_t where  " & _
            "noantrian ='" & Text5.Text & "' and " & _
            "jenis ='" & Text2.Text & "' and " & _
            "tanggalreservasi between '" & Format(Now(), "yyyy-mm-dd 00:00") & "' and '" & Format(Now(), "yyyy-mm-dd 23:59") & "' " & _
            "order by tanggalreservasi"
    Call ReadRs(strSQL)
    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        
        strSQL2 = "update antrianpasienregistrasi_t set statuspanggil ='2' where statuspanggil='1' and tempatlahir ='" & Text1.Text & "'"
        Call ReadRs2(strSQL2)
        strSQL3 = "update antrianpasienregistrasi_t set statuspanggil ='1',tempatlahir='" & Text1.Text & "' where norec='" & RS!norec & "'"
        Call ReadRs3(strSQL3)
    Else
        MsgBox "Tidak ada antrian!", vbInformation, "..:."
    End If
End Sub

Private Sub Command3_Click()
    frmSettingKoneksi.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
    Text4.Text = GetTxt("Setting.ini", "Caller", "IP_Display")
    Text3.Text = GetTxt("Setting.ini", "Caller", "Port")
    Text1.Text = GetTxt("Setting.ini", "Caller", "Loket")
    Text2.Text = GetTxt("Setting.ini", "Caller", "Jenis")
End Sub
