VERSION 5.00
Begin VB.Form frmSettingIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   Icon            =   "frmSettingIP.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtJenis 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLoket 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtIp 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Jenis"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Loket"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Port"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Ip Display"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmSettingIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    SaveTxt "Setting.ini", "Caller", "IP_Display", txtIp.Text
    SaveTxt "Setting.ini", "Caller", "Port", txtPort.Text
    SaveTxt "Setting.ini", "Caller", "Loket", txtLoket.Text
    SaveTxt "Setting.ini", "Caller", "Jenis", txtJenis.Text
End Sub

Private Sub Form_Load()
    txtIp.Text = GetTxt("Setting.ini", "Caller", "IP_Display")
    txtPort.Text = GetTxt("Setting.ini", "Caller", "Port")
    txtLoket.Text = GetTxt("Setting.ini", "Caller", "Loket")
    txtJenis.Text = GetTxt("Setting.ini", "Caller", "Jenis")
End Sub
