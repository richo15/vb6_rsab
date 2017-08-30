VERSION 5.00
Begin VB.Form frmSettingKoneksi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Koneksi"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboVol 
      Height          =   360
      Left            =   1560
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox trg10 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox trg11 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox trg12 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox trg13 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox trg20 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox trg21 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox trg22 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox trg23 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Volume Video :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Display Trigger :"
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Database :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "User Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Port :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Host :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettingKoneksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'    SaveTxt "Setting.ini", "Koneksi", "host", Text1.Text
'    SaveTxt "Setting.ini", "Koneksi", "port", Text2.Text
'    SaveTxt "Setting.ini", "Koneksi", "username", Text3.Text
'    SaveTxt "Setting.ini", "Koneksi", "password", Text4.Text
'    SaveTxt "Setting.ini", "Koneksi", "database", Text5.Text
    
    SaveTxt "Setting.ini", "Koneksi", "a", Text1.Text
    SaveTxt "Setting.ini", "Koneksi", "b", Text2.Text
    SaveTxt "Setting.ini", "Koneksi", "c", Text3.Text
    SaveTxt "Setting.ini", "Koneksi", "d", Text4.Text
    SaveTxt "Setting.ini", "Koneksi", "e", Text5.Text
    
    SaveTxt "Setting.ini", "Volume", "Video", cboVol.Text
    
    SaveTxt "Setting.ini", "DisplayTrigger_1", "IP", trg10.Text
    SaveTxt "Setting.ini", "DisplayTrigger_2", "IP", trg11.Text
    SaveTxt "Setting.ini", "DisplayTrigger_3", "IP", trg12.Text
    SaveTxt "Setting.ini", "DisplayTrigger_4", "IP", trg13.Text
    
    SaveTxt "Setting.ini", "DisplayTrigger_1", "PORT", trg20.Text
    SaveTxt "Setting.ini", "DisplayTrigger_2", "PORT", trg21.Text
    SaveTxt "Setting.ini", "DisplayTrigger_3", "PORT", trg22.Text
    SaveTxt "Setting.ini", "DisplayTrigger_4", "PORT", trg23.Text
    
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    Text1.Text = GetTxt("Setting.ini", "Koneksi", "a")
    Text2.Text = GetTxt("Setting.ini", "Koneksi", "b")
    Text3.Text = GetTxt("Setting.ini", "Koneksi", "c")
    Text4.Text = GetTxt("Setting.ini", "Koneksi", "d")
    Text5.Text = GetTxt("Setting.ini", "Koneksi", "e")
    
    cboVol.Text = GetTxt("Setting.ini", "Volume", "Video")
   
    trg10.Text = GetTxt("Setting.ini", "DisplayTrigger_1", "IP")
    trg11.Text = GetTxt("Setting.ini", "DisplayTrigger_2", "IP")
    trg12.Text = GetTxt("Setting.ini", "DisplayTrigger_3", "IP")
    trg13.Text = GetTxt("Setting.ini", "DisplayTrigger_4", "IP")
    
    trg20.Text = GetTxt("Setting.ini", "DisplayTrigger_1", "PORT")
    trg21.Text = GetTxt("Setting.ini", "DisplayTrigger_2", "PORT")
    trg22.Text = GetTxt("Setting.ini", "DisplayTrigger_3", "PORT")
    trg23.Text = GetTxt("Setting.ini", "DisplayTrigger_4", "PORT")
    
    cboVol.AddItem 0
    cboVol.AddItem 30
    cboVol.AddItem 60
    cboVol.AddItem 100
End Sub
