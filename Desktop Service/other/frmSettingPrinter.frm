VERSION 5.00
Begin VB.Form frmSettingPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Printer"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
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
   ScaleHeight     =   3450
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Printer"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Jenis Cetakan"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettingPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
    Combo2.Text = GetSetting("Jasamedika Service", Combo1.Text, "Printer")
End Sub

Private Sub Combo1_Click()
    Call Combo1_Change
End Sub

Private Sub Command1_Click()
    'SaveSetting "Jasamedika Service", "CetakAntrian", "Jenis", Combo1.Text
    SaveSetting "Jasamedika Service", Combo1.Text, "Printer", Combo2.Text
End Sub

Private Sub Form_Load()
    For Each prn In Printers
'        If prn.DeviceName = "Microsoft XPS Document Writer" Then
'            Set Printer = prn
'            Exit For
'        End If
        Combo2.AddItem prn.DeviceName
    Next prn
    
    Combo1.AddItem "CetakAntrian"
End Sub
