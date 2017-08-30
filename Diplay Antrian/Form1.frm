VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   13500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   900
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1437
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wst4 
      Left            =   7080
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wst3 
      Left            =   6600
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wst2 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wst1 
      Left            =   5640
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WS1 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1440
      Top             =   3600
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   960
      Top             =   3600
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   3600
   End
   Begin MSWinsockLib.Winsock WS2 
      Left            =   2040
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS3 
      Left            =   2520
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS4 
      Left            =   1560
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS5 
      Left            =   2040
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock WS6 
      Left            =   2520
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock ws7 
      Left            =   1560
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock ws8 
      Left            =   2040
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock ws9 
      Left            =   2520
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock ws10 
      Left            =   1560
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   10275
      Left            =   4560
      TabIndex        =   29
      Top             =   1920
      Width           =   11520
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "mini"
      stretchToFit    =   -1  'True
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   20320
      _cy             =   18124
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   0
      TabIndex        =   28
      Top             =   4080
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   0
      TabIndex        =   27
      Top             =   5760
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   8
      Left            =   0
      TabIndex        =   26
      Top             =   7425
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   9
      Left            =   0
      TabIndex        =   25
      Top             =   9015
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 7 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   105
      TabIndex        =   24
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 8 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   105
      TabIndex        =   23
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 9 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   105
      TabIndex        =   22
      Top             =   7005
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 10 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   105
      TabIndex        =   21
      Top             =   8640
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   12600
      Width           =   3975
   End
   Begin VB.Label lblWs 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Data ... ... ..."
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Image pic 
      Height          =   9495
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   11415
   End
   Begin VB.Label runText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SELAMAT DATANG DI RSAB HARAPAN KITA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   20520
      TabIndex        =   18
      Top             =   11160
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 6 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   16440
      TabIndex        =   17
      Top             =   10080
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   16335
      TabIndex        =   16
      Top             =   10440
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 5 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   16440
      TabIndex        =   15
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 4 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   16440
      TabIndex        =   14
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 3 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   16440
      TabIndex        =   13
      Top             =   5085
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 2 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   16440
      TabIndex        =   12
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loket 1 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   16440
      TabIndex        =   11
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblJam 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "11:11:11"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   54.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   35.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   12000
      Width           =   8295
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   35.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   13245
      TabIndex        =   8
      Top             =   12000
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   11280
      Width           =   5295
   End
   Begin VB.Label lblconn 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   16335
      TabIndex        =   4
      Top             =   8700
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   16335
      TabIndex        =   3
      Top             =   7095
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   16335
      TabIndex        =   2
      Top             =   5505
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   16335
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.999"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   16335
      TabIndex        =   0
      Top             =   2160
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SND_APPLICATION = &H80
'The sound is played using an application-specific association
Const SND_ALIAS = &H10000
'The pszSound parameter is a system-event alias in the registry or the WIN.INI file.
'Do not use with either SND_FILENAME or SND_RESOURCE.
Const SND_ALIAS_ID = &H110000
'The pszSound parameter is a predefined sound identifier.
Const SND_ASYNC = &H1
'The sound is played asynchronously and PlaySound returns immediately after beginning the sound. To terminate an asynchronously played waveform sound, call PlaySound with pszSound set to NULL.
Const SND_FILENAME = &H20000
'The pszSound parameter is a filename
Const SND_LOOP = &H8
'The sound plays repeatedly until PlaySound is called again with the pszSound parameter set to NULL. You must also specify the SND_ASYNC flag to indicate an asynchronous sound event
Const SND_MEMORY = &H4
'A sound event’s file is loaded in RAM. The parameter specified by pszSound must point to an image of a sound in memory.
Const SND_NODEFAULT = &H2
'No default sound event is used. If the sound cannot be found, PlaySound returns silently without playing the default sound.
Const SND_NOSTOP = &H10
'The specified sound event will yield to another sound event that is already playing. If a sound cannot be played because the resource needed to generate that sound is busy playing another sound, the function immediately returns FALSE without playing the requested sound. If this flag is not specified, PlaySound attempts to stop the currently playing sound so that the device can be used to play the new sound.
Const SND_NOWAIT = &H2000
'If the driver is busy, return immediately without playing the sound.
Const SND_PURGE = &H40
'Sounds are to be stopped for the calling task. If pszSound is not NULL, all instances of the specified sound are stopped. If pszSound is NULL, all sounds that are playing on behalf of the calling task are stopped. You must also specify the instance handle to stop SND_RESOURCE events.
Const SND_RESOURCE = &H40004 'The pszSound parameter is a resource identifier; hmod must identify the instance that contains the resource.
Const SND_SYNC = &H0
'Synchronous playback of a sound event. PlaySound returns after the sound event completes.

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Dim tmt As Integer
Dim tmt2 As Integer
Dim tmt3 As Integer
Dim tmt4 As Integer
Dim onload As Boolean
Dim loket As Integer
Dim KedipLoket As Integer
Dim jenisAntrian As Integer
Dim TimeToRefresh As Integer
Dim vdeo As Integer
Dim reload As Boolean
Dim sora As Integer



Private Sub File1_DblClick()
'    For i = 0 To File1.ListCount - 1
'        Debug.Print File1.List(i)
'    Next
End Sub

Private Sub Form_DblClick()
    'frmSetServer.Show
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
If KeyCode = 112 Then frmSetServer.Show: Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo hell
    Me.Move 0, 0, Screen.Width, Screen.Height
    lblconn.Caption = dbConn
    Timer1.Enabled = True
    For i = 0 To 9
        lbl(i).Caption = ""
    Next
    File1.Path = App.Path & "\video"
    tmt3 = 10
    tmt = 100
    onload = True
    Label1.Caption = App.Path
    tmt3 = 60
    
'    If GetSetting("Antrian", "Video", "suara") = "ON" Then
'        sora = 70
'    Else
'        sora = 0
'    End If
    
    sora = GetTxt("Setting.ini", "Volume", "Video")
    

'    ##DIRECT SHOW
'    DS_top = GetSetting("Antrian", "Video", "top") '0  '136
'    DS_left = GetSetting("Antrian", "Video", "left") '312
'    DS_width = GetSetting("Antrian", "Video", "width") '761
'    DS_height = GetSetting("Antrian", "Video", "height") '633
'    Fullscreen_Enabled = False
'    vdeo = 0
'    DirectShow_Load_Media App.Path & "\video\" & File1.List(0)
'    DirectShow_Play
'    DirectShow_Volume sora
'    pic.Visible = False
'   ##END DIRECT SHOW

    WindowsMediaPlayer1.URL = App.Path & "\video\" & File1.List(0)
    WindowsMediaPlayer1.settings.Volume = sora
    
    
    Call OpenPortWinsock
    Call loadAntrian
    
    '@IPComputer,@Port,@Loket,@StatusEnabled
'    strSQL = "delete from AntrianIpPort where Loket='1'"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "delete from AntrianIpPort where Loket='2'"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "delete from AntrianIpPort where Loket='3'"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "delete from AntrianIpPort where Loket='4'"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "delete from AntrianIpPort where Loket='5'"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "delete from AntrianIpPort where Loket='6'"
'    Call msubRecFO(rs, strSQL)
'
'    strSQL = "insert into AntrianIpPort values ('" & WS1.LocalIP & "','2001','1','1')"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "insert into AntrianIpPort values ('" & WS1.LocalIP & "','2002','2','1')"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "insert into AntrianIpPort values ('" & WS1.LocalIP & "','2003','3','1')"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "insert into AntrianIpPort values ('" & WS1.LocalIP & "','2004','4','1')"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "insert into AntrianIpPort values ('" & WS1.LocalIP & "','2005','5','1')"
'    Call msubRecFO(rs, strSQL)
'    strSQL = "insert into AntrianIpPort values ('" & WS1.LocalIP & "','2006','6','1')"
'    Call msubRecFO(rs, strSQL)
'    MsgBox "3"
    Exit Sub
    
hell:
    frmSetServer.Show
End Sub

Private Sub LoadTrigger()
Dim a1, a2, b1, b2, c1, c2, d1, d2 As String

    a1 = GetTxt("Setting.ini", "DisplayTrigger_1", "IP")
    b1 = GetTxt("Setting.ini", "DisplayTrigger_2", "IP")
    c1 = GetTxt("Setting.ini", "DisplayTrigger_3", "IP")
    d1 = GetTxt("Setting.ini", "DisplayTrigger_4", "IP")
    
    a2 = GetTxt("Setting.ini", "DisplayTrigger_1", "PORT")
    b2 = GetTxt("Setting.ini", "DisplayTrigger_2", "PORT")
    c2 = GetTxt("Setting.ini", "DisplayTrigger_3", "PORT")
    d2 = GetTxt("Setting.ini", "DisplayTrigger_4", "PORT")
    
    If a1 <> "" And a2 <> "" Then
        If wst1.State <> sckClosed Then wst1.Close
        wst1.Connect a1, a2
    End If
    
    If b1 <> "" And b2 <> "" Then
        If wst2.State <> sckClosed Then wst2.Close
        wst2.Connect b1, b2
    End If
    
    
    If c1 <> "" And c2 <> "" Then
        If wst3.State <> sckClosed Then wst3.Close
        wst3.Connect c1, c2
    End If
    
    
    If d1 <> "" And d2 <> "" Then
        If wst4.State <> sckClosed Then wst4.Close
        wst4.Connect d1, d2
    End If
End Sub

Private Sub Label1_DblClick()
    WindowsMediaPlayer1.Controls.Next
End Sub

Private Sub Label2_Click(Index As Integer)
    frmSettingKoneksi.Show vbModal
End Sub

Private Sub lblJam_DblClick()
    frmSettingKoneksi.Show vbModal
End Sub

Private Sub Timer1_Timer()
    tmt = tmt + 1
    tmt4 = tmt4 + 1
    If tmt4 > 5 Then
    'If Val(Format(Now(), "ss")) Mod 5 = 0 Then
        Call loadAntrian
        'tmt = 0
        tmt4 = 0
    End If
    If tmt > 100 Then
        Timer1.Enabled = False
        tmt = 0
        reload = True
        lblWs.Visible = False
        Call OpenPortWinsock
    End If
'    lblJam.Caption = Format(Now(), "hh:nn:ss")
'    If Val(Format(Now(), "ss")) Mod 10 = 0 Then
'        'pic.Picture = File1.Path & "\File1.Tag"
'        pic.Picture = LoadPicture(File1.Path & "\" & File1.List(Val(File1.Tag)))
'        File1.Tag = Val(File1.Tag) + 1
'        If Val(File1.Tag) > File1.ListCount - 1 Then File1.Tag = 0
'    End If
'    On Error GoTo Error_Handler
'    Label3.Caption = DirectShow_Position.CurrentPosition & "/" & DirectShow_Position.StopTime
'    If DirectShow_Position.CurrentPosition >= DirectShow_Position.StopTime Then
'            'DirectShow_Position.CurrentPosition = 0
'        vdeo = vdeo + 1
'        If vdeo > File1.ListCount - 1 Then vdeo = 0
'        DirectShow_Load_Media App.Path & "\video\" & File1.List(vdeo)
''    DirectShow_Loop
'        DirectShow_Play
'        DirectShow_Volume 0
'    End If
'Error_Handler:
End Sub

Private Sub OpenPortWinsock()
'    Timer1.Enabled = False
    
    If WS1.State <> 0 Then WS1.Close
    WS1.LocalPort = 2001
    WS1.Listen
    
    If WS2.State <> 0 Then WS2.Close
    WS2.LocalPort = 2002
    WS2.Listen
    
    If WS3.State <> 0 Then WS3.Close
    WS3.LocalPort = 2003
    WS3.Listen
    
    If WS4.State <> 0 Then WS4.Close
    WS4.LocalPort = 2004
    WS4.Listen

    If WS5.State <> 0 Then WS5.Close
    WS5.LocalPort = 2005
    WS5.Listen

    If WS6.State <> 0 Then WS6.Close
    WS6.LocalPort = 2006
    WS6.Listen
    
     If ws7.State <> 0 Then ws7.Close
    ws7.LocalPort = 2007
    ws7.Listen
    
     If ws8.State <> 0 Then ws8.Close
    ws8.LocalPort = 2008
    ws8.Listen
    
     If ws9.State <> 0 Then ws9.Close
    ws9.LocalPort = 2009
    ws9.Listen
    
     If ws10.State <> 0 Then ws10.Close
    ws10.LocalPort = 2010
    ws10.Listen
    
    
End Sub


Private Sub ClosePortWinsock()
'    Timer1.Enabled = False
    
    If WS1.State <> 0 Then WS1.Close
    'WS1.LocalPort = 2001
    'WS1.Listen
    
    If WS2.State <> 0 Then WS2.Close
    'WS2.LocalPort = 2002
    'WS2.Listen
    
    If WS3.State <> 0 Then WS3.Close
    'WS3.LocalPort = 2003
    'WS3.Listen
    
    If WS4.State <> 0 Then WS4.Close
    'WS4.LocalPort = 2004
    'WS4.Listen

    If WS5.State <> 0 Then WS5.Close
    'WS5.LocalPort = 2005
    'WS5.Listen

    If WS6.State <> 0 Then WS6.Close
    
    If ws7.State <> 0 Then ws7.Close
    If ws8.State <> 0 Then ws8.Close
    If ws9.State <> 0 Then ws9.Close
    If ws10.State <> 0 Then ws10.Close
    'WS6.LocalPort = 2006
    'WS6.Listen
    
    
End Sub


Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
    If WS1.State <> sckClosed Then
        WS1.Close
    End If
'    lblWs.Visible = True
    WS1.Accept requestID
    WS1.SendData "OK"
    
    Call ClosePortWinsock
    'WS1.Close
'    WS1.LocalPort = 2001
'    WS1.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub ws2_ConnectionRequest(ByVal requestID As Long)
    If WS2.State <> sckClosed Then
        WS2.Close
    End If
'    lblWs.Visible = True
    WS2.Accept requestID
    WS2.SendData "OK"
    
    Call ClosePortWinsock
    'WS2.Close
'    WS2.LocalPort = 2002
'    WS2.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub ws3_ConnectionRequest(ByVal requestID As Long)
    If WS3.State <> sckClosed Then
        WS3.Close
    End If
'    lblWs.Visible = True
    WS3.Accept requestID
    WS3.SendData "OK"
    
    Call ClosePortWinsock
'    WS3.Close
'    WS3.LocalPort = 2003
'    WS3.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub ws4_ConnectionRequest(ByVal requestID As Long)
    If WS4.State <> sckClosed Then
        WS4.Close
    End If
'    lblWs.Visible = True
    WS4.Accept requestID
    WS4.SendData "OK"
    
    Call ClosePortWinsock
'    WS4.Close
'    WS4.LocalPort = 2004
'    WS4.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub ws5_ConnectionRequest(ByVal requestID As Long)
    If WS5.State <> sckClosed Then
        WS5.Close
    End If
'    lblWs.Visible = True
    WS5.Accept requestID
    WS5.SendData "OK"
    
    Call ClosePortWinsock
'    WS5.Close
'    WS5.LocalPort = 2005
'    WS5.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub ws6_ConnectionRequest(ByVal requestID As Long)
    If WS6.State <> sckClosed Then
        WS6.Close
    End If
'    lblWs.Visible = True
    WS6.Accept requestID
    WS6.SendData "OK"
    
    Call ClosePortWinsock
'    WS6.Close
'    WS6.LocalPort = 2006
'    WS6.Listen
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub ws7_ConnectionRequest(ByVal requestID As Long)
    If ws7.State <> sckClosed Then
        ws7.Close
    End If
    ws7.Accept requestID
    ws7.SendData "OK"
    
    Call ClosePortWinsock
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub
Private Sub ws8_ConnectionRequest(ByVal requestID As Long)
    If ws8.State <> sckClosed Then
        ws8.Close
    End If
    ws8.Accept requestID
    ws8.SendData "OK"
    
    Call ClosePortWinsock
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub
Private Sub ws9_ConnectionRequest(ByVal requestID As Long)
    If ws9.State <> sckClosed Then
        ws9.Close
    End If
    ws9.Accept requestID
    ws9.SendData "OK"
    
    Call ClosePortWinsock
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub
Private Sub ws10_ConnectionRequest(ByVal requestID As Long)
    If ws10.State <> sckClosed Then
        ws10.Close
    End If
    ws10.Accept requestID
    ws10.SendData "OK"
    
    Call ClosePortWinsock
    
    lblWs.Tag = "1"
    If reload = False Then Exit Sub
    Call loadAntrian
    Call LoadTrigger
End Sub

Private Sub lblconn_DblClick()
    End
End Sub



Private Sub loadAntrian()
On Error Resume Next
Dim disada As Boolean

'    If reload = False Then Exit Sub
'    reload = True
    lblWs.Visible = True
    'Set RS = Nothing
    strSQL = "select* from antrianpasienregistrasi_t  where statuspanggil ='1' and tanggalreservasi between '" & Format(Now(), "yyyy-mm-dd 00:00") & "' and '" & Format(Now(), "yyyy-mm-dd 23:59") & "'"
    Call ReadRs(strSQL)
'    For i = 0 To 4
'        lbl(i).Caption = "-"
'    Next
    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        For i = 0 To RS.RecordCount - 1
'            If rs!JenisPasien = "BPJS" Then
'                jenisAntrian = 1
'            Else
'                jenisAntrian = 2
'            End If
            If RS!tempatlahir = 1 Then
                If lbl(0).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(0).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(0).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 1
            End If
            If RS!tempatlahir = 2 Then
                If lbl(1).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(1).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(1).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 2
            End If
            If RS!tempatlahir = 3 Then
                If lbl(2).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(2).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(2).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 3
            End If
            If RS!tempatlahir = 4 Then
                If lbl(3).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(3).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(3).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 4
            End If
            If RS!tempatlahir = 5 Then
                If lbl(4).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(4).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 5
            End If
            If RS!tempatlahir = 6 Then
                If lbl(5).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(5).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 6
            End If
            If RS!tempatlahir = 7 Then
                If lbl(6).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(6).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 7
            End If
            If RS!tempatlahir = 8 Then
                If lbl(7).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(7).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 8
            End If
            If RS!tempatlahir = 9 Then
                If lbl(8).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(8).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 9
            End If
            If RS!tempatlahir = 10 Then
                If lbl(9).Caption <> RS!jenis & "-" & Format(RS!noantrian, "0##") Then disada = True
                'lbl(4).Caption = rs!JenisPasien & " : " & Format(rs!NoAntrian, "0##")
                lbl(9).Caption = RS!jenis & "-" & Format(RS!noantrian, "0##")
                loket = 10
            End If
            
            If disada = True Then Call PlaySound(RS!noantrian, UCase(RS!jenis))
            disada = False
            RS.MoveNext
        Next
    End If
    onload = False
    If Timer1.Enabled = False Then Timer1.Enabled = True
End Sub



Private Sub PlaySound(angka As Integer, jenis As String)
Dim t As Single
Dim belas As Boolean
Dim puluh As Boolean
Dim ratus As Boolean

    If onload = True Then Exit Sub
    lbl(loket - 1).BackColor = &H8080FF
    lbl(loket - 1).BackStyle = 1
    Timer2.Enabled = True
    Call sndPlaySound(App.Path & "\sound\nomor-urut.wav", SND_ASYNC Or SND_NODEFAULT)
    
    t = Timer
    Do
        DoEvents
    Loop Until Timer - t > 2
    
    If jenis = "A" Then Call sndPlaySound(App.Path & "\sound\AA.wav", SND_ALIAS Or SND_SYNC)
    If jenis = "B" Then Call sndPlaySound(App.Path & "\sound\BB.wav", SND_ALIAS Or SND_SYNC)
    If jenis = "C" Then Call sndPlaySound(App.Path & "\sound\CC.wav", SND_ALIAS Or SND_SYNC)
    If jenis = "D" Then Call sndPlaySound(App.Path & "\sound\DD.wav", SND_ALIAS Or SND_SYNC)
    If jenis = "E" Then Call sndPlaySound(App.Path & "\sound\EE.wav", SND_ALIAS Or SND_SYNC)
    

'    t = Timer
'    Do
'        DoEvents
'    Loop Until Timer - t > 1
    
    belas = False
    puluh = False
    ratus = False
    
    If angka > 199 And angka < 1000 Then ratus = True
    If angka > 99 And angka < 200 Then Call sndPlaySound(App.Path & "\sound\seratus.wav", SND_ALIAS Or SND_SYNC): angka = angka - 100
    If angka > 19 And angka < 100 Then puluh = True
    
    If angka < 20 And angka > 11 Then angka = angka - 10: belas = True
    
    If Len(CStr(angka)) = 2 And angka = 10 Then Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    If Len(CStr(angka)) = 2 And angka = 11 Then Call sndPlaySound(App.Path & "\sound\sebelas.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    If Len(CStr(angka)) = 3 And angka = 100 Then Call sndPlaySound(App.Path & "\sound\seratus.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    If Len(CStr(angka)) = 4 And angka = 1000 Then Call sndPlaySound(App.Path & "\sound\seribu.wav", SND_ALIAS Or SND_SYNC): GoTo loketttt
    
    For i = 1 To Len(CStr(angka))
        If ratus = False And angka > 200 And Val(Mid(angka, 2, 2)) = 10 Then Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC): Exit For
        If ratus = False And angka > 200 And Val(Mid(angka, 2, 2)) = 11 Then Call sndPlaySound(App.Path & "\sound\sebelas.wav", SND_ALIAS Or SND_SYNC): Exit For
        If ratus = False And angka > 200 Then
            If Val(Mid(angka, 2, 2)) > 19 And puluh = False Then
                puluh = True
            Else
                puluh = False
            End If
        End If
        If ratus = False And angka > 200 And Val(Mid(angka, 2, 2)) < 20 And Val(Mid(angka, 2, 2)) > 11 Then
            'If Val(Mid(angka, 2, 2)) < 20 And Val(Mid(angka, 2, 2)) > 11 Then belas = True:
            Select Case Val(Mid(angka, 2, 2))
                Case 12
                    Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 13
                    Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 14
                    Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 15
                    Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 16
                    Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 17
                    Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 18
                    Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
                Case 19
                    Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
                    Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
            End Select
            Exit For
        End If
        Select Case Mid(CStr(angka), i, 1)
           Case 1
               Call sndPlaySound(App.Path & "\sound\satu.wav", SND_ALIAS Or SND_SYNC)
           Case 2
               Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
           Case 3
               Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
           Case 4
               Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
           Case 5
               Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
           Case 6
               Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
           Case 7
               Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
           Case 8
               Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
           Case 9
               Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
        End Select
        

'        If ratus = False And angka > 200 Then
'            If Val(Mid(angka, 2, 2)) < 20 And Val(Mid(angka, 2, 2)) > 11 Then belas = True
'        End If

        If belas = True Then Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
        If puluh = True Then Call sndPlaySound(App.Path & "\sound\puluh.wav", SND_ALIAS Or SND_SYNC) ': puluh = False
        If angka > 19 And angka < 100 Then puluh = False
        If ratus = True Then Call sndPlaySound(App.Path & "\sound\ratus.wav", SND_ALIAS Or SND_SYNC): ratus = False
belas:
    Next
    
    
'    Select Case angka
'        Case 1
'            Call sndPlaySound(App.Path & "\sound\satu.wav", SND_ALIAS Or SND_SYNC)
'        Case 2
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'        Case 3
'            Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
'        Case 4
'            Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
'        Case 5
'            Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
'        Case 6
'            Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
'        Case 7
'            Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
'        Case 8
'            Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
'        Case 9
'            Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
'        Case 10
'            Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC)
'        Case 11
'            Call sndPlaySound(App.Path & "\sound\sebelas.wav", SND_ALIAS Or SND_SYNC)
'        Case 12
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 13
'            Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 14
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 15
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 16
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 17
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 18
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 19
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'        Case 20
'            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
'            Call sndPlaySound(App.Path & "\sound\belas.wav", SND_ALIAS Or SND_SYNC)
'     End Select
     
     
'1000-turun.wav
'100-turun.wav
'10-turun.wav
'11-turun.wav
'anak.wav
'Antrian.wav
'antrian2.wav
'belas.wav
'belas -turun.wav
'dalam -turun.wav
'delapan.wav
'dua.wav
'empat.wav
'enam.wav
'ke poli.wav
'kosong.wav
'lima.wav
'loket.wav
'LoketPendaftaran.wav
'nol.wav
'nomor -urut.wav
'puluh.wav
'puluh -turun.wav
'ratus.wav
'ratus -turun.wav
'ribu -turun.wav
'Satu.wav
'se.wav
'sebelas.wav
'sembilan.wav
'sepuluh.wav
'seratus.wav
'seribu.wav
'tiga.wav
'tujuh.wav
   
    
    
    
'    t = Timer
'    Do
'        DoEvents
'    Loop Until Timer - t > 1
loketttt:
    Call sndPlaySound(App.Path & "\sound\loket.wav", SND_ASYNC Or SND_NODEFAULT)
    
    t = Timer
    Do
        DoEvents
    Loop Until Timer - t > 1
    Select Case loket
        Case 1
            Call sndPlaySound(App.Path & "\sound\satu.wav", SND_ALIAS Or SND_SYNC)
        Case 2
            Call sndPlaySound(App.Path & "\sound\dua.wav", SND_ALIAS Or SND_SYNC)
        Case 3
            Call sndPlaySound(App.Path & "\sound\tiga.wav", SND_ALIAS Or SND_SYNC)
        Case 4
            Call sndPlaySound(App.Path & "\sound\empat.wav", SND_ALIAS Or SND_SYNC)
        Case 5
            Call sndPlaySound(App.Path & "\sound\lima.wav", SND_ALIAS Or SND_SYNC)
        Case 6
            Call sndPlaySound(App.Path & "\sound\enam.wav", SND_ALIAS Or SND_SYNC)
        Case 7
            Call sndPlaySound(App.Path & "\sound\tujuh.wav", SND_ALIAS Or SND_SYNC)
        Case 8
            Call sndPlaySound(App.Path & "\sound\delapan.wav", SND_ALIAS Or SND_SYNC)
        Case 9
            Call sndPlaySound(App.Path & "\sound\sembilan.wav", SND_ALIAS Or SND_SYNC)
        Case 10
            Call sndPlaySound(App.Path & "\sound\sepuluh.wav", SND_ALIAS Or SND_SYNC)
    End Select
End Sub

Private Sub Timer2_Timer()
'    lbl(KedipLoket).FontBold = Not lbl(KedipLoket).FontBold
    tmt2 = tmt2 + 1
    If tmt2 > 10 Then
        Timer2.Enabled = False
        For i = 0 To 9
'            lbl(i).BackColor = &H8000000F
            lbl(i).BackStyle = 0
        Next
        tmt2 = 0
'        lblWs.Visible = False
    End If
End Sub

Private Sub Timer3_Timer()
'    If runText.Left < 0 - runText.Width Then runText.Left = 1368 'Screen.Width
'    runText.Move runText.Left - 30
    
End Sub

Private Sub Timer4_Timer()
    tmt3 = tmt3 + 1
'    If tmt3 > 60 Then
'        strSQL = "select distinct noantrian from AntrianPasienRegistrasi where TglAntrian > '" & Format(Now(), "yyyy-mm-dd") & " 00:00' and JenisPasien = 'bpjs' " '  group by jenispasien"
'        Call msubRecFO(rsa, strSQL)
'        If rsa.RecordCount <> 0 Then
'            lbl2(0).Caption = "TOTAL BPJS : " & rsa.RecordCount
'        End If
'        strSQL = "select distinct noantrian from AntrianPasienRegistrasi where TglAntrian > '" & Format(Now(), "yyyy-mm-dd") & " 00:00' and JenisPasien = 'UMUM'   " 'group by jenispasien"
'        Call msubRecFO(rsa, strSQL)
'        If rsa.RecordCount <> 0 Then
'            lbl2(1).Caption = "TOTAL UMUM : " & rsa.RecordCount
'        End If
'        tmt3 = 0
'    End If
    lblJam.Caption = Format(Now(), "hh:nn:ss")
'    If Val(Format(Now(), "ss")) Mod 10 = 0 Then
'        'pic.Picture = File1.Path & "\File1.Tag"
'        pic.Picture = LoadPicture(File1.Path & "\" & File1.List(Val(File1.Tag)))
'        File1.Tag = Val(File1.Tag) + 1
'        If Val(File1.Tag) > File1.ListCount - 1 Then File1.Tag = 0
'    End If
    On Error GoTo Error_Handler
'    Label3.Caption = Round(DirectShow_Position.CurrentPosition, 0) & "/" & Round(DirectShow_Position.StopTime, 0)
'    If DirectShow_Position.CurrentPosition >= DirectShow_Position.StopTime Then
'            'DirectShow_Position.CurrentPosition = 0
'        vdeo = vdeo + 1
'        If vdeo > File1.ListCount - 1 Then vdeo = 0
'        DirectShow_Load_Media App.Path & "\video\" & File1.List(vdeo)
''    DirectShow_Loop
'        DirectShow_Play
'        DirectShow_Volume sora
'    End If
    If WindowsMediaPlayer1.Controls.CurrentPosition + 2 > WindowsMediaPlayer1.currentMedia.Duration Then
        vdeo = vdeo + 1
        If vdeo > File1.ListCount - 1 Then vdeo = 0
        'WindowsMediaPlayer1.URL = App.Path & "\video\" & File1.List(0)
        WindowsMediaPlayer1.URL = App.Path & "\video\" & File1.List(vdeo)
    End If
    Label1.Caption = WindowsMediaPlayer1.Controls.CurrentPosition
Error_Handler:
End Sub



