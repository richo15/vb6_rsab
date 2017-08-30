VERSION 5.00
Begin VB.Form SimpleClientMain 
   Caption         =   "Simple client of GossREST"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "SimpleClientMain"
   ScaleHeight     =   5865
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtJSON 
      Height          =   4395
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1440
      Width           =   5895
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Make query"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox txtQuery 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   60
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "JSON results:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Movie query:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "SimpleClientMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TxtQueryRightPadding As Single

Private WithEvents Req As WinHttp.WinHttpRequest
Attribute Req.VB_VarHelpID = -1

Private Sub cmdQuery_Click()
    MousePointer = vbHourglass
    cmdQuery.Enabled = False
    txtQuery.Locked = True
    With Req
        .Open "GET", "http://localhost:8080/query?" & Trim$(txtQuery.Text), Async:=True
        .Send
    End With
End Sub

Private Sub Form_Load()
    With txtQuery
        TxtQueryRightPadding = ScaleWidth - .Left - .Width
    End With
    Set Req = New WinHttp.WinHttpRequest
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        With txtQuery
            .Width = ScaleWidth - .Left - TxtQueryRightPadding
        End With
        With txtJSON
            .Move 0, .Top, ScaleWidth, ScaleHeight - .Top
        End With
    End If
End Sub

Private Sub Req_OnResponseFinished()
    If Req.Status = 200 Then
        With New JNode
            .JSON = Req.ResponseText
            'We could do more, but we'll just dump it to be read:
            txtJSON.Text = .JSON("    ")
        End With
    Else
        txtJSON.Text = "Status " & CStr(Req.Status) & " " & Req.StatusText
    End If
    With txtQuery
        .SelStart = 0
        .SelLength = &H7FFF
        .SetFocus
        .Locked = False
    End With
    cmdQuery.Enabled = True
    MousePointer = vbDefault
End Sub
