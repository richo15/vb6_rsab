VERSION 5.00
Begin VB.Form GossRESTMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Service"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   4560
   Icon            =   "GossRESTMain.frx":0000
   LinkTopic       =   "GossRESTMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin DesktopService.Gossamer Gossamer1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      VDir            =   "VDir"
   End
   Begin VB.Label Label1 
      Caption         =   "PT-jasamedika_saranatama"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Menu asdasd 
      Caption         =   "File"
      Begin VB.Menu mnuSettingPrinter 
         Caption         =   "Setting Printer"
      End
      Begin VB.Menu asdasdasdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "GossRESTMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Initialized during Form_Load via SanitizeInit, used during Sanitize:
Private Const FROM_CHARS As String = "[]\[\]\'\;\%\_\#"
Private Const TO_CHARS As String = "[[]]\[[]\[]]\['']\\[%]\[_]\[#]"
Private FromChars() As String
Private ToChars() As String


Public STM As ADODB.Stream
Private LogFile As Integer

Private Sub SanitizeInit()
    FromChars = Split(FROM_CHARS, "\")
    ToChars = Split(TO_CHARS, "\")
End Sub

Private Function Sanitize(ByVal Arg As String) As String
    'Attempt to "sanitize" argument Arg against SQL Injection errors.
    Dim I As Long
    
    For I = 0 To UBound(FromChars)
        Arg = Replace$(Arg, FromChars(I), ToChars(I))
    Next
    Sanitize = Arg
End Function



Private Sub Command1_Click()
    
End Sub

Private Sub Form_Initialize()
    GossRESTDB.InitializeDB
End Sub

Private Sub Form_Load()
    Call openConnection
    SanitizeInit
'    Set CN = New ADODB.Connection
'    CN.Open GossRESTDB.ConnectionString
'    Set RS = New ADODB.Recordset
'    'Set up as a high-performance server-side cursor with the ability to
'    'be data-bound to an MSHFlexGrid:
'    With RS
'        .CursorLocation = adUseServer
'        .CursorType = adOpenKeyset
'        .LockType = adLockReadOnly
'        .ActiveConnection = CN
'        .Properties("IRowsetIdentity") = True
'    End With
    Set STM = New ADODB.Stream
    
    LogFile = FreeFile(0)
    Open "log.txt" For Append As #LogFile
    Gossamer1.StartListening
    
'    Show
    WindowState = vbHide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Gossamer1.StopListening
    Close #LogFile
'    CN.Close
End Sub



Private Sub Gossamer1_DynamicRequest( _
    ByVal Method As String, _
    ByVal URI As String, _
    ByVal Params As String, _
    ByVal ReqHeaders As Collection, _
    ByRef RespStatus As Single, _
    ByRef RespStatusText As String, _
    ByRef RespMIME As String, _
    ByRef RespExtraHeaders As String, _
    ByRef RespBody() As Byte, _
    ByVal ClientIndex As Integer)
    
    Dim ErrNumber As Long
    Dim ErrDescription As String
    
    If Method = "GET" Then
        'We'll assume URI = "/query" but we'll take any as "/query" in this program.
        On Error Resume Next

        Dim Itm As Variant
        For Each Itm In ReqHeaders
            Debug.Print Itm(0), Itm(1)
        Next Itm
        If URI = "\query" Then RespBody = Query(Gossamer1.URLDecode(Params))
        If URI = "\cetak-antrian" Then RespBody = frmCetakAntrian.CetakAntrian(Gossamer1.URLDecode(Params))
        If URI = "\kasir" Then RespBody = frmKasir.Kasir(Gossamer1.URLDecode(Params))
        If URI = "\tata-rekening" Then RespBody = frmTataRekening.TataRekening(Gossamer1.URLDecode(Params))
        If Err Then
            ErrNumber = Err.Number
            ErrDescription = Err.Description
            On Error GoTo 0
            RespStatus = 500
            RespStatusText = "Internal Server Error"
            Print #LogFile, _
                  "Error "; _
                  CStr(ErrNumber); _
                  " (&H"; Right$("0000000" & Hex$(ErrNumber), 8); ") "; _
                  ErrDescription
            Exit Sub
        End If
        On Error GoTo 0
        RespStatus = 200
        RespStatusText = "ok"
        RespMIME = "application/json" '; charset=utf-8"
        RespExtraHeaders = "Allow: GET" & vbCrLf
        'RespExtraHeaders = "Host: 172.16.16.14:8200"
    Else
        RespStatus = 405
        RespStatusText = "Method Not Allowed"
        RespExtraHeaders = "Allow: GET" & vbCrLf
        
    End If
End Sub

Private Sub Gossamer1_LogEvent(ByVal GossEvent As GossEvent, ByVal ClientIndex As Integer)
    With GossEvent
        Print #LogFile, _
              Format$(.Timestamp, "YYYY-MM-DD HH:NN:SS, "); _
              CStr(ClientIndex); ", "; _
              .IP; ", "; _
              CStr(.EventType); ", "; _
              CStr(.EventSubtype); ", "; _
              .Method; ", "; _
              .Text
    End With
End Sub
Private Function Query(ByVal QueryText As String) As Byte()
'    Const QUERY_TEMPLATE As String = _
'            "SELECT * FROM [Movies] " _
'          & "WHERE " _
'          & "[Title] = '$1$' OR " _
'          & "[Title] LIKE '$1$[ ,.:?]%' OR " _
'          & "[Title] LIKE '%[ ,.:?]$1$' OR " _
'          & "[Title] LIKE '%[ ,.:?]$1$[ ,.:?]%' OR " _
'          & "[Initials1] = '$2$' OR " _
'          & "[Initials2] = '$2$' " _
'          & "ORDER BY [Title] ASC"
'    Dim ActualQuery As String
    Dim Root As JNode
'    Dim RecordIndex As Long
'    Dim FieldIndex As Long
    
'    If Len(QueryText) = 0 Then
'        'Shouldn't match anything, just get an empty Recordset so
'        'we have the field names for headings:
'        ActualQuery = Replace$(Replace$(QUERY_TEMPLATE, _
'                                        "$2$", _
'                                        vbFormFeed), _
'                               "$1$", _
'                               vbFormFeed)
'    Else
'        ActualQuery = Replace$(Replace$(QUERY_TEMPLATE, _
'                                        "$2$", _
'                                        GossRESTDB.AsInitials(QueryText)), _
'                               "$1$", _
'                               Sanitize(QueryText))
'    End If
'    Debug.Print ActualQuery
'    With RS
'        .Open ActualQuery, , , , adCmdText
'        Set Root = New JNode
'        Root("Perusahaan") = "Jasamedika"
'        Set Root("Pegawai") = New JNode
'        Root("Pegawai")(0) = "Riko"
'        Root("Pegawai")(1) = "Agus"
'        For FieldIndex = 0 To .Fields.Count - 1
'            Root("ColumnNames")(FieldIndex) = .Fields(FieldIndex).Name
'        Next
'        If RS.RecordCount > 0 Then
'            Set Root("Rows") = New JNode
'            Root("Rows").MakeArray
'            Do Until .EOF
'                Set Root("Rows")(RecordIndex) = New JNode
'                Root("Rows")(RecordIndex).MakeArray
'                For FieldIndex = 0 To .Fields.Count - 1
'                    Root("Rows")(RecordIndex).Value(FieldIndex) = .Fields(FieldIndex).Value
'                Next
'                .MoveNext
'                RecordIndex = RecordIndex + 1
'            Loop
'        End If
'        .Close
'    End With

If Len(QueryText) > 0 Then
    
    If QueryText = "cetak" Then
        Set Root = New JNode
        'MsgBox "adasdasd", vbInformation, "CETAK SERVICE"
        Root("Status") = "Sukses Dicetak"
'        Set Root("Pegawai") = New JNode
'        Root("Pegawai")(0) = "Riko"
'        Root("Pegawai")(1) = "Agus"
    Else
        Set Root = New JNode
        Root("Status") = "Error"
    End If
End If
    With STM
        .Open
        .Type = adTypeText
        .CharSet = "utf-8"
        .WriteText Root.JSON, adWriteChar
        .Position = 0
        .Type = adTypeBinary
        Query = .Read(adReadAll)
        .Close
    End With
End Function

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSettingPrinter_Click()
    frmSettingPrinter.Show vbModal
End Sub
