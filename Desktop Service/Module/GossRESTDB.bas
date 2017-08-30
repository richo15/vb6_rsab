Attribute VB_Name = "GossRESTDB"
Option Explicit

Private Const MDB_NAME As String = "Data\Movies.mdb"
Private Const TXT_DIR As String = "Data"
Private Const TXT_NAME As String = "Movies.txt"
Private Const PROVIDER_STRING As String = _
      "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;" _
    & "Mode=Share Exclusive;Data Source='$MDB$'"

'Initial letters of "number words" for digits 0 through 9 in English:
Private Const NUMBER_LETTERS As String = "ZOTTFFSSEN"

'Searches a string for the first occurrence of any of a group of characters.
Private Declare Function StrCSpn Lib "shlwapi" Alias "StrCSpnW" ( _
    ByVal pszStr As Long, _
    ByVal pszSet As Long) As Long

Private mConnectionString As String

Public Property Get ConnectionString() As String
    ConnectionString = mConnectionString
End Property

Private Property Let ConnectionString(ByVal RHS As String)
    mConnectionString = RHS
End Property

Private Function InCharSet( _
    ByVal Start As Long, _
    ByRef Text As String, _
    ByVal CharSet As String) As Long
    'Returns 1-based character offset of first character in CharSet,
    'or 0 if none found, using charset in Delimiters.

    If Start <= Len(Text) Then
        InCharSet = StrCSpn(&H80000000 Xor ((StrPtr(Text) Xor &H80000000) + (Start - 1) * 2), _
                            StrPtr(CharSet))
        If InCharSet = Len(Text) + 1 - Start Then
            InCharSet = 0
        Else
            InCharSet = Start + InCharSet
        End If
    End If
End Function

Private Sub CreateDb()
'    Dim CN As ADODB.Connection
'    Dim rsText As ADODB.Recordset
'    Dim rsTableDirect As ADODB.Recordset
'    Dim Fields As Variant
'    Dim Values As Variant
'    Dim Title As String
'    Dim Initials1 As String
'    Dim Initials2 As String
'
'    With CreateObject("ADOX.Catalog")
'        .Create ConnectionString
'        Set CN = .ActiveConnection
'    End With
'    With CN
'        .Execute "CREATE TABLE [Movies](" _
'               & "[ID] LONG CONSTRAINT [pkID] PRIMARY KEY," _
'               & "[Year] SHORT," _
'               & "[Title] TEXT(255) WITH COMPRESSION NOT NULL," _
'               & "[Initials1] TEXT(127) WITH COMPRESSION NOT NULL," _
'               & "[Initials2] TEXT(127) WITH COMPRESSION NOT NULL)", _
'                 , _
'                 adCmdText Or adExecuteNoRecords
'        '-----------------------------------------------
'        'Do a high-performance bulk load operation here:
'        Set rsTableDirect = New ADODB.Recordset
'        With rsTableDirect
'            Set .ActiveConnection = CN
'            .Properties("Append-Only Rowset").Value = True
'            'Do not use [] around the table name when opening direct:
'            .Open "Movies", , , adLockOptimistic, adCmdTableDirect
'        End With
'        Fields = Array(0, 1, 2, 3, 4)
'        ReDim Values(4)
'        Set rsText = .Execute("[Text;Database=" & TXT_DIR & ".;].[" & TXT_NAME & "]", , adCmdTable)
'        With rsText
'            Do Until .EOF
'                Values(0) = ![ID].Value
'                Values(1) = ![Year].Value
'                Values(2) = ![Title].Value
'                FormInitials Values(2), Values(3), Values(4)
'                rsTableDirect.AddNew Fields, Values
'                .MoveNext
'            Loop
'            .Close
'        End With
'        rsTableDirect.Close
'        '-----------------------------------------------
'        .Close
'    End With
End Sub

Private Function Zap(ByVal Text As String) As String
    'Replace chars in CHAR_SET (basically punctuation symbols) with a space.
    Const CHAR_SET As String = """#$%'*,-.:;?¡¿"
    Dim Position As Long
    
    Do
        Position = InCharSet(Position + 1, Text, CHAR_SET)
        If Position > 0 Then
            Mid$(Text, Position, 1) = " "
        End If
    Loop Until Position = 0
    Zap = Text
End Function

Public Sub FormInitials( _
    ByVal Title As String, _
    ByRef Initials1 As Variant, _
    ByRef Initials2 As Variant)
    'Initials1 and Initials2 are Variant (String).
    Dim Words() As String
    Dim i As Long
    Dim Length1 As Long
    Dim Length2 As Long
    Dim Char As String
    
    Title = Trim$(Title)
    If Len(Title) < 1 Then Exit Sub
    Words = Split(Zap(UCase$(Title)), " ")
    Initials1 = Space$(UBound(Words) + 1)
    Initials2 = Space$(UBound(Words) + 1)
    For i = 0 To UBound(Words)
        If Len(Words(i)) > 0 Then
            Char = Left$(Words(i), 1)
            Select Case Char
                Case "A" To "Z"
                    Length1 = Length1 + 1
                    Mid$(Initials1, Length1, 1) = Char
                Case "0" To "9"
                    Length1 = Length1 + 1
                    Mid$(Initials1, Length1, 1) = Mid$(NUMBER_LETTERS, CLng(Char) + 1, 1)
            End Select
            Select Case Words(i)
                Case "A", "AN", "AND", "THE", "HALF", "LA", "LE", "L", "D"
                    'Skip these "words" for purposes of forming the "Initials2" values.
                Case Else
                    Select Case Char
                        Case "A" To "Z"
                            Length2 = Length2 + 1
                            Mid$(Initials2, Length2, 1) = Char
                        Case "0" To "9"
                            Length2 = Length2 + 1
                            Mid$(Initials2, Length2, 1) = Mid$(NUMBER_LETTERS, CLng(Char) + 1, 1)
                    End Select
            End Select
        End If
    Next
    Initials1 = Left$(Initials1, Length1)
    Initials2 = Left$(Initials2, Length2)
End Sub

Public Function AsInitials(ByVal Text As String) As String
    Dim i As Long
    Dim Char As String
    
    AsInitials = Zap(UCase$(Text))
    For i = 1 To Len(Text)
        Char = Mid$(Text, i, 1)
        Select Case Char
            Case "0" To "9"
                Mid$(AsInitials, i, 1) = Mid$(NUMBER_LETTERS, CLng(Char) + 1, 1)
            Case " "
                AsInitials = Left$(AsInitials, i - 1) & Mid$(AsInitials, i + 1)
                i = i + 1
        End Select
    Next
End Function

Public Sub InitializeDB()
'    ChDir App.Path
'    ChDrive App.Path
'    ConnectionString = Replace$(PROVIDER_STRING, "$MDB$", MDB_NAME)
'    On Error Resume Next
'    GetAttr MDB_NAME
'    If Err Then
'        On Error GoTo 0
'        CreateDb
'    End If
End Sub
