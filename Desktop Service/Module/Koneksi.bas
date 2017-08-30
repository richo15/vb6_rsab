Attribute VB_Name = "Koneksi"
'-------------------------------------
' edited by : agus.sustian
' date : 02 agustus 2017
' RSAB Harapan Kita
'-------------------------------------
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RS2 As New ADODB.Recordset
Public CN_String As String
Public strSQL As String
Public strSQL2 As String

Public StatusCN As String
Public Sub openConnection()
 On Error GoTo NoConn
 Dim host, Port, username, password, database As String
   host = GetTxt2("Setting.ini", "Koneksi", "a")
   Port = GetTxt2("Setting.ini", "Koneksi", "b")
   username = GetTxt2("Setting.ini", "Koneksi", "c")
   password = GetTxt2("Setting.ini", "Koneksi", "d")
   database = GetTxt2("Setting.ini", "Koneksi", "e")
'On Error Resume Next
 
    With CN
        If .State = adStateOpen Then Exit Sub
        .CursorLocation = adUseClient
        
        '.ConnectionString = "DRIVER={PostgreSQL Unicode};" & _
                            "SERVER=192.168.12.1;" & _
                            "port=5432;" & _
                            "DATABASE=rsab_hk_production;" & _
                            "UID=postgres;" & _
                            "PWD=root": StatusCN = "192.168.12.1"
        CN_String = "DRIVER={PostgreSQL Unicode};" & _
                            "SERVER=" & host & ";" & _
                            "port=" & Port & ";" & _
                            "DATABASE=" & database & ";" & _
                            "UID=" & username & ";" & _
                            "PWD=" & password & ""
        .ConnectionString = CN_String
        StatusCN = host
        .ConnectionTimeout = 10
        .Open

        If CN.State = adStateOpen Then
        '    Connected sucsessfully"
        Else
            MsgBox "Koneksi ke database error, hubungi administrator !" & vbCrLf & Err.Description & " (" & Err.Number & ")"
            frmSettingKoneksi.Show
        End If
    End With
    

    Exit Sub
NoConn:
    MsgBox "Koneksi ke database error, ganti nama Server dan nama Database", vbCritical, "Validasi"
    frmSettingKoneksi.Show
    
'    frmSetServer.Show
'    blnError = True
'    Unload frmLogin
End Sub

Public Function ReadRs(sql As String)
  Set RS = Nothing
  RS.Open sql, CN, adOpenStatic, adLockReadOnly
End Function

Public Function ReadRs2(sql As String)
  Set RS2 = Nothing
  RS2.Open sql, CN, adOpenStatic, adLockReadOnly
End Function
