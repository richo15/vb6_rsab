Attribute VB_Name = "Koneksi"
'-------------------------------------
' edited by : agus.sustian
' date : 02 agustus 2017
' RSAB Harapan Kita
'-------------------------------------
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RS2 As New ADODB.Recordset
Public RS3 As New ADODB.Recordset
Public strSQL As String
Public strSQL2 As String
Public strSQL3 As String

Public StatusCN As String
Public Sub openConnection()
 On Error GoTo NoConn
 Dim host, port, username, password, database As String
   host = GetTxt("Setting.ini", "Koneksi", "a")
   port = GetTxt("Setting.ini", "Koneksi", "b")
   username = GetTxt("Setting.ini", "Koneksi", "c")
   password = GetTxt("Setting.ini", "Koneksi", "d")
   database = GetTxt("Setting.ini", "Koneksi", "e")
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
        .ConnectionString = "DRIVER={PostgreSQL Unicode};" & _
                            "SERVER=" & host & ";" & _
                            "port=" & port & ";" & _
                            "DATABASE=" & database & ";" & _
                            "UID=" & username & ";" & _
                            "PWD=" & password & "": StatusCN = host
        .ConnectionTimeout = 10
        .Open

        If CN.State = adStateOpen Then
        '    Connected sucsessfully"
        Else
            MsgBox "Koneksi ke database error, hubungi administrator !" & vbCrLf & Err.Description & " (" & Err.Number & ")"
            frmSettingKoneksi.Show vbModal
        End If
    End With
    

    Exit Sub
NoConn:
    MsgBox "Koneksi ke database error, ganti nama Server dan nama Database", vbCritical, "Validasi"
    frmSettingKoneksi.Show vbModal
    
'    frmSetServer.Show
'    blnError = True
'    Unload frmLogin
End Sub
Sub Main()
    Call openConnection
    If CN.State = adStateOpen Then
        Form2.Show
    Else
        End
    End If
    
End Sub

Public Function ReadRs(sql As String)
  Set RS = Nothing
  RS.Open sql, CN, adOpenStatic, adLockPessimistic
End Function

Public Function ReadRs2(sql As String)
  Set RS2 = Nothing
  RS2.Open sql, CN, adOpenStatic, adLockPessimistic
End Function

Public Function ReadRs3(sql As String)
  Set RS3 = Nothing
  RS3.Open sql, CN, adOpenStatic, adLockPessimistic
End Function
