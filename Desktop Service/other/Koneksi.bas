Attribute VB_Name = "Koneksi"
'-------------------------------------
' edited by : agus.sustian
' date : 02 agustus 2017
' RSAB Harapan Kita
'-------------------------------------
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Sub openConnection()
 On Error GoTo NoConn
'On Error Resume Next
 
    With CN
        If .State = adStateOpen Then Exit Sub
        .CursorLocation = adUseClient
        
        .ConnectionString = "DRIVER={PostgreSQL Unicode};" & _
                            "SERVER=192.168.12.1;" & _
                            "port=5432;" & _
                            "DATABASE=rsab_hk_production;" & _
                            "UID=postgres;" & _
                            "PWD=root"
        '.ConnectionString = "DRIVER={PostgreSQL Unicode};" & _
                            "SERVER=192.168.12.4;" & _
                            "port=5432;" & _
                            "DATABASE=rsab_hk_development;" & _
                            "UID=postgres;" & _
                            "PWD=root"
        .ConnectionTimeout = 10
        .Open

        If CN.State = adStateOpen Then
        '    Connected sucsessfully"
        Else
            MsgBox "Koneksi ke database error, hubungi administrator !" & vbCrLf & Err.Description & " (" & Err.Number & ")"
            End
        End If
    End With
    

    Exit Sub
NoConn:
    MsgBox "Koneksi ke database error, ganti nama Server dan nama Database", vbCritical, "Validasi"
    End
'    frmSetServer.Show
'    blnError = True
'    Unload frmLogin
End Sub

