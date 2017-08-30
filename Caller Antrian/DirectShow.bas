Attribute VB_Name = "DirectShow"
Option Explicit

Private Const MAX_VOLUME As Long = 100
Private Const MAX_BALANCE As Long = 100
Private Const MAX_SPEED As Long = 226

Public DirectShow_Event As IMediaEvent
Public DirectShow_Control As IMediaControl
Public DirectShow_Position As IMediaPosition
Public DirectShow_Audio As IBasicAudio
Public DirectShow_Video As IBasicVideo
Public DirectShow_Video_Window As IVideoWindow

Public Video_Media As Boolean
Public Video_Running As Boolean
Public Fullscreen_Enabled As Boolean
Public Fullscreen_Width As Long
Public Fullscreen_Height As Long


Public Function DirectShow_Load_Media(File_Name As String) As Boolean

    On Error GoTo Error_Handler
                
        If Right(File_Name, 4) = ".mp3" Then
                    Video_Media = False
                            Set DirectShow_Control = New FilgraphManager
            DirectShow_Control.RenderFile (File_Name)
        
            Set DirectShow_Audio = DirectShow_Control
            
            DirectShow_Audio.Volume = 0
            DirectShow_Audio.Balance = 0
        
            Set DirectShow_Event = DirectShow_Control
            Set DirectShow_Position = DirectShow_Control
            
            DirectShow_Position.Rate = 1
            
            DirectShow_Position.CurrentPosition = 0
            
        ElseIf Right(File_Name, 4) = ".mpeg" Or _
               Right(File_Name, 4) = ".mpg" Or _
               Right(File_Name, 4) = ".avi" Or _
               Right(File_Name, 4) = ".mp4" Or _
               Right(File_Name, 4) = ".mov" Then
                           Video_Media = True
                           Set DirectShow_Control = New FilgraphManager
            DirectShow_Control.RenderFile (File_Name)

            Set DirectShow_Audio = DirectShow_Control
    
            DirectShow_Audio.Volume = 0
            DirectShow_Audio.Balance = 0

            If Fullscreen_Enabled = True Then
                                Set DirectShow_Video_Window = DirectShow_Control
                DirectShow_Video_Window.WindowStyle = CLng(&H6000000)
                DirectShow_Video_Window.Top = 0
                DirectShow_Video_Window.Left = 0
                DirectShow_Video_Window.Width = Fullscreen_Width
                DirectShow_Video_Window.Height = Fullscreen_Height
                DirectShow_Video_Window.Owner = Form1.hWnd
                
            Else
                Set DirectShow_Video_Window = DirectShow_Control
                DirectShow_Video_Window.WindowStyle = CLng(&H6000000)
                DirectShow_Video_Window.Top = DS_top ' 100
                DirectShow_Video_Window.Left = DS_left ' 400
                DirectShow_Video_Window.Width = DS_width ' 700 'frmMain.ScaleWidth
                DirectShow_Video_Window.Height = DS_height ' 500 'frmMain.ScaleHeight
                DirectShow_Video_Window.Owner = Form1.hWnd
            
            End If
                    Set DirectShow_Event = DirectShow_Control
            Set DirectShow_Position = DirectShow_Control
            
            DirectShow_Position.Rate = 1
            
            DirectShow_Position.CurrentPosition = 0
               
        Else
                    GoTo Error_Handler
        
        End If

    DirectShow_Load_Media = True
        Exit Function
Error_Handler:

    DirectShow_Load_Media = False

End Function


Public Function DirectShow_Play() As Boolean
        On Error GoTo Error_Handler
    
    If Video_Media = True Then Video_Running = True
        DirectShow_Control.Run

    DirectShow_Play = True
        Exit Function

Error_Handler:
    
    DirectShow_Play = False

End Function

Public Function DirectShow_Stop() As Boolean

    On Error GoTo Error_Handler
    
    If Video_Media = True Then
            Video_Running = False
            Video_Media = False
        End If
        DirectShow_Control.Stop
        DirectShow_Position.CurrentPosition = 0

    DirectShow_Stop = True
        Exit Function

Error_Handler:

    DirectShow_Stop = False

End Function

Public Function DirectShow_Pause() As Boolean

    On Error GoTo Error_Handler
    
    DirectShow_Control.Stop

    DirectShow_Pause = True
        Exit Function
Error_Handler:
    
    DirectShow_Pause = False

End Function

Public Function DirectShow_Volume(ByVal Volume As Long) As Boolean

    On Error GoTo Error_Handler
    
    If Volume >= MAX_VOLUME Then Volume = MAX_VOLUME
    
    If Volume <= 0 Then Volume = 0
    
    DirectShow_Audio.Volume = (Volume * MAX_VOLUME) - 10000

    DirectShow_Volume = True
        Exit Function
Error_Handler:

    DirectShow_Volume = False

End Function

Public Function DirectShow_Balance(ByVal Balance As Long) As Boolean

    On Error GoTo Error_Handler
    
    If Balance >= MAX_BALANCE Then Balance = MAX_BALANCE
    
    If Balance <= -MAX_BALANCE Then Balance = -MAX_BALANCE
    
    DirectShow_Audio.Balance = Balance * MAX_BALANCE
    
    DirectShow_Balance = True
        Exit Function
Error_Handler:

    DirectShow_Balance = False

End Function

Public Function DirectShow_Speed(ByVal Speed As Single) As Boolean

    On Error GoTo Error_Handler

    If Speed >= MAX_SPEED Then Speed = MAX_SPEED
    
    If Speed <= 0 Then Speed = 0

    DirectShow_Position.Rate = Speed / 100

    DirectShow_Speed = True
        Exit Function

Error_Handler:

    DirectShow_Speed = False

End Function

Public Function DirectShow_Set_Position(ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, Milliseconds As Single) As Boolean
        On Error GoTo Error_Handler
    
    Dim Max_Position As Single
        Dim Position As Double
        Dim Decimal_Milliseconds As Single
        'Keep minutes within range
            Minutes = Minutes Mod 60
        
    'Keep seconds within range
            Seconds = Seconds Mod 60
        
    'Keep milliseconds within range and keep decimal
            Decimal_Milliseconds = Milliseconds - Int(Milliseconds)
        
        Milliseconds = Milliseconds Mod 1000
        
        Milliseconds = Milliseconds + Decimal_Milliseconds
    
    'Convert Minutes & Seconds to Position time
            Position = (Hours * 3600) + (Minutes * 60) + Seconds + (Milliseconds * 0.001)
    
    Max_Position = DirectShow_Position.StopTime

    If Position >= Max_Position Then
            Position = 0
        
        GoTo Error_Handler
    
    End If
        If Position <= 0 Then
            Position = 0
        
        GoTo Error_Handler
    
    End If
        DirectShow_Position.CurrentPosition = Position
    
    DirectShow_Set_Position = True
        Exit Function
Error_Handler:

    DirectShow_Set_Position = False

End Function

Public Function DirectShow_End() As Boolean

    On Error GoTo Error_Handler
    
    If DirectShow_Loop = False Then
            If DirectShow_Position.CurrentPosition >= DirectShow_Position.StopTime Then DirectShow_Stop
    
    End If
        DirectShow_End = True
        Exit Function
Error_Handler:

    DirectShow_End = False

End Function

Public Function DirectShow_Loop() As Boolean

    On Error GoTo Error_Handler

    If DirectShow_Position.CurrentPosition >= DirectShow_Position.StopTime Then
            DirectShow_Position.CurrentPosition = 0
    
    End If
        DirectShow_Loop = True
        Exit Function
Error_Handler:

    DirectShow_Loop = False

End Function

Public Sub DirectShow_Shutdown()

    Set DirectShow_Video_Window = Nothing
    Set DirectShow_Position = Nothing
    Set DirectShow_Event = Nothing
    Set DirectShow_Audio = Nothing
    Set DirectShow_Control = Nothing

End Sub


