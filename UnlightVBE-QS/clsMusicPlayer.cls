VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMusicPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Player As IMediaControl   'Reference to our player
Private PlayerPos As IMediaPosition 'Reference to determine media position
Private PlayerAU As IBasicAudio     'Reference to determine Audio Volume
Private m_Filepath As String  '檔案路徑
Private m_IsLoop As Boolean '是否循環播放
Private m_Volume As Integer '音量
Private m_Mute As Boolean '是否靜音
Private PlayerIsPlaying As Boolean
Private PlayerIsRender As Boolean
Public WithEvents Timerplay As Timer
Attribute Timerplay.VB_VarHelpID = -1
Public Property Get Filepath() As String
   Filepath = m_Filepath
End Property
Public Property Let Filepath(ByVal New_Filepath As String)
   m_Filepath = New_Filepath
'   PlayerIsRender = False
   '=================
    Set Player = New FilgraphManager
    Set PlayerAU = Player
    Set PlayerPos = Player
    Player.RenderFile Me.Filepath
    PlayerIsRender = True
End Property
Public Property Get IsLoop() As Boolean
   IsLoop = m_IsLoop
End Property
Public Property Let IsLoop(ByVal New_IsLoop As Boolean)
   m_IsLoop = New_IsLoop
End Property
Public Property Get Mute() As Boolean
   Mute = m_Mute
End Property
Public Property Let Mute(ByVal New_Mute As Boolean)
   m_Mute = New_Mute
End Property
Public Property Get Volume() As Integer
   Volume = m_Volume
End Property
Public Property Let Volume(ByVal New_Volume As Integer)
   m_Volume = New_Volume
   '========================
    If m_Mute = False And PlayerIsPlaying = True Then
         Call AdjustVolume
    End If
End Property
Public Sub MusicPlay()
'Set Player = New FilgraphManager
'Set PlayerAU = Player
'Set PlayerPos = Player
'Player.RenderFile Me.Filepath
'PlayerIsRender = True
Call AdjustVolume
Player.Run
PlayerIsPlaying = True
If m_IsLoop = True Then
    If Not Timerplay Is Nothing Then
        Timerplay.Interval = 1000
        Timerplay.Enabled = True
    End If
End If
End Sub
Public Sub MusicStop()
If Not Player Is Nothing Then Player.Stop
PlayerIsPlaying = False
PlayerIsRender = False
If Not Timerplay Is Nothing Then
    Timerplay.Enabled = False
End If
End Sub

Private Sub AdjustVolume()
If PlayerIsRender = True Then
    If m_Mute = True Then
        PlayerAU.Volume = -10000
    Else
        PlayerAU.Volume = (m_Volume * 40) - 4000
    End If
End If
End Sub

Private Sub Class_Initialize()
Set Player = New FilgraphManager
Set PlayerAU = Player
Set PlayerPos = Player
PlayerIsPlaying = False
PlayerIsRender = False
End Sub

Private Sub Class_Terminate()
Set Player = Nothing
Set PlayerAU = Nothing
Set PlayerPos = Nothing
End Sub

Private Sub Timerplay_Timer()
If PlayerIsPlaying = True And PlayerPos.CurrentPosition > 0 Then
    If PlayerPos.CurrentPosition >= PlayerPos.Duration Then
        Me.MusicStop
        If Me.IsLoop = True Then
            Me.MusicPlay
        End If
    End If
End If
End Sub
