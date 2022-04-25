VERSION 5.00
Begin VB.UserControl ucMusicPlayer 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   HitBehavior     =   0  '�L
   ScaleHeight     =   885
   ScaleWidth      =   1605
   Windowless      =   -1  'True
   Begin VB.Timer Timerplay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   240
   End
End
Attribute VB_Name = "ucMusicPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Player As FilgraphManager   'Reference to our player
Dim PlayerPos As IMediaPosition 'Reference to determine media position
Dim PlayerAU As IBasicAudio     'Reference to determine Audio Volume
Dim m_Filepath As String  '�ɮ׸��|
Dim m_IsLoop As Boolean '�O�_�`������
Dim m_Volume As Integer '���q
Dim m_Mute As Boolean '�O�_�R��
Dim PlayerIsPlaying As Boolean
Dim PlayerIsRender As Boolean
Public Property Get Filepath() As String
   Filepath = m_Filepath
End Property
Public Property Let Filepath(ByVal New_Filepath As String)
   m_Filepath = New_Filepath
   PropertyChanged "Filepath"
   '===================
    Set Player = New FilgraphManager
    Set PlayerAU = Player
    Set PlayerPos = Player
    Player.RenderFile m_Filepath
    PlayerIsRender = True
End Property
Public Property Get IsLoop() As Boolean
   IsLoop = m_IsLoop
End Property
Public Property Let IsLoop(ByVal New_IsLoop As Boolean)
   m_IsLoop = New_IsLoop
   PropertyChanged "IsLoop"
End Property
Public Property Get Mute() As Boolean
   Mute = m_Mute
End Property
Public Property Let Mute(ByVal New_Mute As Boolean)
   m_Mute = New_Mute
   PropertyChanged "Mute"
   If PlayerIsPlaying = True Then
         Me.AdjustVolume
    End If
End Property
Public Property Get Volume() As Integer
   Volume = m_Volume
End Property
Public Property Let Volume(ByVal New_Volume As Integer)
   m_Volume = New_Volume
   PropertyChanged "Volume"
   '========================
    If m_Mute = False And PlayerIsPlaying = True Then
         Me.AdjustVolume
    End If
End Property
Public Sub MusicPlay()
Me.AdjustVolume
Player.Run
PlayerIsPlaying = True
Timerplay.Enabled = True
End Sub
Public Sub MusicStop()
Player.Stop
If PlayerPos.CurrentPosition > 0 Then PlayerPos.CurrentPosition = 0
PlayerIsPlaying = False
Timerplay.Enabled = False
PlayerIsRender = False
Me.Filepath = m_Filepath 'ReRender
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

Private Sub UserControl_Initialize()
Set Player = New FilgraphManager
Set PlayerAU = Player
Set PlayerPos = Player
PlayerIsPlaying = False
PlayerIsRender = False
End Sub
Sub AdjustVolume()
If PlayerIsRender = True Then
    If Me.Mute = True Then
        PlayerAU.Volume = -10000
    Else
        PlayerAU.Volume = (m_Volume * 40) - 4000
    End If
End If
End Sub
