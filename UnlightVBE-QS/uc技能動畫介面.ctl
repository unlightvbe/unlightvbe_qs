VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc技能動畫介面 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   CanGetFocus     =   0   'False
   ClientHeight    =   9915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   ClipBehavior    =   0  '無
   ClipControls    =   0   'False
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   9915
   ScaleWidth      =   11340
   Windowless      =   -1  'True
   Begin VB.Timer TimerObj 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10800
      Top             =   9360
   End
   Begin ImageX.aicAlphaImage aicImage 
      Height          =   9900
      Index           =   2
      Left            =   2880
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   17463
      Image           =   "uc技能動畫介面.ctx":0000
      Scaler          =   3
      Mirror          =   1
      Angle           =   90
      Props           =   141
      MaskColor       =   0
      ShadowDepth     =   10
   End
   Begin ImageX.aicAlphaImage aicImage 
      Height          =   9900
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   17463
      Image           =   "uc技能動畫介面.ctx":562F2
      Scaler          =   3
      Angle           =   90
      Props           =   141
      MaskUsed        =   -1  'True
      MaskColor       =   0
      MaskSource      =   1
      Mask            =   0
      ShadowDepth     =   10
   End
End
Attribute VB_Name = "uc技能動畫介面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_AnimatePictureList As Collection
Private m_MusicPlayerObj As ucMusicPlayer
Private m_uscom As Integer
Private m_ImageMaskUse As Boolean
Public Event AnimateCheckPoint(ByVal uscom As Integer)
Public Event AnimateEnd(ByVal uscom As Integer)
Private timernum As Integer
Private Sub Delay(ASecond As Double)
    Dim before
    before = Timer
    Do
        DoEvents
    Loop Until (Timer - before >= ASecond)
End Sub
Public Sub AnimateStart()
timernum = 0
If m_AnimatePictureList.Count() = 16 Then
    TimerObj.Interval = 30
Else
    TimerObj.Interval = 100
End If
TimerObj.Enabled = True
End Sub

Public Property Get AnimatePictureList() As Collection
    Set AnimatePictureList = m_AnimatePictureList
End Property

Public Property Let AnimatePictureList(ByVal vNewValue As Collection)
    If vNewValue.Count() > 0 Then
        Set m_AnimatePictureList = vNewValue
    Else
        Set m_AnimatePictureList = Nothing
    End If
    PropertyChanged "AnimatePictureList"
End Property

Public Property Get uscom() As Integer
    uscom = m_uscom
End Property

Public Property Let uscom(ByVal vNewValue As Integer)
    If vNewValue < 1 Or vNewValue > 2 Then vNewValue = 1
    m_uscom = vNewValue
    PropertyChanged "uscom"
End Property

Private Sub aicImage_FadeTerminated(Index As Integer, ByVal CurrentOpacity As Long)
If CurrentOpacity = 0 Then
    aicImage(m_uscom).Visible = True
    RaiseEvent AnimateEnd(m_uscom)
End If
End Sub

Private Sub TimerObj_Timer()
Call SetImageMaskUse
Select Case m_AnimatePictureList.Count()
    Case 1
        Select Case timernum
            Case 0
                aicImage(m_uscom).Opacity = 0
                aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(1)
                Call SetImageLeftTop(m_uscom)
                aicImage(m_uscom).Visible = True
                aicImage(m_uscom).FadeInOut 100, 20
            Case 1
                m_MusicPlayerObj.MusicStop
                m_MusicPlayerObj.MusicPlay
            Case 5
                RaiseEvent AnimateCheckPoint(m_uscom)
            Case 12
                TimerObj.Enabled = False
                aicImage(m_uscom).FadeInOut 0, 20
        End Select
    Case 2
        Select Case timernum
            Case 0
                aicImage(m_uscom).Opacity = 0
                aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(1)
                Call SetImageLeftTop(m_uscom)
                aicImage(m_uscom).Visible = True
                aicImage(m_uscom).FadeInOut 100, 20
            Case 1
                m_MusicPlayerObj.MusicStop
                m_MusicPlayerObj.MusicPlay
            Case 5
                aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(2)
                RaiseEvent AnimateCheckPoint(m_uscom)
            Case 12
                TimerObj.Enabled = False
                aicImage(m_uscom).FadeInOut 0, 20
        End Select
    Case 16
        Select Case timernum
            Case 0
                aicImage(m_uscom).Opacity = 0
                aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(1)
                Call SetImageLeftTop(m_uscom)
                aicImage(m_uscom).Visible = True
                aicImage(m_uscom).FadeInOut 100, 30
            Case 1
                aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(2)
                m_MusicPlayerObj.MusicStop
                m_MusicPlayerObj.MusicPlay
            Case 7
                aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(8)
                RaiseEvent AnimateCheckPoint(m_uscom)
            Case 16
                TimerObj.Enabled = False
                aicImage(m_uscom).FadeInOut 0, 30
            Case Else
                If timernum < 16 Then
                    aicImage(m_uscom).LoadImage_FromFile m_AnimatePictureList(timernum + 1)
                End If
        End Select
End Select
timernum = timernum + 1
End Sub

Public Property Get MusicPlayerObj() As ucMusicPlayer
    Set MusicPlayerObj = m_MusicPlayerObj
End Property

Public Property Let MusicPlayerObj(ByVal vNewValue As ucMusicPlayer)
    Set m_MusicPlayerObj = vNewValue
    PropertyChanged "MusicPlayerObj"
End Property

Private Sub UserControl_Show()
aicImage(1).Visible = False
aicImage(2).Visible = False
End Sub

Public Property Get ImageMaskUse() As Boolean
    ImageMaskUse = m_ImageMaskUse
End Property

Public Property Let ImageMaskUse(ByVal vNewValue As Boolean)
    m_ImageMaskUse = vNewValue
    PropertyChanged "ImageMaskUse"
    Call SetImageMaskUse
End Property
Private Sub SetImageLeftTop(ByVal uscom As Integer)
Select Case uscom
    Case 1
        aicImage(m_uscom).Left = 0
    Case 2
        aicImage(m_uscom).Left = 11340 - aicImage(m_uscom).Width
End Select
If aicImage(m_uscom).Height < 9900 - 100 Then
    aicImage(m_uscom).Top = 480
Else
    aicImage(m_uscom).Top = 0
End If
End Sub
Private Sub SetImageMaskUse()
If m_ImageMaskUse = True Then
    aicImage(1).MaskUsed = aiUseMaskColor
    aicImage(2).MaskUsed = aiUseMaskColor
Else
    aicImage(1).MaskUsed = aiNoMask
    aicImage(2).MaskUsed = aiNoMask
End If
End Sub
