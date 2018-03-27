VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl ucCard 
   Alignable       =   -1  'True
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   1845
   ScaleWidth      =   1185
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage cgu 
      Height          =   330
      Left            =   240
      Top             =   480
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Image           =   "ucCard.ctx":0000
      Props           =   5
   End
   Begin ImageX.aicAlphaImage cqu 
      Height          =   330
      Left            =   240
      Top             =   480
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Image           =   "ucCard.ctx":02E6
      Props           =   5
   End
   Begin ImageX.aicAlphaImage cge 
      Height          =   330
      Left            =   240
      Top             =   480
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Image           =   "ucCard.ctx":05C3
      Props           =   5
   End
   Begin ImageX.aicAlphaImage cqe 
      Height          =   330
      Left            =   240
      Top             =   480
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Image           =   "ucCard.ctx":0863
      Props           =   5
   End
   Begin ImageX.aicAlphaImage cardup2 
      Height          =   225
      Left            =   240
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   397
      Image           =   "ucCard.ctx":0AFD
      Props           =   5
   End
   Begin ImageX.aicAlphaImage cardup1 
      Height          =   225
      Left            =   240
      Top             =   0
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   397
      Image           =   "ucCard.ctx":1690
      Props           =   5
   End
   Begin ImageX.aicAlphaImage card 
      Height          =   1260
      Left            =   0
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   2223
      Image           =   "ucCard.ctx":2206
      Props           =   13
   End
End
Attribute VB_Name = "ucCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event CardClick()
Event CardButtonClickin()
Event CardButtonClickout()
Event CardMouseMove()
'=======================
Dim m_CardImage As String
Dim m_LocationType As Integer '牌目前狀態(0.未使用/1.手牌-正面/2.出牌-正面/3.背面)
Dim m_CardRotationType As Integer '牌目前反轉紀錄數(1.正面/2.轉牌)
Dim m_CardEventType As Boolean '牌是否顯示細項部分紀錄數
Dim m_CardEnabledType As Boolean '牌是否回應使用者操作紀錄數

Private Sub card_Click(ByVal Button As Integer)
If m_CardEnabledType = True Then RaiseEvent CardClick
End Sub

Sub card_MouseExit()
cardup1.Visible = False
cardup2.Visible = False
cge.Visible = False
cgu.Visible = False
cqu.Visible = False
cqe.Visible = False
End Sub

Private Sub card_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_CardEnabledType = True Then RaiseEvent CardMouseMove
If Me.CardEventType = True Then
    Select Case Me.LocationType
        Case 1
            cardup1.Visible = True
            cardup2.Visible = False
            cge.Visible = True
            cqe.Visible = False
        Case 2
            cardup2.Visible = True
            cardup1.Visible = False
            cqe.Visible = True
            cge.Visible = False
        Case Else
            cardup1.Visible = False
            cardup2.Visible = False
            cge.Visible = False
            cqe.Visible = False
    End Select
    cgu.Visible = False
    cqu.Visible = False
Else
    Call card_MouseExit
End If
End Sub

Private Sub cardup1_Click(ByVal Button As Integer)
Call card_Click(Button)
End Sub

Private Sub cardup1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call card_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cardup2_Click(ByVal Button As Integer)
Call card_Click(Button)
End Sub

Private Sub cardup2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call card_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cgu.Visible = True
cardup1.Visible = True
End Sub

Private Sub cgu_Click(ByVal Button As Integer)
If Me.CardRotationType = 1 Then
    Me.CardRotationType = 2
Else
    Me.CardRotationType = 1
End If
If m_CardEnabledType = True Then RaiseEvent CardButtonClickin
End Sub

Private Sub cqe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cqu.Visible = True
cardup2.Visible = True
End Sub

Private Sub cqu_Click(ByVal Button As Integer)
If Me.CardRotationType = 1 Then
    Me.CardRotationType = 2
Else
    Me.CardRotationType = 1
End If
If m_CardEnabledType = True Then RaiseEvent CardButtonClickout
End Sub

Public Property Get CardImage() As String
   CardImage = m_CardImage
End Property
Public Property Let CardImage(ByVal New_CardImage As String)
   m_CardImage = New_CardImage
   PropertyChanged "CardImage"
   '========================
   card.ClearImage
   card.LoadImage_FromFile Me.CardImage
   Me.LocationType = 1
   Me.CardRotationType = 1
End Property
Public Property Get LocationType() As Integer
   LocationType = m_LocationType
End Property
Public Property Let LocationType(ByVal New_LocationType As Integer)
   m_LocationType = New_LocationType
   PropertyChanged "LocationType"
   '==========================
   If Me.LocationType = 3 Then
       card.ClearImage
       card.ScaleMethod = aiActualSize
       card.Mirror = aiMirrorNone
       card.LoadImage_FromFile App.Path & "\card\cardback.png"
       card.Left = 0
       card.Top = 0
       Call card_MouseExit
   ElseIf Me.LocationType > 0 Then
       card.ClearImage
       card.LoadImage_FromFile m_CardImage
       card.Left = 0
       card.Top = 0
       Me.CardRotationType = m_CardRotationType
   End If
End Property
Public Property Get CardRotationType() As Integer
   CardRotationType = m_CardRotationType
End Property
Public Property Let CardRotationType(ByVal New_CardRotationType As Integer)
   m_CardRotationType = New_CardRotationType
   PropertyChanged "CardRotationType"
   '==========================
   Select Case Me.CardRotationType
       Case 1
          card.Mirror = aiMirrorNone
       Case 2
          card.Mirror = aiMirrorAll
   End Select
End Property
Public Property Get CardEventType() As Boolean
   CardEventType = m_CardEventType
End Property
Public Property Let CardEventType(ByVal New_CardEventType As Boolean)
   m_CardEventType = New_CardEventType
   PropertyChanged "CardEventType"
End Property
Public Property Get CardEnabledType() As Boolean
   CardEnabledType = m_CardEnabledType
End Property
Public Property Let CardEnabledType(ByVal New_CardEnabledType As Boolean)
   m_CardEnabledType = New_CardEnabledType
   PropertyChanged "CardEnabledType"
End Property
Private Sub UserControl_Initialize()
m_CardEnabledType = True
End Sub
