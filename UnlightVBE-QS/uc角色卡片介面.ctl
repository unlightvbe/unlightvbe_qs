VERSION 5.00
Begin VB.UserControl uc角色卡片介面 
   Appearance      =   0  '平面
   BackColor       =   &H00000000&
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   ClipBehavior    =   0  '無
   ClipControls    =   0   'False
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   4545
   ScaleWidth      =   8730
   Windowless      =   -1  'True
   Begin UnlightVBE.uc角色卡片介面_後_主動技能 cardback_active 
      Height          =   3615
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6376
   End
   Begin UnlightVBE.uc角色卡片介面_後_被動技能 cardback_passive 
      Height          =   3615
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6376
   End
   Begin UnlightVBE.uc角色卡片介面_前 card 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6588
   End
End
Attribute VB_Name = "uc角色卡片介面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_cardbackcheck As Integer, m_ShowOnMode As Boolean
Private m_MusicPlayerObj As ucMusicPlayer
Attribute m_MusicPlayerObj.VB_VarUserMemId = 1073938434

Public Sub 更改異常狀態資料(ByVal buffnum As Integer, ByVal ImagePath As String, ByVal num As Integer, ByVal tot As Integer, ByVal isVisible As Boolean)
    card.更改異常狀態資料 buffnum, ImagePath, num, tot, isVisible
End Sub
Public Sub 異常狀態全重設()
    Call card.異常狀態全重設
End Sub
Public Sub CardBack全重設()
    Call cardback_active.全重設
    Call cardback_passive.全重設
End Sub
Public Property Get CardMain_角色圖片() As String
    CardMain_角色圖片 = card.角色圖片
End Property
Public Property Let CardMain_角色圖片(ByVal New_CardMain_角色圖片 As String)
    card.角色圖片 = New_CardMain_角色圖片
    PropertyChanged "CardMain_角色圖片"
End Property
Public Property Get CardMain_角色HP() As Integer
    CardMain_角色HP = card.角色HP
End Property
Public Property Let CardMain_角色HP(ByVal New_CardMain_角色HP As Integer)
    card.角色HP = New_CardMain_角色HP
    PropertyChanged "CardMain_角色HP"
End Property
Public Property Get CardMain_角色HPMAX() As Integer
    CardMain_角色HPMAX = card.角色HPMAX
End Property
Public Property Let CardMain_角色HPMAX(ByVal New_CardMain_角色HPMAX As Integer)
    card.角色HPMAX = New_CardMain_角色HPMAX
    PropertyChanged "CardMain_角色HPMAX"
End Property
Public Property Get CardMain_角色ATK() As Integer
    CardMain_角色ATK = card.角色ATK
End Property
Public Property Let CardMain_角色ATK(ByVal New_CardMain_角色ATK As Integer)
    card.角色ATK = New_CardMain_角色ATK
    PropertyChanged "CardMain_角色ATK"
End Property
Public Property Get CardMain_角色DEF() As Integer
    CardMain_角色DEF = card.角色DEF
End Property
Public Property Let CardMain_角色DEF(ByVal New_CardMain_角色DEF As Integer)
    card.角色DEF = New_CardMain_角色DEF
    PropertyChanged "CardMain_角色DEF"
End Property
Public Property Get CardMain_是否為新樣式資訊() As Boolean
    CardMain_是否為新樣式資訊 = card.是否為新樣式資訊
End Property
Public Property Let CardMain_是否為新樣式資訊(ByVal New_CardMain_是否為新樣式資訊 As Boolean)
    card.是否為新樣式資訊 = New_CardMain_是否為新樣式資訊
    PropertyChanged "CardMain_是否為新樣式資訊"
End Property
Public Property Get MusicPlayerObj() As ucMusicPlayer
    Set MusicPlayerObj = m_MusicPlayerObj
End Property

Public Property Let MusicPlayerObj(ByVal vNewValue As ucMusicPlayer)
    Set m_MusicPlayerObj = vNewValue
    PropertyChanged "MusicPlayerObj"
End Property
Public Property Get ShowOnMode() As Boolean
    ShowOnMode = m_ShowOnMode
End Property

Public Property Let ShowOnMode(ByVal vNewValue As Boolean)
    m_ShowOnMode = vNewValue
    PropertyChanged "ShowOnMode"
    Call ShowOnModeChange
End Property
Public Sub CardBack_主動技_技能名稱(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.主動技_技能名稱 num, skillstr
End Sub
Public Sub CardBack_主動技_技能說明(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.主動技_技能說明 num, skillstr
End Sub
Public Sub CardBack_主動技_階段代碼(ByVal num As Integer, ByVal newTurnNum As Integer)
    cardback_active.主動技_階段代碼 num, newTurnNum
End Sub
Public Sub CardBack_主動技_距離代碼(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.主動技_距離代碼 num, skillstr
End Sub
Public Sub CardBack_主動技_卡片代碼(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.主動技_卡片代碼 num, skillstr
End Sub
Public Sub CardBack_被動技_技能名稱(ByVal num As Integer, ByVal skillstr As String)
    cardback_passive.被動技_技能名稱 num, skillstr
End Sub
Public Sub CardBack_被動技_技能說明(ByVal num As Integer, ByVal skillstr As String)
    cardback_passive.被動技_技能說明 num, skillstr
End Sub

Private Sub card_CardClick()
    If m_cardbackcheck <= 1 Then
        cardback_active.Visible = False
        cardback_active.Left = 0
        cardback_active.Top = 0
        cardback_active.Visible = True
        cardback_active.ZOrder
        m_cardbackcheck = 1
    Else
        cardback_passive.Visible = False
        cardback_passive.Left = 0
        cardback_passive.Top = 0
        cardback_passive.Visible = True
        cardback_passive.ZOrder
        m_cardbackcheck = 2
    End If
    m_MusicPlayerObj.MusicStop
    m_MusicPlayerObj.MusicPlay
    card.Visible = False
End Sub

Private Sub cardback_active_ClickBack()
    card.Visible = False
    card.Left = 0
    card.Top = 0
    card.Visible = True
    card.ZOrder
    cardback_active.Visible = False
    m_cardbackcheck = 1
    m_MusicPlayerObj.MusicStop
    m_MusicPlayerObj.MusicPlay
End Sub

Private Sub cardback_active_ClickPassive()
    cardback_passive.Visible = False
    cardback_passive.Left = 0
    cardback_passive.Top = 0
    cardback_passive.Visible = True
    cardback_passive.ZOrder
    cardback_active.Visible = False
    m_cardbackcheck = 2
End Sub

Private Sub cardback_passive_ClickActive()
    cardback_active.Visible = False
    cardback_active.Left = 0
    cardback_active.Top = 0
    cardback_active.Visible = True
    cardback_active.ZOrder
    cardback_passive.Visible = False
    m_cardbackcheck = 1
End Sub

Private Sub cardback_passive_ClickBack()
    card.Visible = False
    card.Left = 0
    card.Top = 0
    card.Visible = True
    card.ZOrder
    cardback_passive.Visible = False
    m_cardbackcheck = 2
    m_MusicPlayerObj.MusicStop
    m_MusicPlayerObj.MusicPlay
End Sub
Private Sub ShowOnModeChange()
    card.ShowOnMode = m_ShowOnMode
    cardback_active.ShowOnMode = m_ShowOnMode
    cardback_passive.ShowOnMode = m_ShowOnMode
End Sub
