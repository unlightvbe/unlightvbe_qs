VERSION 5.00
Begin VB.UserControl uc����d������ 
   Appearance      =   0  '����
   BackColor       =   &H00000000&
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   ClipBehavior    =   0  '�L
   ClipControls    =   0   'False
   HitBehavior     =   2  '�ϥ�ø�ϰϰ�
   ScaleHeight     =   4545
   ScaleWidth      =   8730
   Windowless      =   -1  'True
   Begin UnlightVBE.uc����d������_��_�D�ʧޯ� cardback_active 
      Height          =   3615
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6376
   End
   Begin UnlightVBE.uc����d������_��_�Q�ʧޯ� cardback_passive 
      Height          =   3615
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6376
   End
   Begin UnlightVBE.uc����d������_�e card 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6588
   End
End
Attribute VB_Name = "uc����d������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_cardbackcheck As Integer, m_ShowOnMode As Boolean
Private m_MusicPlayerObj As ucMusicPlayer
Attribute m_MusicPlayerObj.VB_VarUserMemId = 1073938434

Public Sub ��ﲧ�`���A���(ByVal buffnum As Integer, ByVal ImagePath As String, ByVal num As Integer, ByVal tot As Integer, ByVal isVisible As Boolean)
    card.��ﲧ�`���A��� buffnum, ImagePath, num, tot, isVisible
End Sub
Public Sub ���`���A�����]()
    Call card.���`���A�����]
End Sub
Public Sub CardBack�����]()
    Call cardback_active.�����]
    Call cardback_passive.�����]
End Sub
Public Property Get CardMain_����Ϥ�() As String
    CardMain_����Ϥ� = card.����Ϥ�
End Property
Public Property Let CardMain_����Ϥ�(ByVal New_CardMain_����Ϥ� As String)
    card.����Ϥ� = New_CardMain_����Ϥ�
    PropertyChanged "CardMain_����Ϥ�"
End Property
Public Property Get CardMain_����HP() As Integer
    CardMain_����HP = card.����HP
End Property
Public Property Let CardMain_����HP(ByVal New_CardMain_����HP As Integer)
    card.����HP = New_CardMain_����HP
    PropertyChanged "CardMain_����HP"
End Property
Public Property Get CardMain_����HPMAX() As Integer
    CardMain_����HPMAX = card.����HPMAX
End Property
Public Property Let CardMain_����HPMAX(ByVal New_CardMain_����HPMAX As Integer)
    card.����HPMAX = New_CardMain_����HPMAX
    PropertyChanged "CardMain_����HPMAX"
End Property
Public Property Get CardMain_����ATK() As Integer
    CardMain_����ATK = card.����ATK
End Property
Public Property Let CardMain_����ATK(ByVal New_CardMain_����ATK As Integer)
    card.����ATK = New_CardMain_����ATK
    PropertyChanged "CardMain_����ATK"
End Property
Public Property Get CardMain_����DEF() As Integer
    CardMain_����DEF = card.����DEF
End Property
Public Property Let CardMain_����DEF(ByVal New_CardMain_����DEF As Integer)
    card.����DEF = New_CardMain_����DEF
    PropertyChanged "CardMain_����DEF"
End Property
Public Property Get CardMain_�O�_���s�˦���T() As Boolean
    CardMain_�O�_���s�˦���T = card.�O�_���s�˦���T
End Property
Public Property Let CardMain_�O�_���s�˦���T(ByVal New_CardMain_�O�_���s�˦���T As Boolean)
    card.�O�_���s�˦���T = New_CardMain_�O�_���s�˦���T
    PropertyChanged "CardMain_�O�_���s�˦���T"
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
Public Sub CardBack_�D�ʧ�_�ޯ�W��(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.�D�ʧ�_�ޯ�W�� num, skillstr
End Sub
Public Sub CardBack_�D�ʧ�_�ޯ໡��(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.�D�ʧ�_�ޯ໡�� num, skillstr
End Sub
Public Sub CardBack_�D�ʧ�_���q�N�X(ByVal num As Integer, ByVal newTurnNum As Integer)
    cardback_active.�D�ʧ�_���q�N�X num, newTurnNum
End Sub
Public Sub CardBack_�D�ʧ�_�Z���N�X(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.�D�ʧ�_�Z���N�X num, skillstr
End Sub
Public Sub CardBack_�D�ʧ�_�d���N�X(ByVal num As Integer, ByVal skillstr As String)
    cardback_active.�D�ʧ�_�d���N�X num, skillstr
End Sub
Public Sub CardBack_�Q�ʧ�_�ޯ�W��(ByVal num As Integer, ByVal skillstr As String)
    cardback_passive.�Q�ʧ�_�ޯ�W�� num, skillstr
End Sub
Public Sub CardBack_�Q�ʧ�_�ޯ໡��(ByVal num As Integer, ByVal skillstr As String)
    cardback_passive.�Q�ʧ�_�ޯ໡�� num, skillstr
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
