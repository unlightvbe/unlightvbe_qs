VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc����d������_��_�D�ʧޯ� 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ClipBehavior    =   0  '�L
   HitBehavior     =   2  '�ϥ�ø�ϰϰ�
   ScaleHeight     =   4380
   ScaleWidth      =   6015
   Windowless      =   -1  'True
   Begin VB.Image cardpassiveChickimage 
      Height          =   300
      Left            =   1240
      Top             =   0
      Width           =   1260
   End
   Begin UnlightVBE.uc�d���I��_�D�ʧޯ� cardbackStatus 
      Height          =   495
      Index           =   1
      Left            =   100
      TabIndex        =   0
      Top             =   330
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
   End
   Begin UnlightVBE.uc�d���I��_�D�ʧޯ� cardbackStatus 
      Height          =   495
      Index           =   3
      Left            =   100
      TabIndex        =   1
      Top             =   1280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
   End
   Begin UnlightVBE.uc�d���I��_�D�ʧޯ� cardbackStatus 
      Height          =   495
      Index           =   2
      Left            =   100
      TabIndex        =   2
      Top             =   800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
   End
   Begin UnlightVBE.uc�d���I��_�D�ʧޯ� cardbackStatus 
      Height          =   495
      Index           =   4
      Left            =   100
      TabIndex        =   3
      Top             =   1720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
   End
   Begin VB.Label personcardback_main 
      BackStyle       =   0  '�z��
      Caption         =   "DEF+7�C���m���\�ɡA������P�ҶW�L�����m�P�Ȫ��ˮ`"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese DemiLight"
         Size            =   8.25
         Charset         =   136
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin ImageX.aicAlphaImage aicAlphaImage1 
      Height          =   3600
      Left            =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   6350
      Image           =   "uc����d������_��_�D�ʧޯ�.ctx":0000
      Props           =   5
   End
End
Attribute VB_Name = "uc����d������_��_�D�ʧޯ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_cardback_activecheck As Integer
Private m_ShowOnMode As Boolean
Public Event ClickPassive()
Public Event ClickBack()
Public Sub �����]()
Me.ShowOnMode = False
End Sub
Public Sub �D�ʧ�_�ޯ�W��(ByVal num As Integer, ByVal skillstr As String)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).SkillName = skillstr
    End If
End Sub
Public Sub �D�ʧ�_�ޯ໡��(ByVal num As Integer, ByVal skillstr As String)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).SkillDescription = skillstr
    End If
End Sub
Public Sub �D�ʧ�_���q�N�X(ByVal num As Integer, ByVal newTurnNum As Integer)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).turnnum = newTurnNum
    End If
End Sub
Public Sub �D�ʧ�_�Z���N�X(ByVal num As Integer, ByVal newstr As String)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).RangeStr = newstr
    End If
End Sub
Public Sub �D�ʧ�_�d���N�X(ByVal num As Integer, ByVal newstr As String)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).CardStr = newstr
    End If
End Sub

Private Sub aicAlphaImage1_Click(ByVal Button As Integer)
RaiseEvent ClickBack
End Sub

Private Sub cardpassiveChickimage_Click()
RaiseEvent ClickPassive
End Sub

Private Sub cardbackStatus_ClickBR(Index As Integer)
If m_ShowOnMode = True Then
    m_cardback_activecheck = Index
    personcardback_main.Caption = cardbackStatus(Index).SkillDescription
End If
End Sub

Private Sub personcardback_main_Click()
RaiseEvent ClickBack
End Sub

Public Property Get ShowOnMode() As Boolean
    ShowOnMode = m_ShowOnMode
End Property

Public Property Let ShowOnMode(ByVal vNewValue As Boolean)
    m_ShowOnMode = vNewValue
    PropertyChanged "ShowOnMode"
    Call ShowOnModeChange
End Property

Private Sub ShowOnModeChange()
Dim i As Integer
For i = 1 To 4
    cardbackStatus(i).ShowOnMode = m_ShowOnMode
Next
personcardback_main.Caption = ""
m_cardback_activecheck = 0
End Sub
