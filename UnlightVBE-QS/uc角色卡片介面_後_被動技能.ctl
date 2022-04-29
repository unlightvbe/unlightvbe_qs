VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc角色卡片介面_後_被動技能 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   4230
   ScaleWidth      =   4260
   Windowless      =   -1  'True
   Begin UnlightVBE.uc卡片背面_被動技能 cardbackStatus 
      Height          =   375
      Index           =   1
      Left            =   80
      TabIndex        =   1
      Top             =   380
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc卡片背面_被動技能 cardbackStatus 
      Height          =   375
      Index           =   2
      Left            =   80
      TabIndex        =   2
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc卡片背面_被動技能 cardbackStatus 
      Height          =   375
      Index           =   3
      Left            =   80
      TabIndex        =   3
      Top             =   1280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc卡片背面_被動技能 cardbackStatus 
      Height          =   375
      Index           =   4
      Left            =   80
      TabIndex        =   4
      Top             =   1740
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
   End
   Begin VB.Label personcardback_passivemain 
      BackStyle       =   0  '透明
      Caption         =   "DEF+7。防禦成功時，對手受到與所超過之防禦同值的傷害"
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
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Image cardactiveChickimage 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   1260
   End
   Begin ImageX.aicAlphaImage aicAlphaImage1 
      Height          =   3600
      Left            =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   6350
      Image           =   "uc角色卡片介面_後_被動技能.ctx":0000
      Props           =   5
   End
End
Attribute VB_Name = "uc角色卡片介面_後_被動技能"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_cardback_passivecheck As Integer
Private m_ShowOnMode As Boolean
Public Event ClickActive()
Public Event ClickBack()
Public Sub 全重設()
Me.ShowOnMode = False
End Sub
Public Sub 被動技_技能名稱(ByVal num As Integer, ByVal skillstr As String)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).SkillName = skillstr
    End If
End Sub
Public Sub 被動技_技能說明(ByVal num As Integer, ByVal skillstr As String)
    If num >= 1 And num <= 4 Then
        cardbackStatus(num).SkillDescription = skillstr
    End If
End Sub

Private Sub aicAlphaImage1_Click(ByVal Button As Integer)
RaiseEvent ClickBack
End Sub

Private Sub cardactiveChickimage_Click()
RaiseEvent ClickActive
End Sub

Private Sub cardbackStatus_ClickBR(Index As Integer)
If m_ShowOnMode = True Then
    m_cardback_passivecheck = Index
    personcardback_passivemain.Caption = cardbackStatus(Index).SkillDescription
End If
End Sub

Private Sub personcardback_passivemain_Click()
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
personcardback_passivemain.Caption = ""
m_cardback_passivecheck = 0
End Sub
