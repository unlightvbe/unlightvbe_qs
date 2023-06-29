VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc卡片背面_被動技能 
   Appearance      =   0  '平面
   BackColor       =   &H00404040&
   BackStyle       =   0  '透明
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   1545
   ScaleWidth      =   3810
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage aicAlphaImageBar 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   873
      Image           =   "uc卡片背面_被動技能.ctx":0000
      Opacity         =   0
   End
   Begin VB.Label personcardback_passivetext 
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin ImageX.aicAlphaImage cardbackpassiveBR 
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   688
      Image           =   "uc卡片背面_被動技能.ctx":04BB
      Opacity         =   50
      Props           =   5
   End
End
Attribute VB_Name = "uc卡片背面_被動技能"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_SkillName As String
Private m_SkillDescription As String
Private m_ShowOnMode As Boolean
Public Event ClickBR()

Public Property Get SkillName() As String
    SkillName = m_SkillName
End Property

Public Property Let SkillName(ByVal vNewValue As String)
    m_SkillName = vNewValue
    PropertyChanged "SkillName"
    '=====================
    If m_SkillName <> "" And m_ShowOnMode = True Then
        personcardback_passivetext.Caption = m_SkillName
        personcardback_passivetext.Visible = True
    Else
        personcardback_passivetext.Visible = False
    End If
End Property
Public Property Get SkillDescription() As String
    SkillDescription = m_SkillDescription
End Property

Public Property Let SkillDescription(ByVal vNewValue As String)
    m_SkillDescription = vNewValue
    PropertyChanged "SkillDescription"
End Property

Public Property Get ShowOnMode() As Boolean
    ShowOnMode = m_ShowOnMode
End Property

Public Property Let ShowOnMode(ByVal vNewValue As Boolean)
    m_ShowOnMode = vNewValue
    PropertyChanged "ShowOnMode"
    Call ShowOnModeChange
End Property

Private Sub ShowOnModeChange()
    If m_ShowOnMode = True Then
        Me.SkillName = m_SkillName
    Else
        personcardback_passivetext.Visible = False
    End If
End Sub
Public Sub ResetAll()
    Me.SkillName = ""
    Me.SkillDescription = ""
    Me.ShowOnMode = False
End Sub

Private Sub aicAlphaImageBar_Click(ByVal Button As Integer)
    RaiseEvent ClickBR
End Sub

Private Sub aicAlphaImageBar_MouseEnter()
    cardbackpassiveBR.Opacity = 50
End Sub

Private Sub aicAlphaImageBar_MouseExit()
    cardbackpassiveBR.Opacity = 0
End Sub

Private Sub UserControl_Show()
    cardbackpassiveBR.Opacity = 0
End Sub
