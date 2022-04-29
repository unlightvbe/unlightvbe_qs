VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc卡片背面_主動技能 
   Appearance      =   0  '平面
   BackColor       =   &H00404040&
   BackStyle       =   0  '透明
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   2310
   ScaleWidth      =   3945
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage aicAlphaImageBar 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   1085
      Image           =   "uc卡片背面_主動技能.ctx":0000
      Opacity         =   0
   End
   Begin UnlightVBE.uc卡片背面 personcardback_turn 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   238
   End
   Begin UnlightVBE.uc卡片背面 personcardback_num 
      Height          =   255
      Index           =   1
      Left            =   930
      TabIndex        =   2
      Top             =   280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_range 
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   280
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_range 
      Height          =   255
      Index           =   2
      Left            =   630
      TabIndex        =   4
      Top             =   280
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_range 
      Height          =   255
      Index           =   3
      Left            =   780
      TabIndex        =   5
      Top             =   280
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_num 
      Height          =   255
      Index           =   2
      Left            =   1230
      TabIndex        =   6
      Top             =   280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_num 
      Height          =   255
      Index           =   3
      Left            =   1530
      TabIndex        =   7
      Top             =   280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_num 
      Height          =   255
      Index           =   4
      Left            =   1830
      TabIndex        =   8
      Top             =   280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
   End
   Begin UnlightVBE.uc卡片背面 personcardback_num 
      Height          =   255
      Index           =   5
      Left            =   2130
      TabIndex        =   9
      Top             =   280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
   End
   Begin VB.Label personcardback_text 
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese DemiLight"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin ImageX.aicAlphaImage cardbackBR 
      Height          =   435
      Left            =   20
      Top             =   0
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   767
      Image           =   "uc卡片背面_主動技能.ctx":04BB
      Props           =   13
   End
End
Attribute VB_Name = "uc卡片背面_主動技能"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_SkillName As String
Private m_TurnNum As Integer
Private m_RangeStr As String
Private m_CardStr As String
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
        personcardback_text.Caption = m_SkillName
        personcardback_text.Visible = True
    Else
        personcardback_text.Visible = False
    End If
End Property
Public Sub ResetAll()
    Me.SkillName = ""
    Me.turnnum = 0
    Me.RangeStr = ""
    Me.CardStr = ""
    Me.SkillDescription = ""
    Me.ShowOnMode = False
End Sub

Public Property Get turnnum() As Integer
    turnnum = m_TurnNum
End Property

Public Property Let turnnum(ByVal vNewValue As Integer)
    If vNewValue < 1 Or vNewValue > 3 Then
        m_TurnNum = 0
    Else
        m_TurnNum = vNewValue
    End If
    PropertyChanged "turnnum"
    '=====================
    If m_TurnNum <> 0 And m_ShowOnMode = True Then
        personcardback_turn.物件類別 = 3
        personcardback_turn.圖片 = App.Path & "\gif\system\cardback\CBturn.png"
        personcardback_turn.項目編號 = m_TurnNum
        personcardback_turn.Visible = True
    Else
        personcardback_turn.Visible = False
    End If
End Property

Public Property Get RangeStr() As String
    RangeStr = m_RangeStr
End Property

Public Property Let RangeStr(ByVal vNewValue As String)
    m_RangeStr = vNewValue
    PropertyChanged "RangeStr"
    '=====================
    If m_RangeStr <> "" And m_ShowOnMode = True Then
        Dim k As Integer
        For k = 1 To 3
             personcardback_range(k).物件類別 = 2
             personcardback_range(k).圖片 = App.Path & "\gif\system\cardback\CBrge.png"
             If Mid(m_RangeStr, k, 1) = 1 Then
                 If k < 3 Then
                     personcardback_range(k).項目編號 = 1
                 Else
                     personcardback_range(k).項目編號 = 3
                 End If
             Else
                 personcardback_range(k).項目編號 = 2
             End If
        Next
    Else
        For k = 1 To 3
            personcardback_range(k).物件類別 = 2
            personcardback_range(k).圖片 = App.Path & "\gif\system\cardback\CBrge.png"
            personcardback_range(k).項目編號 = 2
        Next
    End If
End Property

Public Property Get CardStr() As String
    CardStr = m_CardStr
End Property

Public Property Let CardStr(ByVal vNewValue As String)
    m_CardStr = vNewValue
    PropertyChanged "CardStr"
    '=====================
    If m_CardStr <> "" And m_ShowOnMode = True Then
        Dim strw As Variant, k As Integer, n As Integer
        strw = Split(m_CardStr, "&")
        For k = 0 To UBound(strw)
            If Len(strw(k)) = 3 Then
                   personcardback_num(k + 1).物件類別 = 1
                   personcardback_num(k + 1).圖片 = App.Path & "\gif\system\cardback\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                   If Mid(strw(k), 2, 1) = "a" Then
                        n = 10
                   ElseIf Mid(strw(k), 2, 1) = "b" Then
                        n = 11
                   Else
                        n = Val(Mid(strw(k), 2, 1))
                   End If
                   personcardback_num(k + 1).項目編號 = n
                   personcardback_num(k + 1).Visible = True
            Else
                   personcardback_num(k + 1).Visible = False
            End If
        Next
        For k = UBound(strw) + 1 To 4
            personcardback_num(k + 1).Visible = False
        Next
    Else
        For k = 0 To 4
            personcardback_num(k + 1).Visible = False
        Next
    End If
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
    Me.turnnum = m_TurnNum
    Me.RangeStr = m_RangeStr
    Me.CardStr = m_CardStr
Else
    Dim k As Integer
    personcardback_turn.Visible = False
    personcardback_text.Visible = False
    For k = 1 To 5
        personcardback_num(k).Visible = False
    Next
    '================
    For k = 1 To 3
        personcardback_range(k).物件類別 = 2
        personcardback_range(k).圖片 = App.Path & "\gif\system\cardback\CBrge.png"
        personcardback_range(k).項目編號 = 2
    Next
    '================
End If
End Sub

Private Sub aicAlphaImageBar_Click(ByVal Button As Integer)
    RaiseEvent ClickBR
End Sub

Private Sub aicAlphaImageBar_MouseEnter()
    cardbackBR.Opacity = 100
End Sub

Private Sub aicAlphaImageBar_MouseExit()
    cardbackBR.Opacity = 0
End Sub

Private Sub UserControl_Show()
    cardbackBR.Opacity = 0
End Sub

Public Property Get SkillDescription() As String
    SkillDescription = m_SkillDescription
End Property

Public Property Let SkillDescription(ByVal vNewValue As String)
    m_SkillDescription = vNewValue
    PropertyChanged "SkillDescription"
End Property
