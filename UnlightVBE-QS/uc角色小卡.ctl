VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc角色小卡 
   Appearance      =   0  '平面
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  '透明
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   705
   ScaleWidth      =   2760
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage ImageMouseMove 
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   926
      Image           =   "uc角色小卡.ctx":0000
      Scaler          =   1
      Props           =   5
   End
   Begin VB.Label labelallhp 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label labeldef 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Label labelatk 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.Label labellv 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label labelcurrenthp 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Bradley Gratis"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label labelname 
      BackStyle       =   0  '透明
      Caption         =   "艾伯李斯特"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin ImageX.aicAlphaImage aicAlphaImage1 
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   900
      Image           =   "uc角色小卡.ctx":026F
      Props           =   5
   End
End
Attribute VB_Name = "uc角色小卡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PersonName As String, m_CurrentHP As Integer, m_Level As Integer, m_atk As Integer, m_def As Integer, m_allhp As Integer, m_ShowOnMode As Boolean
Attribute m_CurrentHP.VB_VarUserMemId = 1073938432
Attribute m_Level.VB_VarUserMemId = 1073938432
Attribute m_atk.VB_VarUserMemId = 1073938432
Attribute m_def.VB_VarUserMemId = 1073938432
Attribute m_allhp.VB_VarUserMemId = 1073938432
Attribute m_ShowOnMode.VB_VarUserMemId = 1073938432
Public Event MouseMove()
Public Event MouseEnter()
Public Event MouseExit()
Public Property Get personName() As String
    personName = m_PersonName
End Property

Public Property Let personName(ByVal vNewValue As String)
    m_PersonName = vNewValue
    PropertyChanged "PersonName"
    If m_ShowOnMode = True Then
        labelname.Caption = m_PersonName
        labelname.Visible = True
    Else
        labelname.Visible = False
    End If
End Property

Public Property Get CurrentHP() As Integer
    CurrentHP = m_CurrentHP
End Property

Public Property Let CurrentHP(ByVal vNewValue As Integer)
    If vNewValue <= 0 Then
        m_CurrentHP = 0
    Else
        m_CurrentHP = vNewValue
    End If
    PropertyChanged "CurrentHP"
    If labelcurrenthp.FontName <> "Bradley Gratis" Then
        labelcurrenthp.FontSize = 14
    End If
    If m_ShowOnMode = True Then
        labelcurrenthp.Caption = m_CurrentHP
    Else
        labelcurrenthp.Caption = m_CurrentHP - m_allhp
    End If
    labelcurrenthp.Visible = True
    Call HPColorChange
End Property

Public Property Get Level() As Integer
    Level = m_Level
End Property

Public Property Let Level(ByVal vNewValue As Integer)
    m_Level = vNewValue
    PropertyChanged "Level"
    If m_ShowOnMode = True Then
        labellv.Caption = m_Level
    Else
        labellv.Caption = "?"
    End If
End Property

Public Property Get ATK() As Integer
    ATK = m_atk
End Property

Public Property Let ATK(ByVal vNewValue As Integer)
    m_atk = vNewValue
    PropertyChanged "ATK"
    If m_ShowOnMode = True Then
        labelatk.Caption = m_atk
    Else
        labelatk.Caption = "?"
    End If
End Property

Public Property Get DEF() As Integer
    DEF = m_def
End Property

Public Property Let DEF(ByVal vNewValue As Integer)
    m_def = vNewValue
    PropertyChanged "DEF"
    If m_ShowOnMode = True Then
        labeldef.Caption = m_def
    Else
        labeldef.Caption = "?"
    End If
End Property

Public Property Get AllHP() As Integer
    AllHP = m_allhp
End Property

Public Property Let AllHP(ByVal vNewValue As Integer)
    m_allhp = vNewValue
    PropertyChanged "AllHP"
    If m_ShowOnMode = True Then
        labelallhp.Caption = m_allhp
        Call HPColorChange
    Else
        labelallhp.Caption = "?"
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

Private Sub aicAlphaImage1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub ImageMouseMove_MouseEnter()
    RaiseEvent MouseEnter
End Sub

Private Sub ImageMouseMove_MouseExit()
    RaiseEvent MouseExit
End Sub

Private Sub labelallhp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub labelatk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub labeldef_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub labellv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub labelname_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub
Private Sub HPColorChange()
    If m_ShowOnMode = True Then
        If m_CurrentHP = m_allhp Then
            labelcurrenthp.ForeColor = RGB(255, 255, 255)
        End If
        If m_CurrentHP < m_allhp Then
            labelcurrenthp.ForeColor = RGB(255, 255, 128)
        End If
        If m_CurrentHP <= m_allhp / 3 Then
            labelcurrenthp.ForeColor = RGB(255, 0, 0)
        End If
    Else
        If m_CurrentHP = m_allhp Then
            labelcurrenthp.ForeColor = RGB(255, 255, 255)
        Else
            labelcurrenthp.ForeColor = RGB(255, 0, 0)
        End If
    End If
End Sub
Private Sub ShowOnModeChange()
    If m_ShowOnMode = True Then
        labelcurrenthp.Caption = m_CurrentHP
        labelcurrenthp.Visible = True
        labelname.Caption = m_PersonName
        labelname.Visible = True
        labellv.Caption = m_Level
        labelatk.Caption = m_atk
        labeldef.Caption = m_def
        labelallhp.Caption = m_allhp
        Call HPColorChange
    Else
        If m_CurrentHP - m_allhp = 0 Then
            labelcurrenthp.Visible = False
        Else
            labelcurrenthp.Caption = m_CurrentHP - m_allhp
            labelcurrenthp.Visible = True
        End If
        labelname.Visible = False
        labellv.Caption = "?"
        labelatk.Caption = "?"
        labeldef.Caption = "?"
        labelallhp.Caption = "?"
    End If
End Sub
