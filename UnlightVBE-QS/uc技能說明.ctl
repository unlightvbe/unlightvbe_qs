VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc技能說明 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   3930
   ScaleWidth      =   2805
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage ImageMouseMove 
      Height          =   3570
      Left            =   0
      Top             =   0
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   6297
      Image           =   "uc技能說明.ctx":0000
      Props           =   5
   End
   Begin VB.Label atkinghelpi3 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "距離："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label atkinghelpt3 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "む卡片めめめめめめめめめ"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label atkinghelpt2 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "む距離め"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label atkinghelpi4 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "卡片："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label atkinghelpt1 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "む階段め"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label atkinghelpi2 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "階段："
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label atkinghelpt4 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "む這裡是技能效果區めめめめめめめめめめめめめめめめめめ"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label atkinghelpi5 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "「效果」"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label atkinghelpi1 
      BackColor       =   &H00000000&
      BackStyle       =   0  '透明
      Caption         =   "「條件」"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin ImageX.aicAlphaImage aicAlphaImage1 
      Height          =   7200
      Left            =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   12700
      Image           =   "uc技能說明.ctx":026F
      Opacity         =   75
      Props           =   5
   End
End
Attribute VB_Name = "uc技能說明"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_stage As String, m_distance As String, m_card As String, m_effect As String

Public Event MouseEnter()
Public Property Get Stage() As String
    Stage = m_stage
End Property

Public Property Let Stage(ByVal vNewValue As String)
    m_stage = vNewValue
    PropertyChanged "Stage"
    '==================
    atkinghelpt1.Caption = m_stage
End Property

Public Property Get Distance() As String
    Distance = m_distance
End Property

Public Property Let Distance(ByVal vNewValue As String)
    m_distance = vNewValue
    PropertyChanged "Distance"
    '==================
    atkinghelpt2.Caption = m_distance
End Property

Public Property Get card() As String
    card = m_card
End Property

Public Property Let card(ByVal vNewValue As String)
    m_card = vNewValue
    PropertyChanged "Card"
    '==================
    atkinghelpt3.Caption = m_card
End Property

Public Property Get Effect() As Variant
    Effect = m_effect
End Property

Public Property Let Effect(ByVal vNewValue As Variant)
    m_effect = vNewValue
    PropertyChanged "Effect"
    '==================
    atkinghelpt4.Caption = m_effect
End Property

Private Sub ImageMouseMove_MouseEnter()
RaiseEvent MouseEnter
End Sub
