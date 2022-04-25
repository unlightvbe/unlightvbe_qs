VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc戰鬥系統牌型介面 
   Appearance      =   0  '平面
   BackColor       =   &H00808080&
   BackStyle       =   0  '透明
   ClientHeight    =   9915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   9915
   ScaleWidth      =   11340
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage ImageMouseMoveActivecom 
      Height          =   330
      Index           =   4
      Left            =   0
      Top             =   600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":0000
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActivecom 
      Height          =   330
      Index           =   3
      Left            =   2280
      Top             =   600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":026F
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActivecom 
      Height          =   330
      Index           =   2
      Left            =   4560
      Top             =   600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":04DE
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActivecom 
      Height          =   330
      Index           =   1
      Left            =   6840
      Top             =   600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":074D
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActiveus 
      Height          =   330
      Index           =   4
      Left            =   9240
      Top             =   6240
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":09BC
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActiveus 
      Height          =   330
      Index           =   3
      Left            =   6960
      Top             =   6240
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":0C2B
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActiveus 
      Height          =   330
      Index           =   2
      Left            =   4800
      Top             =   6240
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":0E9A
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage ImageMouseMoveActiveus 
      Height          =   330
      Index           =   1
      Left            =   2640
      Top             =   6240
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      Image           =   "uc戰鬥系統牌型介面.ctx":1109
      Scaler          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage bnok 
      Height          =   1050
      Left            =   7320
      Top             =   8160
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1852
      Image           =   "uc戰鬥系統牌型介面.ctx":1378
      Props           =   5
   End
   Begin VB.Label activecom 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label activecom 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   17
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label activecom 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   16
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label activecom 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   15
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label activeus 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   4
      Left            =   9120
      TabIndex        =   14
      Top             =   6240
      Width           =   2205
   End
   Begin VB.Label activeus 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   13
      Top             =   6240
      Width           =   2205
   End
   Begin VB.Label activeus 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Top             =   6240
      Width           =   2205
   End
   Begin VB.Label activeus 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      Caption         =   "人物技能"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   6240
      Width           =   2205
   End
   Begin VB.Image cardpagejpg 
      Height          =   465
      Left            =   120
      Picture         =   "uc戰鬥系統牌型介面.ctx":3899
      Top             =   960
      Width           =   570
   End
   Begin VB.Label passivetext_com 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Regular"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   8760
      TabIndex        =   10
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label passivetext_com 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Regular"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   8760
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label passivetext_com 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Regular"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label passivetext_com 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Regular"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin ImageX.aicAlphaImage passivelight_com 
      Height          =   255
      Index           =   4
      Left            =   9600
      Top             =   2280
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":3D5A
      Opacity         =   70
      Props           =   5
   End
   Begin ImageX.aicAlphaImage passivelight_com 
      Height          =   255
      Index           =   3
      Left            =   9600
      Top             =   2040
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":4698
      Opacity         =   70
      Props           =   5
   End
   Begin ImageX.aicAlphaImage passivelight_com 
      Height          =   255
      Index           =   2
      Left            =   9600
      Top             =   1800
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":4FD6
      Opacity         =   70
      Props           =   5
   End
   Begin ImageX.aicAlphaImage passivelight_com 
      Height          =   255
      Index           =   1
      Left            =   9600
      Top             =   1560
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":5914
      Opacity         =   70
      Props           =   5
   End
   Begin VB.Label passivetext_us 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese DemiLight"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label passivetext_us 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese DemiLight"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label passivetext_us 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese DemiLight"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label passivetext_us 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "精密射擊"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese DemiLight"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin ImageX.aicAlphaImage passivelight_us 
      Height          =   255
      Index           =   4
      Left            =   0
      Top             =   2280
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":6252
      Opacity         =   70
      Mirror          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage passivelight_us 
      Height          =   255
      Index           =   3
      Left            =   0
      Top             =   2040
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":6BBC
      Opacity         =   70
      Mirror          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage passivelight_us 
      Height          =   255
      Index           =   2
      Left            =   0
      Top             =   1800
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":7526
      Opacity         =   70
      Mirror          =   1
      Props           =   5
   End
   Begin ImageX.aicAlphaImage passivelight_us 
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   1560
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   450
      Image           =   "uc戰鬥系統牌型介面.ctx":7E90
      Opacity         =   70
      Mirror          =   1
      Props           =   5
   End
   Begin VB.Image stagejpgn 
      Height          =   270
      Left            =   9120
      Picture         =   "uc戰鬥系統牌型介面.ctx":87FA
      Top             =   1080
      Width           =   2280
   End
   Begin VB.Label pageul 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "57"
      BeginProperty Font 
         Name            =   "Bradley Gratis"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   560
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin UnlightVBE.uc訊息視窗 messagetext 
      Height          =   1200
      Left            =   2640
      TabIndex        =   1
      Top             =   8100
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2196
   End
   Begin VB.Image cardbackjpg 
      Height          =   1455
      Left            =   2535
      Picture         =   "uc戰鬥系統牌型介面.ctx":8CD7
      Top             =   6600
      Width           =   8910
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  '實心
      Height          =   3615
      Left            =   2535
      Top             =   6240
      Width           =   9135
   End
   Begin VB.Label turnnum 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Bradley Gratis"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   10200
      TabIndex        =   0
      Top             =   480
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Image turnpe 
      Height          =   420
      Left            =   10200
      Picture         =   "uc戰鬥系統牌型介面.ctx":33111
      Top             =   480
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  '實心
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11415
   End
   Begin ImageX.aicAlphaImage cardunderjpg 
      Height          =   360
      Left            =   0
      Top             =   1020
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      Image           =   "uc戰鬥系統牌型介面.ctx":335BB
      Props           =   5
   End
End
Attribute VB_Name = "uc戰鬥系統牌型介面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Turn As Integer, m_cardnum As Integer, m_passivevisble As Boolean, m_bnoktype As Integer

Public Event ActiveMouseMove(ByVal uscom As Integer, ByVal num As Integer)
Public Event InterfaceMouseMove()
Public Event BnOKMouseMove()
Public Event BnOKClick()
Public Event ActiveMouseEnter(ByVal uscom As Integer, ByVal num As Integer)
Public Event ActiveMouseExit(ByVal uscom As Integer, ByVal num As Integer)
Public Property Get turn() As Integer
   turn = m_Turn
End Property
Public Property Let turn(ByVal New_Turn As Integer)
   m_Turn = New_Turn
   PropertyChanged "Turn"
   '=================
   turnnum.Caption = Me.turn
   If turnnum.FontName <> "Bradley Gratis" Then
        turnnum.FontSize = 20
   Else
        turnnum.FontSize = 24
   End If
End Property
Public Property Get Cardnum() As Integer
   Cardnum = m_cardnum
End Property
Public Property Let Cardnum(ByVal New_Cardnum As Integer)
   m_cardnum = New_Cardnum
   PropertyChanged "Cardnum"
   '=================
   pageul.Caption = Me.Cardnum
   If pageul.FontName <> "Bradley Gratis" Then
        pageul.FontSize = 16
   Else
        pageul.FontSize = 20
   End If
End Property
Public Property Get Passive_介面顯示() As Boolean
   Passive_介面顯示 = m_passivevisble
End Property
Public Property Let Passive_介面顯示(ByVal New_Passive_介面顯示 As Boolean)
   m_passivevisble = New_Passive_介面顯示
   PropertyChanged "Passive_介面顯示"
   '=================
   Dim i As Integer
   If Me.Passive_介面顯示 = False Then
       cardunderjpg.Visible = False
       cardpagejpg.Visible = False
       pageul.Visible = False
       stagejpgn.Visible = False
       For i = 1 To 4
          Me.Passive_使用者_技能隱藏 i
          Me.Passive_電腦_技能隱藏 i
       Next
   Else
       cardunderjpg.Visible = True
       cardpagejpg.Visible = True
       pageul.Visible = True
       pageul.ZOrder
       stagejpgn.Visible = True
   End If
End Property

Sub MessageClear()
messagetext.MessageTextClear
End Sub
Sub Passive_技能一方全重設(ByVal uscom As Integer)
Dim i As Integer
Select Case uscom
    Case 1
        For i = 1 To 4
            passivelight_us(i).Visible = False
            Me.Passive_使用者_技能燈變暗 i
            passivetext_us(i).Visible = False
            passivetext_us(i).Caption = ""
        Next
    Case 2
        For i = 1 To 4
            passivelight_com(i).Visible = False
            Me.Passive_電腦_技能燈變暗 i
            passivetext_com(i).Visible = False
            passivetext_com(i).Caption = ""
        Next
End Select
End Sub
Sub Passive_電腦_技能名稱(ByVal name As String, ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivetext_com(num).Caption = name
End Sub
Sub Passive_電腦_技能燈變暗(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_com(num).ClearImage
    passivelight_com(num).LoadImage_FromFile App.Path & "\gif\system\passivelightoff.png"
    passivelight_com(num).Mirror = aiMirrorNone
End Sub
Sub Passive_電腦_技能燈發亮(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_com(num).ClearImage
    passivelight_com(num).LoadImage_FromFile App.Path & "\gif\system\passivelighton.png"
    passivelight_com(num).Mirror = aiMirrorNone
End Sub
Sub Passive_電腦_技能隱藏(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_com(num).Visible = False
    passivetext_com(num).Visible = False
End Sub
Sub Passive_電腦_技能顯示(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_com(num).Visible = True
    passivetext_com(num).Visible = True
End Sub
Sub Passive_使用者_技能隱藏(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_us(num).Visible = False
    passivetext_us(num).Visible = False
End Sub
Sub Passive_使用者_技能顯示(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_us(num).Visible = True
    passivetext_us(num).Visible = True
End Sub
Sub Passive_使用者_技能燈變暗(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_us(num).ClearImage
    passivelight_us(num).LoadImage_FromFile App.Path & "\gif\system\passivelightoff.png"
    passivelight_us(num).Mirror = aiMirrorHorizontal
End Sub
Sub Passive_使用者_技能燈發亮(ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivelight_us(num).ClearImage
    passivelight_us(num).LoadImage_FromFile App.Path & "\gif\system\passivelighton.png"
    passivelight_us(num).Mirror = aiMirrorHorizontal
End Sub
Sub Passive_使用者_技能名稱(ByVal name As String, ByVal num As Integer)
    If num < 1 Or num > 4 Then Exit Sub
    passivetext_us(num).Caption = name
End Sub
Sub stagejpg(ByVal filestr As String)
    If filestr <> "" Then
       stagejpgn.Picture = LoadPicture(filestr)
   End If
End Sub
Sub Message(ByVal msgstr As String)
    messagetext.MeaageText msgstr
End Sub

Private Sub activecom_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent ActiveMouseMove(2, Index)
End Sub

Private Sub activeus_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent ActiveMouseMove(1, Index)
End Sub

Private Sub bnok_Click(ByVal Button As Integer)
RaiseEvent BnOKClick
End Sub

Private Sub bnok_MouseEnter()
bnok.LoadImage_FromFile app_path & "gif\system\ok_2.jpg"
m_bnoktype = 2
End Sub

Private Sub bnok_MouseExit()
bnok.LoadImage_FromFile app_path & "gif\system\ok_1.jpg"
m_bnoktype = 1
End Sub

Private Sub bnok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent BnOKMouseMove
End Sub

Private Sub cardunderjpg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent InterfaceMouseMove
End Sub

Private Sub ImageMouseMoveActivecom_MouseEnter(Index As Integer)
RaiseEvent ActiveMouseEnter(2, Index)
End Sub

Private Sub ImageMouseMoveActivecom_MouseExit(Index As Integer)
RaiseEvent ActiveMouseExit(2, Index)
End Sub

Private Sub ImageMouseMoveActiveus_MouseEnter(Index As Integer)
RaiseEvent ActiveMouseEnter(1, Index)
End Sub

Private Sub ImageMouseMoveActiveus_MouseExit(Index As Integer)
RaiseEvent ActiveMouseExit(1, Index)
End Sub

Private Sub stagejpgn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent InterfaceMouseMove
End Sub
Sub ActiveDescription(ByVal uscom As Integer, ByVal num As Integer, ByVal skillobj As clsPersonActiveSkill)
Me.ActiveSkillName uscom, num, skillobj.name
Me.ActiveSkillNameFontSize uscom, num, skillobj.NameFontSize
End Sub

Sub ActiveSkillLight(ByVal uscom As Integer, ByVal num As Integer, ByVal isOn As Boolean)
Select Case uscom
 Case 1
    If isOn = True Then
       activeus(num).ForeColor = RGB(255, 255, 0)
       activeus(num).BackColor = RGB(47, 94, 94)
    Else
       activeus(num).ForeColor = RGB(192, 192, 192)
       activeus(num).BackColor = RGB(0, 0, 0)
    End If
 Case 2
    If isOn = True Then
       activecom(num).ForeColor = RGB(255, 255, 0)
       activecom(num).BackColor = RGB(47, 94, 94)
    Else
       activecom(num).ForeColor = RGB(192, 192, 192)
       activecom(num).BackColor = RGB(0, 0, 0)
    End If
End Select
    
End Sub
Sub ActiveSkillVisable(ByVal uscom As Integer, ByVal num As Integer, ByVal isOn As Boolean)
Select Case uscom
 Case 1
    activeus(num).Visible = isOn
 Case 2
    activecom(num).Visible = isOn
End Select
End Sub
Sub ActiveSkillName(ByVal uscom As Integer, ByVal num As Integer, ByVal namestr As String)
Select Case uscom
 Case 1
    activeus(num).Caption = namestr
 Case 2
    activecom(num).Caption = namestr
End Select
End Sub
Sub ActiveSkillNameFontSize(ByVal uscom As Integer, ByVal num As Integer, ByVal sizenum As Integer)
If sizenum <= 7 Then Exit Sub
Select Case uscom
 Case 1
    activeus(num).FontSize = sizenum
 Case 2
    activecom(num).FontSize = sizenum
End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent InterfaceMouseMove
End Sub
Sub BnOKStartListen()
bnok.LoadImage_FromFile app_path & "gif\system\ok_1.jpg"
m_bnoktype = 1
bnok.Enabled = True
bnok.Visible = True
End Sub
Sub BnOKStopListen()
bnok.LoadImage_FromFile app_path & "gif\system\ok_3.jpg"
m_bnoktype = 3
bnok.Enabled = False
End Sub
Sub BnOKVisable(ByVal isOn As Boolean)
bnok.Visible = isOn
End Sub
Sub BnOKEnabled(ByVal isOn As Boolean)
bnok.Enabled = isOn
End Sub
