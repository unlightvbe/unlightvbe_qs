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
   HitBehavior     =   0  '無
   ScaleHeight     =   9915
   ScaleWidth      =   11340
   Windowless      =   -1  'True
   Begin VB.Image cardpagejpg 
      Height          =   465
      Left            =   240
      Picture         =   "uc戰鬥系統牌型介面.ctx":0000
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
      Image           =   "uc戰鬥系統牌型介面.ctx":04C1
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
      Image           =   "uc戰鬥系統牌型介面.ctx":0DFF
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
      Image           =   "uc戰鬥系統牌型介面.ctx":173D
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
      Image           =   "uc戰鬥系統牌型介面.ctx":207B
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
      Image           =   "uc戰鬥系統牌型介面.ctx":29B9
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
      Image           =   "uc戰鬥系統牌型介面.ctx":3323
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
      Image           =   "uc戰鬥系統牌型介面.ctx":3C8D
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
      Image           =   "uc戰鬥系統牌型介面.ctx":45F7
      Opacity         =   70
      Mirror          =   1
      Props           =   5
   End
   Begin VB.Image stagejpgn 
      Height          =   270
      Left            =   9120
      Picture         =   "uc戰鬥系統牌型介面.ctx":4F61
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
      Left            =   720
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
      Picture         =   "uc戰鬥系統牌型介面.ctx":543E
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
      Picture         =   "uc戰鬥系統牌型介面.ctx":2F878
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
   Begin VB.Image cardunderjpg 
      Height          =   270
      Left            =   0
      Picture         =   "uc戰鬥系統牌型介面.ctx":2FD22
      Top             =   1080
      Width           =   2280
   End
End
Attribute VB_Name = "uc戰鬥系統牌型介面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Turn As Integer, m_cardnum As Integer, m_passivevisble As Boolean
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
        pageul.FontSize = 20
   Else
        pageul.FontSize = 24
   End If
End Property
Public Property Get stagejpg() As String
   stagejpg = ""
End Property
Public Property Let stagejpg(ByVal New_Stagejpg As String)
   PropertyChanged "Stagejpg"
   '=================
   If New_Stagejpg <> "" Then
       stagejpgn.Picture = LoadPicture(New_Stagejpg)
   End If
End Property
Public Property Get Passive_使用者_技能名稱() As String
   Passive_使用者_技能名稱 = ""
End Property
Public Property Let Passive_使用者_技能名稱(ByVal New_Passive_使用者_技能名稱 As String)
   PropertyChanged "Passive_使用者_技能名稱"
   '=================
   Dim pstr() As String
   pstr = Split(New_Passive_使用者_技能名稱, "#")
   If pstr(0) <> "" And Val(pstr(1)) >= 1 And Val(pstr(1)) <= 4 Then
       passivetext_us(Val(pstr(1))).Caption = pstr(0)
   End If
End Property
Public Property Get Passive_使用者_技能燈發亮() As Integer
   Passive_使用者_技能燈發亮 = 0
End Property
Public Property Let Passive_使用者_技能燈發亮(ByVal New_Passive_使用者_技能燈發亮 As Integer)
   PropertyChanged "Passive_使用者_技能燈發亮"
   '=================
   If New_Passive_使用者_技能燈發亮 >= 1 And New_Passive_使用者_技能燈發亮 <= 4 Then
       passivelight_us(New_Passive_使用者_技能燈發亮).ClearImage
       passivelight_us(New_Passive_使用者_技能燈發亮).LoadImage_FromFile App.Path & "\gif\system\passivelighton.png"
       passivelight_us(New_Passive_使用者_技能燈發亮).Mirror = aiMirrorHorizontal
'        passivelight_us(New_Passive_使用者_技能燈發亮).Picture = LoadPicture(App.Path & "\gif\system\passivelightonus.gif")
   End If
End Property
Public Property Get Passive_使用者_技能燈變暗() As Integer
   Passive_使用者_技能燈變暗 = 0
End Property
Public Property Let Passive_使用者_技能燈變暗(ByVal New_Passive_使用者_技能燈變暗 As Integer)
   PropertyChanged "Passive_使用者_技能燈變暗"
   '=================
   If New_Passive_使用者_技能燈變暗 >= 1 And New_Passive_使用者_技能燈變暗 <= 4 Then
       passivelight_us(New_Passive_使用者_技能燈變暗).ClearImage
       passivelight_us(New_Passive_使用者_技能燈變暗).LoadImage_FromFile App.Path & "\gif\system\passivelightoff.png"
       passivelight_us(New_Passive_使用者_技能燈變暗).Mirror = aiMirrorHorizontal
   End If
End Property
Public Property Get Passive_使用者_技能顯示() As Integer
   Passive_使用者_技能顯示 = 0
End Property
Public Property Let Passive_使用者_技能顯示(ByVal New_Passive_使用者_技能顯示 As Integer)
   PropertyChanged "Passive_使用者_技能顯示"
   '=================
   If New_Passive_使用者_技能顯示 >= 1 And New_Passive_使用者_技能顯示 <= 4 Then
       passivelight_us(New_Passive_使用者_技能顯示).Visible = True
       passivetext_us(New_Passive_使用者_技能顯示).Visible = True
   End If
End Property
Public Property Get Passive_使用者_技能隱藏() As Integer
   Passive_使用者_技能隱藏 = 0
End Property
Public Property Let Passive_使用者_技能隱藏(ByVal New_Passive_使用者_技能隱藏 As Integer)
   PropertyChanged "Passive_使用者_技能隱藏"
   '=================
   If New_Passive_使用者_技能隱藏 >= 1 And New_Passive_使用者_技能隱藏 <= 4 Then
       passivelight_us(New_Passive_使用者_技能隱藏).Visible = False
       passivetext_us(New_Passive_使用者_技能隱藏).Visible = False
   End If
End Property
Public Property Get Passive_電腦_技能顯示() As Integer
   Passive_電腦_技能顯示 = 0
End Property
Public Property Let Passive_電腦_技能顯示(ByVal New_Passive_電腦_技能顯示 As Integer)
   PropertyChanged "Passive_電腦_技能顯示"
   '=================
   If New_Passive_電腦_技能顯示 >= 1 And New_Passive_電腦_技能顯示 <= 4 Then
       passivelight_com(New_Passive_電腦_技能顯示).Visible = True
       passivetext_com(New_Passive_電腦_技能顯示).Visible = True
   End If
End Property
Public Property Get Passive_電腦_技能隱藏() As Integer
   Passive_電腦_技能隱藏 = 0
End Property
Public Property Let Passive_電腦_技能隱藏(ByVal New_Passive_電腦_技能隱藏 As Integer)
   PropertyChanged "Passive_電腦_技能隱藏"
   '=================
   If New_Passive_電腦_技能隱藏 >= 1 And New_Passive_電腦_技能隱藏 <= 4 Then
       passivelight_com(New_Passive_電腦_技能隱藏).Visible = False
       passivetext_com(New_Passive_電腦_技能隱藏).Visible = False
   End If
End Property
Public Property Get Passive_電腦_技能燈發亮() As Integer
   Passive_電腦_技能燈發亮 = 0
End Property
Public Property Let Passive_電腦_技能燈發亮(ByVal New_Passive_電腦_技能燈發亮 As Integer)
   PropertyChanged "Passive_電腦_技能燈發亮"
   '=================
   If New_Passive_電腦_技能燈發亮 >= 1 And New_Passive_電腦_技能燈發亮 <= 4 Then
       passivelight_com(New_Passive_電腦_技能燈發亮).ClearImage
       passivelight_com(New_Passive_電腦_技能燈發亮).LoadImage_FromFile App.Path & "\gif\system\passivelighton.png"
       passivelight_com(New_Passive_電腦_技能燈發亮).Mirror = aiMirrorNone
'        passivelight_com(New_Passive_使用者_技能燈發亮).Picture = LoadPicture(App.Path & "\gif\system\passivelightoncom.gif")
   End If
End Property
Public Property Get Passive_電腦_技能燈變暗() As Integer
   Passive_電腦_技能燈變暗 = 0
End Property
Public Property Let Passive_電腦_技能燈變暗(ByVal New_Passive_電腦_技能燈變暗 As Integer)
   PropertyChanged "Passive_電腦_技能燈變暗"
   '=================
   If New_Passive_電腦_技能燈變暗 >= 1 And New_Passive_電腦_技能燈變暗 <= 4 Then
       passivelight_com(New_Passive_電腦_技能燈變暗).ClearImage
       passivelight_com(New_Passive_電腦_技能燈變暗).LoadImage_FromFile App.Path & "\gif\system\passivelightoff.png"
       passivelight_com(New_Passive_電腦_技能燈變暗).Mirror = aiMirrorNone
   End If
End Property
Public Property Get Passive_電腦_技能名稱() As String
   Passive_電腦_技能名稱 = ""
End Property
Public Property Let Passive_電腦_技能名稱(ByVal New_Passive_電腦_技能名稱 As String)
   PropertyChanged "Passive_電腦_技能名稱"
   '=================
   Dim pstr() As String
   pstr = Split(New_Passive_電腦_技能名稱, "#")
   If pstr(0) <> "" And Val(pstr(1)) >= 1 And Val(pstr(1)) <= 4 Then
       passivetext_com(Val(pstr(1))).Caption = pstr(0)
   End If
End Property
Public Property Get Passive_技能一方全重設() As Integer
   Passive_技能一方全重設 = 0
End Property
Public Property Let Passive_技能一方全重設(ByVal New_Passive_技能一方全重設 As Integer)
   PropertyChanged "Passive_技能一方全重設"
   '=================
   Dim i As Integer
   Select Case New_Passive_技能一方全重設
       Case 1
           For i = 1 To 4
               passivelight_us(i).Visible = False
               Me.Passive_使用者_技能燈變暗 = i
               passivetext_us(i).Visible = False
               passivetext_us(i).Caption = ""
           Next
       Case 2
           For i = 1 To 4
               passivelight_com(i).Visible = False
               Me.Passive_電腦_技能燈變暗 = i
               passivetext_com(i).Visible = False
               passivetext_com(i).Caption = ""
           Next
   End Select
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
          Me.Passive_使用者_技能隱藏 = i
          Me.Passive_電腦_技能隱藏 = i
       Next
   Else
       cardunderjpg.Visible = True
       cardpagejpg.Visible = True
       pageul.Visible = True
       pageul.ZOrder
       stagejpgn.Visible = True
   End If
End Property
Public Property Get Message() As String
   Message = ""
End Property
Public Property Let Message(ByVal New_Message As String)
   messagetext.MeaageText = New_Message
   PropertyChanged "Message"
End Property
Sub MessageClear()
messagetext.MessageTextClear
End Sub

