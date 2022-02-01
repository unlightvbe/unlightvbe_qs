VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl 小人物形象 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ClipBehavior    =   0  '無
   ScaleHeight     =   5535
   ScaleWidth      =   2880
   Windowless      =   -1  'True
   Begin VB.Timer t2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   4800
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   4800
   End
   Begin ImageX.aicAlphaImage image1 
      Height          =   3255
      Left            =   120
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5741
      Image           =   "uc小人物形象.ctx":0000
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image2 
      Height          =   855
      Left            =   120
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      Image           =   "uc小人物形象.ctx":0018
      Scaler          =   3
   End
End
Attribute VB_Name = "小人物形象"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_totwidth As Integer
Dim m_totheight As Integer
Dim m_smalldowntop As Integer
Dim m_smallimage As String
Dim m_smallimagedown As String
Dim m_smalldownleft As Integer
Dim m_smallhei As Integer
Dim m_smallwh As Integer
Dim m_smalldeath As Boolean
Dim m_smalllive As Boolean
Dim m_smallreturn As Boolean
Dim m_smallreset As Boolean

Public Property Get 小人物影子Left() As Integer
   小人物影子Left = m_smalldownleft
End Property
Public Property Get 小人物影子top差() As Integer
   小人物影子top差 = m_smalldowntop
End Property
Public Property Get 小人物圖片() As String
   小人物圖片 = m_smallimage
End Property
Public Property Get 小人物圖片height() As Integer
   小人物圖片height = m_smallhei
End Property
Public Property Get 小人物圖片width() As Integer
   小人物圖片width = m_smallwh
End Property
Public Property Get 小人物影子圖片() As String
   小人物影子圖片 = m_smallimagedown
End Property
Public Property Let 小人物影子Left(ByVal New_小人物影子Left As Integer)
   m_smalldownleft = New_小人物影子Left
   PropertyChanged "小人物影子Left"
   Image2.Left = Me.小人物影子Left
End Property
Public Property Let 小人物影子top差(ByVal New_小人物影子top差 As Integer)
   m_smalldowntop = New_小人物影子top差
   PropertyChanged "小人物影子top差"
   Image2.Top = image1.Height + Me.小人物影子top差
End Property
Public Property Let 小人物圖片(ByVal New_小人物圖片 As String)
   m_smallimage = New_小人物圖片
   PropertyChanged "小人物圖片"
   If Me.小人物圖片 <> "" Then
       image1.AutoRedraw = True
       image1.AutoSize = True
       image1.LoadImage_FromFile Me.小人物圖片
       image1.Left = 0
       image1.Top = 0
       Me.小人物圖片height = image1.Height
       Me.小人物圖片width = image1.Width
       Me.小人物消失 = False
       Me.小人物顯現 = False
   End If
End Property
Public Property Let 小人物影子圖片(ByVal New_小人物影子圖片 As String)
   m_smallimagedown = New_小人物影子圖片
   PropertyChanged "小人物影子圖片"
   If Me.小人物影子圖片 <> "" Then
       Image2.AutoRedraw = True
       Image2.AutoSize = True
       Image2.LoadImage_FromFile Me.小人物影子圖片
       Image2.Left = 0
       Image2.Top = image1.Height
       Image2.Opacity = 100
   End If
End Property
Public Property Let 小人物圖片height(ByVal New_小人物圖片height As Integer)
   m_smallhei = New_小人物圖片height
   PropertyChanged "小人物圖片height"
End Property
Public Property Let 小人物圖片width(ByVal New_小人物圖片width As Integer)
   m_smallwh = New_小人物圖片width
   PropertyChanged "小人物圖片width"
End Property
Public Property Get 小人物消失() As Boolean
   小人物消失 = m_smalldeath
End Property
Public Property Let 小人物消失(ByVal New_小人物消失 As Boolean)
   m_smalldeath = New_小人物消失
   PropertyChanged "小人物消失"
   '=====================
   If Me.小人物消失 = True Then
       t1.Enabled = True
   End If
End Property
Public Property Get 小人物重設() As Boolean
   小人物重設 = m_smallreset
End Property
Public Property Let 小人物重設(ByVal New_小人物重設 As Boolean)
   m_smallreset = New_小人物重設
   PropertyChanged "小人物重設"
   '=====================
   If Me.小人物重設 = True Then
       image1.Opacity = 100
       Image2.Opacity = 100
       Me.小人物影像反轉 = False
       Me.小人物消失 = False
       Me.小人物顯現 = False
       Me.小人物重設 = False
   End If
End Property
Public Property Get 小人物顯現() As Boolean
   小人物顯現 = m_smalllive
End Property
Public Property Let 小人物顯現(ByVal New_小人物顯現 As Boolean)
   m_smalllive = New_小人物顯現
   PropertyChanged "小人物顯現"
   '=====================
   If Me.小人物顯現 = True Then
       t2.Enabled = True
   End If
End Property
Public Property Get 小人物影像反轉() As Boolean
   小人物影像反轉 = m_smallreturn
End Property
Public Property Let 小人物影像反轉(ByVal New_小人物影像反轉 As Boolean)
   m_smallreturn = New_小人物影像反轉
   PropertyChanged "小人物影像反轉"
   '=====================
   If Me.小人物影像反轉 = True Then
       image1.Mirror = aiMirrorHorizontal
       Image2.Mirror = aiMirrorHorizontal
   Else
       image1.Mirror = aiMirrorNone
       Image2.Mirror = aiMirrorNone
   End If
End Property
Private Sub t1_Timer()
If image1.Opacity <> 0 Then
    image1.Opacity = Val(image1.Opacity) - 1
End If
If Image2.Opacity <> 0 Then
    Image2.Opacity = Val(Image2.Opacity) - 1
End If
If image1.Opacity = 0 And Image2.Opacity = 0 Then
    t1.Enabled = False
    Me.小人物消失 = False
End If
End Sub

Private Sub t2_Timer()
If image1.Opacity <> 100 Then
    image1.Opacity = Val(image1.Opacity) + 2
End If
If Image2.Opacity <> 100 Then
    Image2.Opacity = Val(Image2.Opacity) + 2
End If
If image1.Opacity = 100 And Image2.Opacity = 100 Then
    t2.Enabled = False
    Me.小人物顯現 = False
End If
End Sub
