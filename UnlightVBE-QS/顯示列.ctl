VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl 顯示列 
   BackStyle       =   0  '透明
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12720
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   3150
   ScaleWidth      =   12720
   Windowless      =   -1  'True
   Begin VB.Timer trmovehide 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   8330
      Top             =   1200
   End
   Begin VB.Timer trmoveshow 
      Enabled         =   0   'False
      Interval        =   130
      Left            =   5160
      Top             =   1080
   End
   Begin VB.Label g2 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Black"
         Size            =   36
         Charset         =   136
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7590
      TabIndex        =   1
      Top             =   -120
      Width           =   1400
   End
   Begin VB.Label g1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Black"
         Size            =   39.75
         Charset         =   136
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2440
      TabIndex        =   0
      Top             =   0
      Width           =   1400
   End
   Begin ImageX.aicAlphaImage bn42 
      Height          =   1335
      Left            =   4320
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "顯示列.ctx":0000
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn4 
      Height          =   1335
      Left            =   10320
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "顯示列.ctx":0018
      Scaler          =   3
   End
   Begin VB.Image moverightjpg 
      Height          =   720
      Index           =   0
      Left            =   10080
      Picture         =   "顯示列.ctx":0030
      Top             =   480
      Width           =   525
   End
   Begin VB.Image moveleftjpg 
      Height          =   720
      Index           =   0
      Left            =   2760
      Picture         =   "顯示列.ctx":01F5
      Top             =   360
      Width           =   525
   End
   Begin ImageX.aicAlphaImage bn32 
      Height          =   1335
      Left            =   2760
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "顯示列.ctx":03B8
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn22 
      Height          =   1095
      Left            =   1800
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      Image           =   "顯示列.ctx":03D0
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn12 
      Height          =   975
      Left            =   600
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Image           =   "顯示列.ctx":03E8
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn3 
      Height          =   1455
      Left            =   8880
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2566
      Image           =   "顯示列.ctx":0400
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn2 
      Height          =   1215
      Left            =   7320
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Image           =   "顯示列.ctx":0418
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn1 
      Height          =   1215
      Left            =   6120
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Image           =   "顯示列.ctx":0430
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image2 
      Height          =   1095
      Left            =   6960
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      Image           =   "顯示列.ctx":0448
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image1 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      Image           =   "顯示列.ctx":0460
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage aie1 
      Height          =   1575
      Left            =   -120
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   2778
      Image           =   "顯示列.ctx":0478
      Scaler          =   3
      Props           =   17
   End
End
Attribute VB_Name = "顯示列"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_smallimage As String
Dim m_smallimageus As String
Dim m_smallimagecom As String
Dim m_movetn As Boolean
Dim m_g1 As Integer
Dim m_g2 As Integer
Dim m_smallimageusleft As Integer
Dim m_smallimagecomleft As Integer
Dim m_g1v As Boolean
Dim m_g2v As Boolean
Dim m_bnc As Integer
Dim m_moveleftnum As Integer
Dim m_moveleftio As Integer
Dim m_moverightnum As Integer
Dim m_moverightio As Integer
Dim 移動圖片顯示數(1 To 2, 1 To 3) As Integer '移動顯示計數器暫時變數(1.使用者/2.電腦,1.目前數/2.方向-(1)向內(2)向外/3.目標最大數)
Dim 移動圖片顯示完成數(1 To 2) As Boolean '移動顯示計數器是否已完成變數(1.使用者/2.電腦)
Dim trmovehidetime As Integer '移動顯示計數器暫時變數
Dim m_moveleftrightc As Boolean
Dim m_smallimageuswidth As Integer
Dim m_smallimagecomwidth As Integer
Dim m_personvs As Integer

Public Property Get 人物戰鬥人數() As Integer
   人物戰鬥人數 = m_personvs
End Property
Public Property Let 人物戰鬥人數(ByVal new_人物戰鬥人數 As Integer)
   m_personvs = new_人物戰鬥人數
   PropertyChanged "人物戰鬥人數"
   Select Case Me.人物戰鬥人數
       Case 1
            bn1.Left = 4200
            bn1.Top = 100
            bn2.Left = 5340
            bn2.Top = 100
            bn3.Left = 6480
            bn3.Top = 100
            bn12.Left = 4200
            bn12.Top = 100
            bn22.Left = 5340
            bn22.Top = 100
            bn32.Left = 6480
            bn32.Top = 100
            bn4.Visible = False
            bn42.Visible = False
       Case 3
            bn1.Left = 4080
            bn1.Top = 100
            bn2.Left = 4920
            bn2.Top = 100
            bn3.Left = 6600
            bn3.Top = 100
            bn4.Left = 5760
            bn4.Top = 100
            bn12.Left = 4080
            bn12.Top = 100
            bn22.Left = 4920
            bn22.Top = 100
            bn32.Left = 6600
            bn32.Top = 100
            bn42.Left = 5760
            bn42.Top = 100
            bn4.Visible = True
            bn42.Visible = True
   End Select
End Property
Public Property Get 顯示列圖片() As String
   顯示列圖片 = m_smallimage
End Property
Public Property Get 使用者方小人物圖片width() As Integer
   使用者方小人物圖片width = m_smallimageuswidth
End Property
Public Property Get 電腦方小人物圖片width() As Integer
   電腦方小人物圖片width = m_smallimagecomwidth
End Property
Public Property Get 移動方向圖片顯示() As Boolean
   移動方向圖片顯示 = m_moveleftrightc
End Property
Public Property Get 使用者方移動值() As Integer
   使用者方移動值 = m_moveleftnum
End Property
Public Property Get 使用者方移動內外() As Integer
   使用者方移動內外 = m_moveleftio
End Property
Public Property Get 電腦方移動值() As Integer
   電腦方移動值 = m_moverightnum
End Property
Public Property Get 電腦方移動內外() As Integer
   電腦方移動內外 = m_moverightio
End Property
Public Property Get 移動階段圖顯示() As Boolean
   移動階段圖顯示 = m_movetn
End Property
Public Property Get goi1() As Integer
   goi1 = m_g1
End Property
Public Property Get goi2() As Integer
   goi2 = m_g2
End Property
Public Property Get 移動階段選擇值() As Integer
   移動階段選擇值 = m_bnc
End Property
Public Property Get 使用者方小人物圖片left() As Integer
   使用者方小人物圖片left = m_smallimageusleft
End Property
Public Property Get 電腦方小人物圖片left() As Integer
   電腦方小人物圖片left = m_smallimagecomleft
End Property
Public Property Get goi1顯示() As Boolean
   goi1顯示 = m_g1v
End Property
Public Property Get goi2顯示() As Boolean
   goi2顯示 = m_g2v
End Property
Public Property Get 使用者方小人物圖片() As String
   使用者方小人物圖片 = m_smallimageus
End Property
Public Property Let 使用者方小人物圖片(ByVal New_使用者方小人物圖片 As String)
   m_smallimageus = New_使用者方小人物圖片
   PropertyChanged "使用者方小人物圖片"
   If Me.使用者方小人物圖片 <> "" Then
       Image1.AutoSize = True
       Image1.AutoRedraw = True
       Image1.LoadImage_FromFile Me.使用者方小人物圖片
       Image1.Top = 0
       Image1.Left = 0
       Me.使用者方小人物圖片width = Image1.Width
       Image1.Mirror = aiMirrorNone
    End If
End Property
Public Property Get 電腦方小人物圖片() As String
   電腦方小人物圖片 = m_smallimagecom
End Property
Public Property Let 電腦方小人物圖片(ByVal New_電腦方小人物圖片 As String)
   m_smallimagecom = New_電腦方小人物圖片
   PropertyChanged "電腦方小人物圖片"
   If Me.電腦方小人物圖片 <> "" Then
       Image2.AutoSize = True
       Image2.AutoRedraw = True
       Image2.LoadImage_FromFile Me.電腦方小人物圖片
       Image2.Top = 0
       Image2.Left = 7680
       Image2.Mirror = aiMirrorHorizontal
    End If
    Me.電腦方小人物圖片width = Image2.Width
End Property
Public Property Let 顯示列圖片(ByVal new_顯示列圖片 As String)
   m_smallimage = new_顯示列圖片
   PropertyChanged "顯示列圖片"
   If Me.顯示列圖片 <> "" Then
       aie1.AutoRedraw = True
       aie1.AutoSize = True
       aie1.LoadImage_FromFile Me.顯示列圖片
       aie1.Left = 0
       aie1.Top = 0
   End If
End Property
Public Property Let goi2(ByVal newgoi2 As Integer)
   m_g2 = newgoi2
   PropertyChanged "goi2"
   g2.Caption = Me.goi2
End Property
Public Property Let goi1(ByVal newgoi1 As Integer)
   m_g1 = newgoi1
   PropertyChanged "goi1"
   g1.Caption = Me.goi1
End Property
Public Property Let goi1顯示(ByVal newgoi1v As Boolean)
   m_g1v = newgoi1v
   PropertyChanged "goi1顯示"
   If Me.goi1顯示 = False Then
       g1.Visible = False
    Else
       g1.Visible = True
        If g1.FontName = "Noto Sans T Chinese Black" Then
            g1.Top = -160
            g1.FontSize = 40
        Else
            g1.Top = 0
            g1.FontSize = 36
        End If
    End If
End Property
Public Property Let goi2顯示(ByVal newgoi2v As Boolean)
   m_g2v = newgoi2v
   PropertyChanged "goi2顯示"
   If Me.goi2顯示 = False Then
       g2.Visible = False
    Else
       g2.Visible = True
        If g2.FontName = "Noto Sans T Chinese Black" Then
            g2.Top = -160
            g2.FontSize = 40
        Else
            g2.Top = 0
            g2.FontSize = 36
        End If
    End If
End Property
Public Property Let 使用者方小人物圖片left(ByVal new使用者方小人物圖片left As Integer)
    m_smallimageusleft = new使用者方小人物圖片left
   PropertyChanged "使用者方小人物圖片left"
   Image1.Left = Me.使用者方小人物圖片left
End Property
Public Property Let 電腦方小人物圖片left(ByVal new電腦方小人物圖片left As Integer)
    m_smallimagecomleft = new電腦方小人物圖片left
   PropertyChanged "電腦方小人物圖片left"
   Image2.Left = Me.電腦方小人物圖片left
End Property

Public Property Let 移動階段選擇值(ByVal new移動階段選擇值 As Integer)
    m_bnc = new移動階段選擇值
   PropertyChanged "移動階段選擇值"
   移動階段圖顯示_階段
End Property
Public Property Let 使用者方移動值(ByVal new使用者方移動值 As Integer)
   m_moveleftnum = new使用者方移動值
   PropertyChanged "使用者方移動值"
   移動圖片顯示數(1, 3) = Me.使用者方移動值
End Property
Public Property Let 使用者方移動內外(ByVal new使用者方移動內外 As Integer)
   m_moveleftio = new使用者方移動內外
   PropertyChanged "使用者方移動內外"
   移動圖片顯示數(1, 2) = Me.使用者方移動內外
End Property
Public Property Let 電腦方移動值(ByVal new電腦方移動值 As Integer)
   m_moverightnum = new電腦方移動值
   PropertyChanged "電腦方移動值"
   移動圖片顯示數(2, 3) = Me.電腦方移動值
End Property
Public Property Let 電腦方移動內外(ByVal new電腦方移動內外 As Integer)
   m_moverightio = new電腦方移動內外
   PropertyChanged "電腦方移動內外"
   移動圖片顯示數(2, 2) = Me.電腦方移動內外
End Property
Public Property Let 使用者方小人物圖片width(ByVal new使用者方小人物圖片width As Integer)
   m_smallimageuswidth = new使用者方小人物圖片width
   PropertyChanged "使用者方小人物圖片width"
End Property
Public Property Let 電腦方小人物圖片width(ByVal new電腦方小人物圖片width As Integer)
   m_smallimagecomwidth = new電腦方小人物圖片width
   PropertyChanged "電腦方小人物圖片width"
End Property
Public Property Let 移動方向圖片顯示(ByVal new移動方向圖片顯示 As Boolean)
   m_moveleftrightc = new移動方向圖片顯示
   PropertyChanged "移動方向圖片顯示"
   If Me.移動方向圖片顯示 = True Then
         移動圖片顯示數(1, 1) = 1
         移動圖片顯示數(2, 1) = 1
         移動圖片顯示完成數(1) = False
         移動圖片顯示完成數(2) = False
         '=============================
         For i = 1 To 移動圖片顯示數(1, 3)
             Load moveleftjpg(i)
             moveleftjpg(i).Left = 2400 + i * 300
             moveleftjpg(i).Top = 120
         Next
         For i = 1 To 移動圖片顯示數(2, 3)
             Load moverightjpg(i)
             moverightjpg(i).Left = 8520 - i * 300
             moverightjpg(i).Top = 120
         Next
         '=============================
         trmovehidetime = 1
         trmoveshow.Enabled = True
    End If
End Property
Public Property Let 移動階段圖顯示(ByVal new移動階段圖顯示 As Boolean)
    m_movetn = new移動階段圖顯示
   PropertyChanged "移動階段圖顯示"
   If Me.移動階段圖顯示 = True Then
       bn1.Visible = True
       bn2.Visible = True
       bn3.Visible = True
       If Me.人物戰鬥人數 = 3 Then
           bn4.Visible = True
       Else
           bn4.Visible = False
       End If
       Me.移動階段選擇值 = 0
    Else
       bn1.Visible = False
       bn2.Visible = False
       bn3.Visible = False
       bn4.Visible = False
       bn12.Visible = False
       bn22.Visible = False
       bn32.Visible = False
       bn42.Visible = False
    End If
End Property
Sub 移動階段圖顯示_階段()
   Select Case Me.移動階段選擇值
      Case 0
            bn12.Visible = False
            bn22.Visible = False
            bn32.Visible = False
            bn42.Visible = False
      Case 1
            bn12.Visible = True
            bn22.Visible = False
            bn32.Visible = False
            bn42.Visible = False
      Case 2
            bn12.Visible = False
            bn22.Visible = True
            bn32.Visible = False
            bn42.Visible = False
      Case 3
            bn12.Visible = False
            bn22.Visible = False
            bn32.Visible = True
            bn42.Visible = False
      Case 4
            bn12.Visible = False
            bn22.Visible = False
            bn32.Visible = False
            bn42.Visible = True
   End Select
End Sub
Private Sub aie1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
移動階段圖顯示_階段
End Sub

Private Sub bn1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn12.Visible = True
End Sub

Private Sub bn12_Click(ByVal Button As Integer)
Me.移動階段選擇值 = 1
End Sub

Private Sub bn2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn22.Visible = True
End Sub

Private Sub bn22_Click(ByVal Button As Integer)
Me.移動階段選擇值 = 2
End Sub

Private Sub bn3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn32.Visible = True
End Sub

Private Sub bn32_Click(ByVal Button As Integer)
Me.移動階段選擇值 = 3
End Sub



Private Sub bn4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn42.Visible = True
End Sub

Private Sub bn42_Click(ByVal Button As Integer)
Me.移動階段選擇值 = 4
End Sub

Private Sub trmovehide_Timer()
Select Case trmovehidetime
 Case 2
   If 移動圖片顯示數(1, 2) = 1 And 移動圖片顯示數(2, 2) = 2 Then
     If 移動圖片顯示數(1, 3) > 0 And 移動圖片顯示數(2, 3) > 0 Then
       If moveleftjpg(移動圖片顯示數(1, 3)).Visible = True And moverightjpg(移動圖片顯示數(2, 3)).Visible = True Then
          moveleftjpg(移動圖片顯示數(1, 3)).Visible = False
          moverightjpg(移動圖片顯示數(2, 3)).Visible = False
          移動圖片顯示數(1, 3) = 移動圖片顯示數(1, 3) - 1
          移動圖片顯示數(2, 3) = 移動圖片顯示數(2, 3) - 1
          Exit Sub
       End If
     End If
   ElseIf 移動圖片顯示數(1, 2) = 2 And 移動圖片顯示數(2, 2) = 1 Then
     If 移動圖片顯示數(1, 3) > 0 And 移動圖片顯示數(2, 3) > 0 Then
       If moveleftjpg(移動圖片顯示數(1, 3)).Visible = True And moverightjpg(移動圖片顯示數(2, 3)).Visible = True Then
          moveleftjpg(移動圖片顯示數(1, 3)).Visible = False
          moverightjpg(移動圖片顯示數(2, 3)).Visible = False
          移動圖片顯示數(1, 3) = 移動圖片顯示數(1, 3) - 1
          移動圖片顯示數(2, 3) = 移動圖片顯示數(2, 3) - 1
          Exit Sub
       End If
     End If
   End If
   trmovehidetime = trmovehidetime + 1
 Case 10
      '===假如都不符合條件時之動作
      trmovehide.Enabled = False
      '=========================
      For i = 1 To moveleftjpg.UBound
          moveleftjpg(i).Visible = False
          Unload moveleftjpg(i)
      Next
      For i = 1 To moverightjpg.UBound
          moverightjpg(i).Visible = False
          Unload moverightjpg(i)
      Next
      '=========================
      Me.移動方向圖片顯示 = False
 Case Else
      trmovehidetime = trmovehidetime + 1
End Select
End Sub

Private Sub trmoveshow_Timer()
If 移動圖片顯示數(1, 1) <= 移動圖片顯示數(1, 3) Then
   If 移動圖片顯示數(1, 2) = 1 Then
      moveleftjpg(移動圖片顯示數(1, 1)).Picture = LoadPicture(App.Path & "\gif\system\movein.gif")
      moveleftjpg(移動圖片顯示數(1, 1)).Visible = True
      moveleftjpg(移動圖片顯示數(1, 1)).ZOrder
   Else
      moveleftjpg(移動圖片顯示數(1, 1)).Picture = LoadPicture(App.Path & "\gif\system\moveout.gif")
      moveleftjpg(移動圖片顯示數(1, 1)).Visible = True
      moveleftjpg(移動圖片顯示數(1, 1)).ZOrder
   End If
   移動圖片顯示數(1, 1) = 移動圖片顯示數(1, 1) + 1
Else
   移動圖片顯示完成數(1) = True
End If

If 移動圖片顯示數(2, 1) <= 移動圖片顯示數(2, 3) Then
   If 移動圖片顯示數(2, 2) = 1 Then
      moverightjpg(移動圖片顯示數(2, 1)).Picture = LoadPicture(App.Path & "\gif\system\moveout.gif")
      moverightjpg(移動圖片顯示數(2, 1)).Visible = True
      moverightjpg(移動圖片顯示數(2, 1)).ZOrder
   Else
      moverightjpg(移動圖片顯示數(2, 1)).Picture = LoadPicture(App.Path & "\gif\system\movein.gif")
      moverightjpg(移動圖片顯示數(2, 1)).Visible = True
      moverightjpg(移動圖片顯示數(2, 1)).ZOrder
   End If
   移動圖片顯示數(2, 1) = 移動圖片顯示數(2, 1) + 1
Else
   移動圖片顯示完成數(2) = True
End If

If 移動圖片顯示完成數(1) = True And 移動圖片顯示完成數(2) = True Then
trmoveshow.Enabled = False
移動圖片顯示完成數(1) = False
移動圖片顯示完成數(2) = False
trmovehide.Enabled = True
End If
End Sub

Private Sub UserControl_Show()
bn1.AutoRedraw = True
bn1.AutoSize = True
bn2.AutoRedraw = True
bn2.AutoSize = True
bn3.AutoRedraw = True
bn3.AutoSize = True
bn4.AutoRedraw = True
bn4.AutoSize = True
bn12.AutoRedraw = True
bn12.AutoSize = True
bn22.AutoRedraw = True
bn22.AutoSize = True
bn32.AutoRedraw = True
bn32.AutoSize = True
bn42.AutoRedraw = True
bn42.AutoSize = True
bn1.LoadImage_FromFile App.Path & "\gif\system\left_1.png"
bn2.LoadImage_FromFile App.Path & "\gif\system\rest_1.png"
bn3.LoadImage_FromFile App.Path & "\gif\system\right_1.png"
bn4.LoadImage_FromFile App.Path & "\gif\system\change_1.png"
bn12.LoadImage_FromFile App.Path & "\gif\system\left_2.png"
bn22.LoadImage_FromFile App.Path & "\gif\system\rest_2.png"
bn32.LoadImage_FromFile App.Path & "\gif\system\right_2.png"
bn42.LoadImage_FromFile App.Path & "\gif\system\change_2.png"
'Me.goi1顯示 = True
'Me.goi2顯示 = True
'Me.使用者方小人物圖片left = 0
'Me.電腦方小人物圖片left = 7680
'bn1.Left = 4200
'bn1.Top = 100
'bn2.Left = 5340
'bn2.Top = 100
'bn3.Left = 6480
'bn3.Top = 100
'bn12.Left = 4200
'bn12.Top = 100
'bn22.Left = 5340
'bn22.Top = 100
'bn32.Left = 6480
'bn32.Top = 100
'bn4.Visible = False
'bn42.Visible = False
Me.移動階段圖顯示 = False
Me.移動階段選擇值 = 0

moveleftjpg(0).Left = 2400
moveleftjpg(0).Top = 120
moverightjpg(0).Left = 8520
moverightjpg(0).Top = 120
moveleftjpg(0).Visible = False
moverightjpg(0).Visible = False
Me.移動方向圖片顯示 = False
End Sub


