VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc角色卡片介面_前 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ClipBehavior    =   0  '無
   HitBehavior     =   2  '使用繪圖區域
   ScaleHeight     =   5190
   ScaleWidth      =   4785
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage cardbackclick 
      Height          =   795
      Left            =   480
      Top             =   1560
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1402
      Image           =   "uc角色卡片介面_前.ctx":0000
      Props           =   13
   End
   Begin ImageX.aicAlphaImage cardback 
      Height          =   3600
      Left            =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   6350
      Image           =   "uc角色卡片介面_前.ctx":2735
      Props           =   9
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   1
      Left            =   330
      TabIndex        =   0
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   7
      Left            =   330
      TabIndex        =   1
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   6
      Left            =   330
      TabIndex        =   2
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   5
      Left            =   330
      TabIndex        =   3
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   4
      Left            =   330
      TabIndex        =   4
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   3
      Left            =   330
      TabIndex        =   5
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   2
      Left            =   330
      TabIndex        =   6
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   14
      Left            =   1320
      TabIndex        =   7
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   13
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   12
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   10
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   11
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   12
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin UnlightVBE.uc異常狀態 personspe 
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   13
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin VB.Label personlabeldef 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "2"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label personlabelatk 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "2"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label personlabelhp 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "2"
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
      Left            =   600
      TabIndex        =   14
      Top             =   3240
      Width           =   375
   End
   Begin ImageX.aicAlphaImage cardImage 
      Height          =   3585
      Left            =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   6324
      Image           =   "uc角色卡片介面_前.ctx":2980
      Props           =   9
   End
End
Attribute VB_Name = "uc角色卡片介面_前"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_cardmain_jpg As String
Dim m_cardmain_personhp As Integer, m_cardmain_personhpmax As Integer, m_cardmain_personhp41 As Integer
Dim m_cardmain_personatk As Integer
Dim m_cardmain_persondef As Integer
Dim m_cardmain_isnewtype As Boolean
Public Event CardClick()
Public Sub 更改異常狀態資料(ByVal buffnum As Integer, ByVal ImagePath As String, ByVal num As Integer, ByVal tot As Integer, ByVal isVisible As Boolean)
If buffnum >= 1 And buffnum <= 14 Then
    If isVisible = False Then
        personspe(buffnum).Visible = False
    Else
        personspe(buffnum).person_num = num
        personspe(buffnum).person_turn = tot
        personspe(buffnum).異常狀態圖片 = ImagePath
        personspe(buffnum).Visible = True
    End If
End If
End Sub
Public Sub 異常狀態全重設()
Dim i As Integer
For i = 1 To 14
    personspe(i).Visible = False
Next
End Sub

Public Property Get 角色圖片() As String
   角色圖片 = m_cardmain_jpg
End Property
Public Property Let 角色圖片(ByVal New_角色圖片 As String)
   m_cardmain_jpg = New_角色圖片
   PropertyChanged "角色圖片"
   If m_cardmain_jpg <> "" Then
       cardImage.LoadImage_FromFile m_cardmain_jpg
   End If
End Property
Public Property Get 角色HP() As Integer
   角色HP = m_cardmain_personhp
End Property
Public Property Let 角色HP(ByVal New_角色HP As Integer)
   m_cardmain_personhp = New_角色HP
   PropertyChanged "角色HP"
   '========================
   If m_cardmain_personhp = -99 Then
       personlabelhp.Caption = "?"
   Else
       personlabelhp.Caption = m_cardmain_personhp
   End If
   If m_cardmain_personhp = m_cardmain_personhpmax Or m_cardmain_personhp = -99 Then
        personlabelhp.ForeColor = RGB(255, 255, 255)
        personlabelhp.ForeColor = RGB(255, 255, 255)
        cardback.Opacity = 0
   ElseIf m_cardmain_personhp < m_cardmain_personhpmax And m_cardmain_personhp > m_cardmain_personhp41 Then
        personlabelhp.ForeColor = RGB(255, 255, 128)
        personlabelhp.ForeColor = RGB(255, 255, 128)
        cardback.Opacity = 0
   ElseIf m_cardmain_personhp <= m_cardmain_personhp41 Then
        personlabelhp.ForeColor = RGB(255, 0, 0)
        personlabelhp.ForeColor = RGB(255, 0, 0)
        cardback.Opacity = 0
   End If
   If m_cardmain_personhp = 0 Then
        cardback.Opacity = 100
        cardback.ZOrder
        cardbackclick.Visible = False
   End If
End Property
Public Property Get 角色HPMAX() As Integer
   角色HPMAX = m_cardmain_personhpmax
End Property
Public Property Let 角色HPMAX(ByVal New_角色HPMAX As Integer)
   m_cardmain_personhpmax = New_角色HPMAX
   PropertyChanged "角色HPMAX"
   If m_cardmain_personhpmax < 0 Then m_cardmain_personhpmax = 0
   If m_cardmain_personhpmax > 1 Then
       m_cardmain_personhp41 = Int(m_cardmain_personhpmax / 3 + 0.9)
   Else
       m_cardmain_personhp41 = 0
   End If
   Me.角色HP = m_cardmain_personhp '更改HP狀態
End Property
Public Property Get 角色ATK() As Integer
   角色ATK = m_cardmain_personatk
End Property
Public Property Let 角色ATK(ByVal New_角色ATK As Integer)
   m_cardmain_personatk = New_角色ATK
   PropertyChanged "角色ATK"
   If m_cardmain_personatk = -99 Then
       personlabelatk.Caption = "?"
   ElseIf m_cardmain_personatk < 0 Then
       m_cardmain_personatk = 0
       personlabelatk.Caption = m_cardmain_personatk
   Else
       personlabelatk.Caption = m_cardmain_personatk
   End If
End Property
Public Property Get 角色DEF() As Integer
   角色DEF = m_cardmain_persondef
End Property
Public Property Let 角色DEF(ByVal New_角色DEF As Integer)
   m_cardmain_persondef = New_角色DEF
   PropertyChanged "角色DEF"
   If m_cardmain_persondef = -99 Then
       personlabeldef.Caption = "?"
   ElseIf m_cardmain_persondef < 0 Then
       m_cardmain_persondef = 0
       personlabeldef.Caption = m_cardmain_persondef
   Else
       personlabeldef.Caption = m_cardmain_persondef
   End If
End Property
Public Property Get 是否為新樣式資訊() As Boolean
   是否為新樣式資訊 = m_cardmain_isnewtype
End Property
Public Property Let 是否為新樣式資訊(ByVal New_是否為新樣式資訊 As Boolean)
   m_cardmain_isnewtype = New_是否為新樣式資訊
   PropertyChanged "是否為新樣式資訊"
   If m_cardmain_isnewtype = False Then
        personlabelhp.Left = 555
        personlabelhp.Top = 3240
        personlabelatk.Left = 1200
        personlabelatk.Top = 3240
        personlabeldef.Left = 1920
        personlabeldef.Top = 3240
   Else
        personlabelhp.Left = 300
        personlabelhp.Top = 3220
        personlabelatk.Left = 960
        personlabelatk.Top = 3220
        personlabeldef.Left = 1820
        personlabeldef.Top = 3220
   End If
End Property

Private Sub cardback_Click(ByVal Button As Integer)
RaiseEvent CardClick
End Sub

Private Sub cardback_MouseEnter()
cardbackclick.Visible = True
cardbackclick.ZOrder
End Sub

Private Sub cardback_MouseExit()
cardbackclick.Visible = False
End Sub

Private Sub cardbackclick_Click(ByVal Button As Integer)
RaiseEvent CardClick
End Sub

Private Sub UserControl_Show()
If personlabelhp.FontName <> "Bradley Gratis" Then
    personlabelhp.FontSize = 14
    personlabelatk.FontSize = 14
    personlabeldef.FontSize = 14
End If
End Sub
