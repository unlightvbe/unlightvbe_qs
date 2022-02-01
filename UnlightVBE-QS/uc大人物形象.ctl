VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl 大人物形像 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   ScaleHeight     =   9405
   ScaleWidth      =   9675
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage image1 
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   16113
      Image           =   "uc大人物形象.ctx":0000
      Scaler          =   3
   End
End
Attribute VB_Name = "大人物形像"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_smallimage As String
Dim m_bighei As Integer
Dim m_bigwh As Integer
Dim m_bigreturn As Boolean
Public Property Get 大人物圖片() As String
   大人物圖片 = m_smallimage
End Property
Public Property Get 大人物圖片height() As Integer
   大人物圖片height = m_bighei
End Property
Public Property Get 大人物圖片width() As Integer
   大人物圖片width = m_bigwh
End Property
Public Property Let 大人物圖片(ByVal New_大人物圖片 As String)
   m_smallimage = New_大人物圖片
   PropertyChanged "大人物圖片"
   If Me.大人物圖片 <> "" Then
       image1.AutoSize = True
       image1.AutoRedraw = True
       image1.LoadImage_FromFile Me.大人物圖片
       image1.Top = 0
       image1.Left = 0
    End If
    Me.大人物圖片height = image1.Height
    Me.大人物圖片width = image1.Width
End Property
Public Property Let 大人物圖片height(ByVal New_大人物圖片height As Integer)
   m_bighei = New_大人物圖片height
   PropertyChanged "大人物圖片height"
End Property
Public Property Let 大人物圖片width(ByVal New_大人物圖片width As Integer)
   m_bigwh = New_大人物圖片width
   PropertyChanged "大人物圖片width"
End Property
Public Property Get 大人物影像反轉() As Boolean
   大人物影像反轉 = m_bigreturn
End Property
Public Property Let 大人物影像反轉(ByVal New_大人物影像反轉 As Boolean)
   m_bigreturn = New_大人物影像反轉
   PropertyChanged "大人物影像反轉"
   '=====================
   If Me.大人物影像反轉 = True Then
       image1.Mirror = aiMirrorHorizontal
   Else
       image1.Mirror = aiMirrorNone
   End If
End Property
