VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc卡片背面 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   ScaleHeight     =   2775
   ScaleWidth      =   5295
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage image1 
      Height          =   165
      Left            =   0
      Top             =   0
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   291
      Image           =   "uc卡片背面.ctx":0000
      Scaler          =   3
      Props           =   13
   End
End
Attribute VB_Name = "uc卡片背面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_invent As Integer
Dim m_picture As String
Dim m_num As Integer
Public Property Get 物件類別() As Integer
   物件類別 = m_invent
End Property
Public Property Let 物件類別(ByVal New_物件類別 As Integer)
   m_invent = New_物件類別
   PropertyChanged "物件類別"
End Property
Public Property Get 圖片() As String
   圖片 = m_picture
End Property
Public Property Let 圖片(ByVal New_圖片 As String)
   m_picture = New_圖片
   PropertyChanged "圖片"
   '============
   Image1.LoadImage_FromFile Me.圖片
   Image1.Left = 0
   Image1.Top = 0
End Property
Public Property Get 項目編號() As Integer
   項目編號 = m_num
End Property
Public Property Let 項目編號(ByVal New_項目編號 As Integer)
   m_num = New_項目編號
   PropertyChanged "項目編號"
   '==============
   Select Case Me.物件類別
        Case 1
             Image1.Left = Val(Me.項目編號) * -300
             Image1.Top = 0
        Case 2
             Image1.Left = (Val(Me.項目編號) - 1) * -120
             Image1.Top = 0
        Case 3
             Image1.Left = 0
             Image1.Top = (Val(Me.項目編號) - 1) * -210
    End Select
End Property

