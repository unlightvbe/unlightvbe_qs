VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc技能視窗 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11490
   ScaleHeight     =   10500
   ScaleWidth      =   11490
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage Image1 
      Height          =   10155
      Left            =   -120
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   17912
      Image           =   "uc技能視窗.ctx":0000
   End
End
Attribute VB_Name = "uc技能視窗"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim m_smallimage As String
Public Property Get 技能圖片() As String
   技能圖片 = m_smallimage
End Property
'Public Property Get 大人物圖片height() As Integer
'   大人物圖片height = m_bighei
'End Property
'Public Property Get 大人物圖片width() As Integer
'   大人物圖片width = m_bigwh
'End Property
Public Property Let 技能圖片(ByVal New_技能圖片 As String)
   m_smallimage = New_技能圖片
   PropertyChanged "技能圖片"
   If Me.技能圖片 <> "" Then
       Image1.LoadImage_FromFile Me.技能圖片
       Image1.Top = 0
       Image1.Left = 0
    End If
'    Me.技能圖片height = Image1.Height
'    Me.技能圖片width = Image1.Width
End Property
'Public Property Let 大人物圖片height(ByVal New_大人物圖片height As Integer)
'   m_bighei = New_大人物圖片height
'   PropertyChanged "大人物圖片height"
'End Property
'Public Property Let 大人物圖片width(ByVal New_大人物圖片width As Integer)
'   m_bigwh = New_大人物圖片width
'   PropertyChanged "大人物圖片width"
'End Property
Private Sub UserControl_Initialize()

End Sub
