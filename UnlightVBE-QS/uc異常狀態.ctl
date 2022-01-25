VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc異常狀態 
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  '透明
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ScaleHeight     =   1950
   ScaleWidth      =   3255
   Windowless      =   -1  'True
   Begin VB.Label personturn 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   25
      Width           =   285
   End
   Begin VB.Label personnum 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   340
      TabIndex        =   0
      Top             =   25
      Width           =   285
   End
   Begin ImageX.aicAlphaImage personimg 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Image           =   "uc異常狀態.ctx":0000
      Scaler          =   2
      Props           =   17
      MaskColor       =   16777215
   End
   Begin ImageX.aicAlphaImage aie2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Image           =   "uc異常狀態.ctx":0119
      Scaler          =   1
      Props           =   25
      MaskColor       =   16777215
   End
End
Attribute VB_Name = "uc異常狀態"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_personimg As String
Dim m_personnum As Integer
Dim m_personturn As Integer
Public Property Get 異常狀態圖片() As String
   異常狀態圖片 = m_personimg
End Property
Public Property Get person_num() As Integer
   person_num = m_personnum
End Property
Public Property Get person_turn() As Integer
   person_turn = m_personturn
End Property
Public Property Let 異常狀態圖片(ByVal New_異常狀態圖片 As String)
   m_personimg = New_異常狀態圖片
   PropertyChanged "異常狀態圖片"
   If Me.異常狀態圖片 <> "" Then
       personimg.LoadImage_FromFile Me.異常狀態圖片
    End If
End Property
Public Property Let person_num(ByVal New_person_num As Integer)
   m_personnum = New_person_num
   PropertyChanged "person_num"
   personnum.Caption = m_personnum
   If m_personnum = 0 Then
       personnum.Visible = False
   Else
       personnum.Visible = True
   End If
End Property
Public Property Let person_turn(ByVal New_person_turn As Integer)
   m_personturn = New_person_turn
   PropertyChanged "person_turn"
   personturn.Caption = m_personturn
End Property

