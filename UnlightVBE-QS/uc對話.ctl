VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc對話 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4695
   ScaleWidth      =   7815
   Windowless      =   -1  'True
   Begin VB.Label talktext 
      BackStyle       =   0  '透明
      Caption         =   "為劍而生的東西也將為劍而死。僅此而已。"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4560
      WordWrap        =   -1  'True
   End
   Begin ImageX.aicAlphaImage image1 
      Height          =   1680
      Left            =   240
      Top             =   240
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   2963
      Image           =   "uc對話.ctx":0000
      Props           =   13
   End
End
Attribute VB_Name = "uc對話"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_persontalkstr As String
Dim m_persontalkyn As Boolean
Public Property Get 對話文字() As String
   對話文字 = m_persontalkstr
End Property
Public Property Let 對話文字(ByVal New_對話文字 As String)
   m_persontalkstr = New_對話文字
   PropertyChanged "對話文字"
   '================
   talktext.Caption = Me.對話文字
End Property
Public Property Get 對話文字顯示() As Boolean
   對話文字顯示 = m_persontalkyn
End Property
Public Property Let 對話文字顯示(ByVal New_對話文字顯示 As Boolean)
   m_persontalkyn = New_對話文字顯示
   PropertyChanged "對話文字顯示"
   '================
   talktext.Visible = Me.對話文字顯示
End Property

