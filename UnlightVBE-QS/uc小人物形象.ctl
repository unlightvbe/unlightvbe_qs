VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl �p�H���ζH 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ClipBehavior    =   0  '�L
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
      Image           =   "uc�p�H���ζH.ctx":0000
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image2 
      Height          =   855
      Left            =   120
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      Image           =   "uc�p�H���ζH.ctx":0018
      Scaler          =   3
   End
End
Attribute VB_Name = "�p�H���ζH"
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

Public Property Get �p�H���v�lLeft() As Integer
   �p�H���v�lLeft = m_smalldownleft
End Property
Public Property Get �p�H���v�ltop�t() As Integer
   �p�H���v�ltop�t = m_smalldowntop
End Property
Public Property Get �p�H���Ϥ�() As String
   �p�H���Ϥ� = m_smallimage
End Property
Public Property Get �p�H���Ϥ�height() As Integer
   �p�H���Ϥ�height = m_smallhei
End Property
Public Property Get �p�H���Ϥ�width() As Integer
   �p�H���Ϥ�width = m_smallwh
End Property
Public Property Get �p�H���v�l�Ϥ�() As String
   �p�H���v�l�Ϥ� = m_smallimagedown
End Property
Public Property Let �p�H���v�lLeft(ByVal New_�p�H���v�lLeft As Integer)
   m_smalldownleft = New_�p�H���v�lLeft
   PropertyChanged "�p�H���v�lLeft"
   Image2.Left = Me.�p�H���v�lLeft
End Property
Public Property Let �p�H���v�ltop�t(ByVal New_�p�H���v�ltop�t As Integer)
   m_smalldowntop = New_�p�H���v�ltop�t
   PropertyChanged "�p�H���v�ltop�t"
   Image2.Top = image1.Height + Me.�p�H���v�ltop�t
End Property
Public Property Let �p�H���Ϥ�(ByVal New_�p�H���Ϥ� As String)
   m_smallimage = New_�p�H���Ϥ�
   PropertyChanged "�p�H���Ϥ�"
   If Me.�p�H���Ϥ� <> "" Then
       image1.AutoRedraw = True
       image1.AutoSize = True
       image1.LoadImage_FromFile Me.�p�H���Ϥ�
       image1.Left = 0
       image1.Top = 0
       Me.�p�H���Ϥ�height = image1.Height
       Me.�p�H���Ϥ�width = image1.Width
       Me.�p�H������ = False
       Me.�p�H����{ = False
   End If
End Property
Public Property Let �p�H���v�l�Ϥ�(ByVal New_�p�H���v�l�Ϥ� As String)
   m_smallimagedown = New_�p�H���v�l�Ϥ�
   PropertyChanged "�p�H���v�l�Ϥ�"
   If Me.�p�H���v�l�Ϥ� <> "" Then
       Image2.AutoRedraw = True
       Image2.AutoSize = True
       Image2.LoadImage_FromFile Me.�p�H���v�l�Ϥ�
       Image2.Left = 0
       Image2.Top = image1.Height
       Image2.Opacity = 100
   End If
End Property
Public Property Let �p�H���Ϥ�height(ByVal New_�p�H���Ϥ�height As Integer)
   m_smallhei = New_�p�H���Ϥ�height
   PropertyChanged "�p�H���Ϥ�height"
End Property
Public Property Let �p�H���Ϥ�width(ByVal New_�p�H���Ϥ�width As Integer)
   m_smallwh = New_�p�H���Ϥ�width
   PropertyChanged "�p�H���Ϥ�width"
End Property
Public Property Get �p�H������() As Boolean
   �p�H������ = m_smalldeath
End Property
Public Property Let �p�H������(ByVal New_�p�H������ As Boolean)
   m_smalldeath = New_�p�H������
   PropertyChanged "�p�H������"
   '=====================
   If Me.�p�H������ = True Then
       t1.Enabled = True
   End If
End Property
Public Property Get �p�H�����]() As Boolean
   �p�H�����] = m_smallreset
End Property
Public Property Let �p�H�����](ByVal New_�p�H�����] As Boolean)
   m_smallreset = New_�p�H�����]
   PropertyChanged "�p�H�����]"
   '=====================
   If Me.�p�H�����] = True Then
       image1.Opacity = 100
       Image2.Opacity = 100
       Me.�p�H���v������ = False
       Me.�p�H������ = False
       Me.�p�H����{ = False
       Me.�p�H�����] = False
   End If
End Property
Public Property Get �p�H����{() As Boolean
   �p�H����{ = m_smalllive
End Property
Public Property Let �p�H����{(ByVal New_�p�H����{ As Boolean)
   m_smalllive = New_�p�H����{
   PropertyChanged "�p�H����{"
   '=====================
   If Me.�p�H����{ = True Then
       t2.Enabled = True
   End If
End Property
Public Property Get �p�H���v������() As Boolean
   �p�H���v������ = m_smallreturn
End Property
Public Property Let �p�H���v������(ByVal New_�p�H���v������ As Boolean)
   m_smallreturn = New_�p�H���v������
   PropertyChanged "�p�H���v������"
   '=====================
   If Me.�p�H���v������ = True Then
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
    Me.�p�H������ = False
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
    Me.�p�H����{ = False
End If
End Sub
