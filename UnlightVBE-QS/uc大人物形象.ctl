VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl �j�H���ι� 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
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
      Image           =   "uc�j�H���ζH.ctx":0000
      Scaler          =   3
   End
End
Attribute VB_Name = "�j�H���ι�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_smallimage As String
Dim m_bighei As Integer
Dim m_bigwh As Integer
Dim m_bigreturn As Boolean
Public Property Get �j�H���Ϥ�() As String
   �j�H���Ϥ� = m_smallimage
End Property
Public Property Get �j�H���Ϥ�height() As Integer
   �j�H���Ϥ�height = m_bighei
End Property
Public Property Get �j�H���Ϥ�width() As Integer
   �j�H���Ϥ�width = m_bigwh
End Property
Public Property Let �j�H���Ϥ�(ByVal New_�j�H���Ϥ� As String)
   m_smallimage = New_�j�H���Ϥ�
   PropertyChanged "�j�H���Ϥ�"
   If Me.�j�H���Ϥ� <> "" Then
       image1.AutoSize = True
       image1.AutoRedraw = True
       image1.LoadImage_FromFile Me.�j�H���Ϥ�
       image1.Top = 0
       image1.Left = 0
    End If
    Me.�j�H���Ϥ�height = image1.Height
    Me.�j�H���Ϥ�width = image1.Width
End Property
Public Property Let �j�H���Ϥ�height(ByVal New_�j�H���Ϥ�height As Integer)
   m_bighei = New_�j�H���Ϥ�height
   PropertyChanged "�j�H���Ϥ�height"
End Property
Public Property Let �j�H���Ϥ�width(ByVal New_�j�H���Ϥ�width As Integer)
   m_bigwh = New_�j�H���Ϥ�width
   PropertyChanged "�j�H���Ϥ�width"
End Property
Public Property Get �j�H���v������() As Boolean
   �j�H���v������ = m_bigreturn
End Property
Public Property Let �j�H���v������(ByVal New_�j�H���v������ As Boolean)
   m_bigreturn = New_�j�H���v������
   PropertyChanged "�j�H���v������"
   '=====================
   If Me.�j�H���v������ = True Then
       image1.Mirror = aiMirrorHorizontal
   Else
       image1.Mirror = aiMirrorNone
   End If
End Property
