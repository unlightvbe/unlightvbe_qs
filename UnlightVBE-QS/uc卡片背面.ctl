VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc�d���I�� 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
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
      Image           =   "uc�d���I��.ctx":0000
      Scaler          =   3
      Props           =   13
   End
End
Attribute VB_Name = "uc�d���I��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_invent As Integer
Dim m_picture As String
Dim m_num As Integer
Public Property Get �������O() As Integer
   �������O = m_invent
End Property
Public Property Let �������O(ByVal New_�������O As Integer)
   m_invent = New_�������O
   PropertyChanged "�������O"
End Property
Public Property Get �Ϥ�() As String
   �Ϥ� = m_picture
End Property
Public Property Let �Ϥ�(ByVal New_�Ϥ� As String)
   m_picture = New_�Ϥ�
   PropertyChanged "�Ϥ�"
   '============
   image1.LoadImage_FromFile Me.�Ϥ�
   image1.Left = 0
   image1.Top = 0
End Property
Public Property Get ���ؽs��() As Integer
   ���ؽs�� = m_num
End Property
Public Property Let ���ؽs��(ByVal New_���ؽs�� As Integer)
   m_num = New_���ؽs��
   PropertyChanged "���ؽs��"
   '==============
   Select Case Me.�������O
        Case 1
             image1.Left = Val(Me.���ؽs��) * -315
             image1.Top = 0
        Case 2
             image1.Left = (Val(Me.���ؽs��) - 1) * -120
             image1.Top = 0
        Case 3
             image1.Left = 0
             image1.Top = (Val(Me.���ؽs��) - 1) * -210
    End Select
End Property

