VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc��� 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "�L�n������"
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
      BackStyle       =   0  '�z��
      Caption         =   "���C�ӥͪ��F��]�N���C�Ӧ��C�Ȧ��Ӥw�C"
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
      Image           =   "uc���.ctx":0000
      Props           =   13
   End
End
Attribute VB_Name = "uc���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_persontalkstr As String
Dim m_persontalkyn As Boolean
Public Property Get ��ܤ�r() As String
   ��ܤ�r = m_persontalkstr
End Property
Public Property Let ��ܤ�r(ByVal New_��ܤ�r As String)
   m_persontalkstr = New_��ܤ�r
   PropertyChanged "��ܤ�r"
   '================
   talktext.Caption = Me.��ܤ�r
End Property
Public Property Get ��ܤ�r���() As Boolean
   ��ܤ�r��� = m_persontalkyn
End Property
Public Property Let ��ܤ�r���(ByVal New_��ܤ�r��� As Boolean)
   m_persontalkyn = New_��ܤ�r���
   PropertyChanged "��ܤ�r���"
   '================
   talktext.Visible = Me.��ܤ�r���
End Property

