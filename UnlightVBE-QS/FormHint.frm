VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form FormHint 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   2805
   ClientLeft      =   3360
   ClientTop       =   4395
   ClientWidth     =   9120
   Icon            =   "FormHint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   9120
   StartUpPosition =   1  '���ݵ�������
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      Begin ImageX.aicAlphaImage imageBlau 
         Height          =   7290
         Left            =   -240
         Top             =   0
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   12859
         Image           =   "FormHint.frx":0CCA
         Props           =   5
      End
   End
   Begin VB.Label bnet 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "��^"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�y���@�U�A�j�p�j�C�z�٨S�������]�w�ڡC"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '���z��
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   11295
   End
   Begin VB.Image bne 
      Height          =   615
      Left            =   7440
      Picture         =   "FormHint.frx":121A2
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1470
   End
End
Attribute VB_Name = "FormHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub bne_Click()
FormHint.Visible = False
�@��t����.���ļ��� 11
���ϥΪ̨ƥ� = True
���q���ƥ� = True
FormMainMode.PEGameFreeModeSettingForm.Enabled = True
End Sub

Private Sub bnet_Click()
FormHint.Visible = False
�@��t����.���ļ��� 11
���ϥΪ̨ƥ� = True
���q���ƥ� = True
FormMainMode.PEGameFreeModeSettingForm.Enabled = True
End Sub

