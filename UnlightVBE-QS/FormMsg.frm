VERSION 5.00
Begin VB.Form FormMessage 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "UnlightVBE"
   ClientHeight    =   8370
   ClientLeft      =   3360
   ClientTop       =   4395
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "�L�n������"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9120
   StartUpPosition =   2  '�ù�����
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      Picture         =   "FormMsg.frx":0CCA
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   5655
      Left            =   9000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FormMsg.frx":3C5E
      Top             =   2040
      Width           =   8175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "Label2"
      Height          =   5655
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   8655
   End
   Begin VB.Label bnet 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�T�w"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "���j�p�j���@�h�q��"
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
      Height          =   1725
      Left            =   0
      Top             =   0
      Width           =   11295
   End
   Begin VB.Image bne 
      Height          =   615
      Left            =   7440
      Picture         =   "FormMsg.frx":3E61
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   1470
   End
End
Attribute VB_Name = "FormMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub bne_Click()
FormMessage.Visible = False
End Sub

Private Sub bnet_Click()
FormMessage.Visible = False
End Sub

