VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "����UnlightVBE-QS"
   ClientHeight    =   5055
   ClientLeft      =   5865
   ClientTop       =   3945
   ClientWidth     =   8400
   Icon            =   "b.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8400
   StartUpPosition =   1  '���ݵ�������
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   8415
      TabIndex        =   5
      Top             =   0
      Width           =   8415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '����
         BackColor       =   &H00000000&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   1200
         Picture         =   "b.frx":0CCA
         ScaleHeight     =   1575
         ScaleWidth      =   7815
         TabIndex        =   6
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�T�w"
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
      Left            =   6840
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   6600
      Picture         =   "b.frx":95E8
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "2019 By Andy Ciu."
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label aboutvn 
      Caption         =   "Bulid 0966"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Version Origin"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "UnlightVBE Type QS"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Visible = False
End Sub


Private Sub Form_Load()
'=====�H�U���]�p�������
'Label5.Visible = False
'Label7.Visible = False

End Sub

Private Sub Image2_Click()
Form2.Visible = False
End Sub

Private Sub Label6_Click()
Form2.Visible = False
End Sub
