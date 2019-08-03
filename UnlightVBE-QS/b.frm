VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  '單線固定
   Caption         =   "關於UnlightVBE-QS"
   ClientHeight    =   5535
   ClientLeft      =   5865
   ClientTop       =   3945
   ClientWidth     =   8400
   Icon            =   "b.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8400
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   8415
      TabIndex        =   5
      Top             =   0
      Width           =   8415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         BorderStyle     =   0  '沒有框線
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
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "(CC BY-ND 4.0) CPA Co.,Ltd."
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "本程式內使用之Unlight相關素材，皆授權自CPA。"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
   End
   Begin ImageX.aicAlphaImage aicAlphaImage1 
      Height          =   465
      Left            =   6960
      Top             =   2280
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   820
      Image           =   "b.frx":95E8
      Props           =   5
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   6720
      Picture         =   "b.frx":9C22
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "2019 By Andy Ciu."
      BeginProperty Font 
         Name            =   "微軟正黑體"
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
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label aboutvn 
      Caption         =   "Bulid 0967"
      BeginProperty Font 
         Name            =   "微軟正黑體"
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
         Name            =   "微軟正黑體"
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
         Name            =   "微軟正黑體"
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
'=====以下為設計物件顯示
'Label5.Visible = False
'Label7.Visible = False

End Sub

Private Sub Image2_Click()
Form2.Visible = False
End Sub

Private Sub Label6_Click()
Form2.Visible = False
End Sub
