VERSION 5.00
Begin VB.Form FormDevSetting 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE-影子設定"
   ClientHeight    =   6375
   ClientLeft      =   9135
   ClientTop       =   3540
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDevSetting.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  '像素
   ScaleWidth      =   432
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      Caption         =   "顯示列"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   40
      Top             =   5400
      Width           =   6495
      Begin VB.CommandButton Command3 
         Caption         =   "時間軸開始"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   50
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "時間軸停止"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         TabIndex        =   49
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton personfcomright 
         Caption         =   "R"
         Height          =   375
         Left            =   4080
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton personfcomleft 
         Caption         =   "L"
         Height          =   375
         Left            =   3600
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton personfusright 
         Caption         =   "R"
         Height          =   375
         Left            =   2520
         TabIndex        =   46
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton personfusleft 
         Caption         =   "L"
         Height          =   375
         Left            =   2040
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.Label personfcom 
         Height          =   375
         Left            =   5520
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '透明
         Caption         =   "右Left:"
         Height          =   375
         Left            =   4680
         TabIndex        =   43
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label personfus 
         Height          =   375
         Left            =   840
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "左Left:"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '沒有框線
      Height          =   2655
      Left            =   3240
      ScaleHeight     =   2655
      ScaleWidth      =   3255
      TabIndex        =   30
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton smallpntdncom 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton smallpntupcom 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   33
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallpnrcom 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   32
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallpnlcom 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   31
         Top             =   960
         Width           =   495
      End
      Begin VB.Label smallpntopcom 
         Height          =   375
         Left            =   1440
         TabIndex        =   39
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label smallpnleftcom 
         Height          =   375
         Left            =   1560
         TabIndex        =   38
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   37
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "COM(人物)"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '沒有框線
      Height          =   2655
      Left            =   -120
      ScaleHeight     =   2655
      ScaleWidth      =   3255
      TabIndex        =   20
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton smallpnlus 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallpnrus 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   23
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallpntupus 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallpntdnus 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label smallpnleftus 
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label smallpntopus 
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "1P(人物)"
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '沒有框線
      Height          =   2655
      Left            =   3240
      ScaleHeight     =   2655
      ScaleWidth      =   3255
      TabIndex        =   9
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton smallclcom 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallcrcom 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallctupcom 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallctdncom 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "COM(影子)"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label smallleftcom 
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label smalltopcom 
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '沒有框線
      Height          =   2655
      Left            =   -120
      ScaleHeight     =   2655
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton smallctdnus 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton smallctupus 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallcrus 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallclus 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "1P(影子)"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label smalltopus 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label smallleftus 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   1680
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormDevSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
戰鬥系統類.時間軸_停止
End Sub

Private Sub Command3_Click()
FormMainMode.trtimeline.Enabled = True
End Sub

Private Sub personfcomleft_Click()
FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.顯示列1.電腦方小人物圖片left - 10
personfcom.Caption = personfcom.Caption - 10
End Sub

Private Sub personfcomright_Click()
FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.顯示列1.電腦方小人物圖片left + 10
personfcom.Caption = personfcom.Caption + 10
End Sub

Private Sub personfusleft_Click()
FormMainMode.顯示列1.使用者方小人物圖片left = FormMainMode.顯示列1.使用者方小人物圖片left - 10
personfus.Caption = personfus.Caption - 10
End Sub

Private Sub personfusright_Click()
FormMainMode.顯示列1.使用者方小人物圖片left = FormMainMode.顯示列1.使用者方小人物圖片left + 10
personfus.Caption = personfus.Caption + 10
End Sub

Private Sub smallclcom_Click()
FormMainMode.personcomminijpg.小人物影子Left = Val(FormMainMode.personcomminijpg.小人物影子Left) - 10
smallleftcom.Caption = Val(smallleftcom.Caption) - 10
End Sub

Private Sub smallclus_Click()
FormMainMode.personusminijpg.小人物影子Left = Val(FormMainMode.personusminijpg.小人物影子Left) - 10
smallleftus.Caption = Val(smallleftus.Caption) - 10
End Sub

Private Sub smallcrcom_Click()
FormMainMode.personcomminijpg.小人物影子Left = Val(FormMainMode.personcomminijpg.小人物影子Left) + 10
smallleftcom.Caption = Val(smallleftcom.Caption) + 10
End Sub

Private Sub smallcrus_Click()
FormMainMode.personusminijpg.小人物影子Left = Val(FormMainMode.personusminijpg.小人物影子Left) + 10
smallleftus.Caption = Val(smallleftus.Caption) + 10
End Sub

Private Sub smallctdncom_Click()
FormMainMode.personcomminijpg.小人物影子top差 = Val(FormMainMode.personcomminijpg.小人物影子top差) + 10
smalltopcom.Caption = Val(smalltopcom.Caption) + 10
End Sub

Private Sub smallctdnus_Click()
FormMainMode.personusminijpg.小人物影子top差 = Val(FormMainMode.personusminijpg.小人物影子top差) + 10
smalltopus.Caption = Val(smalltopus.Caption) + 10
End Sub

Private Sub smallctupcom_Click()
FormMainMode.personcomminijpg.小人物影子top差 = Val(FormMainMode.personcomminijpg.小人物影子top差) - 10
smalltopcom.Caption = Val(smalltopcom.Caption) - 10
End Sub

Private Sub smallctupus_Click()
FormMainMode.personusminijpg.小人物影子top差 = Val(FormMainMode.personusminijpg.小人物影子top差) - 10
smalltopus.Caption = Val(smalltopus.Caption) - 10
End Sub

Private Sub smallpnlcom_Click()
FormMainMode.personcomminijpg.Left = Val(FormMainMode.personcomminijpg.Left) - 10
smallpnleftcom.Caption = Val(smallpnleftcom.Caption) - 10
End Sub



Private Sub smallpnlus_Click()
FormMainMode.personusminijpg.Left = Val(FormMainMode.personusminijpg.Left) - 10
smallpnleftus.Caption = Val(smallpnleftus.Caption) - 10
End Sub

Private Sub smallpnrcom_Click()
FormMainMode.personcomminijpg.Left = Val(FormMainMode.personcomminijpg.Left) + 10
smallpnleftcom.Caption = Val(smallpnleftcom.Caption) + 10
End Sub

Private Sub smallpnrus_Click()
FormMainMode.personusminijpg.Left = Val(FormMainMode.personusminijpg.Left) + 10
smallpnleftus.Caption = Val(smallpnleftus.Caption) + 10
End Sub

Private Sub smallpntdncom_Click()
FormMainMode.personcomminijpg.Top = Val(FormMainMode.personcomminijpg.Top) + 10
smallpntopcom.Caption = Val(smallpntopcom.Caption) + 10
End Sub

Private Sub smallpntdnus_Click()
FormMainMode.personusminijpg.Top = Val(FormMainMode.personusminijpg.Top) + 10
smallpntopus.Caption = Val(smallpntopus.Caption) + 10
End Sub

Private Sub smallpntupcom_Click()
FormMainMode.personcomminijpg.Top = Val(FormMainMode.personcomminijpg.Top) - 10
smallpntopcom.Caption = Val(smallpntopcom.Caption) - 10
End Sub

Private Sub smallpntupus_Click()
FormMainMode.personusminijpg.Top = Val(FormMainMode.personusminijpg.Top) - 10
smallpntopus.Caption = Val(smallpntopus.Caption) - 10
End Sub
