VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form FormMainMode 
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE-QS Origin"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20400
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMainMode.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   20400
   StartUpPosition =   2  '螢幕中央
   Tag             =   "UnlightVBE-QS Origin"
   Begin VB.PictureBox PEAttackingForm 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9910
      Left            =   4800
      Picture         =   "FormMainMode.frx":0CCA
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   11340
      Begin VB.CommandButton 影子設定 
         Caption         =   "影子設定"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8760
         TabIndex        =   125
         Top             =   9360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "測試"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   124
         Top             =   9360
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "離開"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   123
         Top             =   9360
         Width           =   1215
      End
      Begin VB.PictureBox PEAFtoolbox 
         Height          =   2175
         Left            =   4920
         ScaleHeight     =   2115
         ScaleWidth      =   6075
         TabIndex        =   111
         Top             =   6240
         Visible         =   0   'False
         Width           =   6135
         Begin VB.Timer 攻擊階段_階段初始 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   4680
            Top             =   240
         End
         Begin VB.Timer 移動階段_階段初始 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   1920
            Top             =   1200
         End
         Begin VB.Timer 防禦階段_階段初始 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   4680
            Top             =   840
         End
         Begin VB.Timer NextTurn_階段2 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   3840
            Top             =   1200
         End
         Begin VB.CommandButton cn1 
            Caption         =   "發牌"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   119
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cnmove2 
            Caption         =   "OK"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   118
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cnmove 
            Caption         =   "下一步"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   117
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn32 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   116
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn22 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   115
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn3 
            Caption         =   "下一步"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   114
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn2 
            Caption         =   "下一步"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   113
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn4 
            Caption         =   "Next Turn"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   112
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Timer OK按鈕牌完成移動檢查 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   1200
            Top             =   0
         End
         Begin VB.Timer 對齊完成檢查 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   600
            Top             =   0
         End
         Begin VB.Timer 攻擊階段_階段1 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   5040
            Top             =   240
         End
         Begin VB.Timer 攻擊階段_階段2 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   5400
            Top             =   240
         End
         Begin VB.Timer 使用者出牌_手牌對齊 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   0
            Top             =   120
         End
      End
      Begin MSScriptControlCtl.ScriptControl PEAFvssc 
         Index           =   0
         Left            =   2640
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         UseSafeSubset   =   -1  'True
      End
      Begin VB.Timer 使用者出牌_AI出牌控制_事件卡 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3720
         Top             =   5640
      End
      Begin VB.Timer 使用者出牌_AI出牌控制 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3240
         Top             =   5640
      End
      Begin VB.Timer 人物消失檢查 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   2400
         Top             =   2640
      End
      Begin VB.Timer tr牌組_回牌_電腦 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1320
         Top             =   3840
      End
      Begin VB.Timer atkingtrus 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   10920
         Top             =   5400
      End
      Begin VB.Timer trend 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   10920
         Top             =   4920
      End
      Begin VB.Timer trnextend 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   5040
      End
      Begin VB.Timer 牌移動 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   960
         Top             =   2760
      End
      Begin VB.Timer 發牌_使用者階段 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   480
         Top             =   2520
      End
      Begin VB.Timer 發牌_電腦階段 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   480
         Top             =   3000
      End
      Begin VB.Timer 發牌檢查 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   2760
      End
      Begin VB.Timer 牌移動_收牌 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   1680
      End
      Begin VB.Timer 使用者出牌_出牌對齊_靠左 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   5520
         Top             =   5520
      End
      Begin VB.Timer 使用者出牌_出牌對齊_靠右 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   5760
         Top             =   5520
      End
      Begin VB.Timer atkingtrcom 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   10920
         Top             =   3120
      End
      Begin VB.Timer 電腦出牌 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8280
         Top             =   120
      End
      Begin VB.Timer 電腦出牌_出牌對齊_靠左 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7320
         Top             =   1080
      End
      Begin VB.Timer 電腦出牌_手牌對齊 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7800
         Top             =   120
      End
      Begin VB.Timer 電腦出牌_亮牌 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   7440
         Top             =   1560
      End
      Begin VB.Timer 收牌階段_計算 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   2160
      End
      Begin VB.Timer 骰子執行完啟動 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   720
         Top             =   5040
      End
      Begin VB.Timer 等待時間 
         Enabled         =   0   'False
         Interval        =   375
         Left            =   10920
         Top             =   2640
      End
      Begin VB.Timer 小人物頭像移動_使用者 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3840
         Top             =   1080
      End
      Begin VB.Timer trgoi1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3000
         Top             =   3120
      End
      Begin VB.Timer trgoi2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   8160
         Top             =   3120
      End
      Begin VB.Timer 小人物頭像移動_電腦 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4200
         Top             =   1080
      End
      Begin VB.Timer 移動圖片完成檢查 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1680
         Top             =   1920
      End
      Begin VB.Timer tr電腦牌_翻牌 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   8280
         Top             =   1080
      End
      Begin VB.Timer tr電腦牌_偷牌 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   8280
         Top             =   1560
      End
      Begin VB.Timer tr牌組_回牌_使用者 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1080
         Top             =   3840
      End
      Begin VB.Timer tr使用者_棄牌 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1080
         Top             =   4440
      End
      Begin VB.Timer tr電腦牌_棄牌 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   7920
         Top             =   1560
      End
      Begin VB.Timer tr牌組_抽牌_使用者 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1200
         Top             =   4440
      End
      Begin VB.Timer tr牌組_抽牌_電腦 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1920
         Top             =   1440
      End
      Begin VB.Timer trtimeline 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   5520
         Top             =   4920
      End
      Begin VB.Timer 血量載入動畫 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7440
         Top             =   5640
      End
      Begin VB.Timer 等待時間_2 
         Enabled         =   0   'False
         Interval        =   187
         Left            =   10560
         Top             =   2640
      End
      Begin VB.Timer tr使用者牌_偷牌 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   5520
      End
      Begin VB.Timer 電腦出牌_出牌對齊_靠右 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7560
         Top             =   1080
      End
      Begin UnlightVBE.uc角色卡片介面 cardcom 
         Height          =   3615
         Index           =   0
         Left            =   600
         TabIndex        =   129
         Top             =   6120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   2355
         _ExtentY        =   3625
      End
      Begin UnlightVBE.uc角色卡片介面 cardus 
         Height          =   3615
         Index           =   0
         Left            =   0
         TabIndex        =   128
         Top             =   6120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   2355
         _ExtentY        =   3625
      End
      Begin UnlightVBE.uc技能動畫介面 PEAFAnimateInterface 
         Height          =   9910
         Left            =   0
         Top             =   0
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   17489
      End
      Begin UnlightVBE.uc角色小卡 PEAFpersoncardcom 
         Height          =   495
         Index           =   3
         Left            =   5040
         TabIndex        =   138
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc角色小卡 PEAFpersoncardcom 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   137
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc角色小卡 PEAFpersoncardcom 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc角色小卡 PEAFpersoncardus 
         Height          =   495
         Index           =   3
         Left            =   5040
         TabIndex        =   135
         Top             =   9360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc角色小卡 PEAFpersoncardus 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   134
         Top             =   9360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc角色小卡 PEAFpersoncardus 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   133
         Top             =   9360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc技能說明 PEAFatkinghelpc 
         Height          =   3255
         Left            =   2640
         TabIndex        =   132
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5741
      End
      Begin UnlightVBE.uc戰鬥系統牌型介面 PEAFInterface 
         Height          =   9915
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   11340
         _ExtentX        =   2143
         _ExtentY        =   2778
      End
      Begin VB.Label bloodnumcom2 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11040
         TabIndex        =   2
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label bloodnumcom1 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   10560
         TabIndex        =   3
         Top             =   5600
         Width           =   375
      End
      Begin VB.Image PEAFbloodbackimage2 
         Height          =   690
         Left            =   10080
         Picture         =   "FormMainMode.frx":22BA9
         Top             =   5440
         Width           =   1275
      End
      Begin VB.Label bloodnumus2 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   590
         TabIndex        =   4
         Top             =   5820
         Width           =   300
      End
      Begin VB.Label bloodnumus1 
         Alignment       =   1  '靠右對齊
         BackStyle       =   0  '透明
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   30
         TabIndex        =   5
         Top             =   5600
         Width           =   450
      End
      Begin VB.Image PEAFbloodbackimage1 
         Height          =   690
         Left            =   0
         Picture         =   "FormMainMode.frx":232F0
         Top             =   5440
         Width           =   1290
      End
      Begin UnlightVBE.uc擲骰介面 PEAFDiceInterface 
         Height          =   9910
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Width           =   11340
         _ExtentX        =   2566
         _ExtentY        =   1296
      End
      Begin VB.Label pageusglead 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   122
         Top             =   6600
         Width           =   135
      End
      Begin UnlightVBE.ucCard card 
         Height          =   1335
         Index           =   0
         Left            =   5280
         TabIndex        =   121
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2355
      End
      Begin VB.Label pagecomglead 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   120
         Top             =   720
         Width           =   135
      End
      Begin UnlightVBE.顯示列 顯示列1 
         Height          =   1215
         Left            =   0
         TabIndex        =   1
         Top             =   3520
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2143
      End
      Begin VB.Image atkdef2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":23A80
         Top             =   1860
         Width           =   2280
      End
      Begin VB.Image atkdef1 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":241CA
         Top             =   1590
         Width           =   2280
      End
      Begin VB.Image draw2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":24920
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Image move2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":24F87
         Top             =   1320
         Width           =   2280
      End
      Begin VB.Image PEAFtest 
         Height          =   2865
         Index           =   1
         Left            =   2040
         Picture         =   "FormMainMode.frx":256A9
         Tag             =   "personusminijpg"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label pageusqlead 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   5880
         Width           =   135
      End
      Begin VB.Label pagecomqlead 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   2160
         Width           =   135
      End
      Begin VB.Image Image2 
         Height          =   120
         Left            =   5280
         Picture         =   "FormMainMode.frx":271D7
         Top             =   6120
         Width           =   780
      End
      Begin VB.Shape bloodlineout1 
         BorderStyle     =   0  '透明
         FillColor       =   &H000000FF&
         FillStyle       =   0  '實心
         Height          =   80
         Left            =   0
         Top             =   6160
         Width           =   5295
      End
      Begin VB.Shape bloodlineout2 
         BorderStyle     =   0  '透明
         FillColor       =   &H000000FF&
         FillStyle       =   0  '實心
         Height          =   75
         Left            =   6060
         Top             =   6160
         Width           =   5295
      End
      Begin VB.Image timeup 
         Height          =   105
         Left            =   5290
         Picture         =   "FormMainMode.frx":2726A
         Top             =   4720
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Line timelineout1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   0
         X2              =   5250
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Line timelineout2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   6060
         X2              =   11310
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Image PEAFtest 
         Height          =   2880
         Index           =   2
         Left            =   9960
         Picture         =   "FormMainMode.frx":272D6
         Tag             =   "personcomminijpg"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape bloodlinein1 
         BorderStyle     =   6  '內實線
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  '實心
         Height          =   90
         Left            =   0
         Top             =   6150
         Width           =   5295
      End
      Begin VB.Shape bloodlinein2 
         BorderStyle     =   6  '內實線
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  '實心
         Height          =   90
         Left            =   6060
         Top             =   6150
         Width           =   5295
      End
      Begin VB.Shape timelinein1 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  '內實線
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  '實心
         Height          =   90
         Left            =   0
         Top             =   4720
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Shape timelinein2 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  '內實線
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  '實心
         Height          =   90
         Left            =   6050
         Top             =   4720
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Image draw1 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":28D00
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Image move1 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":28E7F
         Top             =   1340
         Width           =   2040
      End
      Begin VB.Image move3 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":2901B
         Top             =   1610
         Width           =   2040
      End
      Begin VB.Image move4 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":292A4
         Top             =   1880
         Width           =   2040
      End
      Begin UnlightVBE.小人物形象 personusminijpg 
         Height          =   4935
         Left            =   0
         TabIndex        =   9
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8705
      End
      Begin VB.Image cardpagejpg 
         Height          =   915
         Left            =   0
         Picture         =   "FormMainMode.frx":29521
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label pageul 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "57"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   1100
         Width           =   855
      End
      Begin UnlightVBE.小人物形象 personcomminijpg 
         Height          =   4935
         Left            =   5520
         TabIndex        =   10
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8705
      End
      Begin ImageX.aicAlphaImage PEAFMoveRange 
         Height          =   1080
         Left            =   2880
         Top             =   2160
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   1905
         Image           =   "FormMainMode.frx":29D84
         Scaler          =   3
         Props           =   13
      End
   End
   Begin VB.PictureBox PEGameFreeModeSettingForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   2760
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   11340
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '平面
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   11415
         TabIndex        =   36
         Top             =   4320
         Width           =   11415
         Begin VB.ComboBox personlevelus 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox personlevelus 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2760
            TabIndex        =   50
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox personlevelus 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5400
            TabIndex        =   49
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox personnameus 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   48
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox personnameus 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3720
            TabIndex        =   47
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox personnameus 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6360
            TabIndex        =   46
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox personlevelcom 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3360
            TabIndex        =   45
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox personlevelcom 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   6000
            TabIndex        =   44
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox personlevelcom 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   8640
            TabIndex        =   43
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox personnamecom 
            Appearance      =   0  '平面
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4320
            TabIndex        =   42
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox personnamecom 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   6960
            TabIndex        =   41
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox personnamecom 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   9600
            TabIndex        =   40
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton opnpersonvs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3v3"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton personreadifus 
            Caption         =   "讀入..."
            Height          =   495
            Left            =   2040
            TabIndex        =   37
            Top             =   720
            Width           =   975
         End
         Begin MSComDlg.CommonDialog cdgpersonus 
            Left            =   3000
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "UnlightVBE-卡片人物資訊-開啟檔案"
         End
         Begin VB.OptionButton opnpersonvs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1v1"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   855
         End
         Begin ImageX.aicAlphaImage aicAlphaImage1 
            Height          =   660
            Left            =   5160
            Top             =   600
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   1164
            Image           =   "FormMainMode.frx":2E650
            Props           =   5
         End
         Begin UnlightVBE.大人物形像 personfus 
            Height          =   1215
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2143
         End
         Begin VB.Image PEGFbnstart 
            Height          =   510
            Left            =   9600
            Picture         =   "FormMainMode.frx":2F53D
            Top             =   600
            Width           =   1440
         End
         Begin VB.Image bnabout 
            Height          =   390
            Left            =   8280
            Picture         =   "FormMainMode.frx":30093
            Top             =   720
            Width           =   1320
         End
         Begin VB.Image bnconfig 
            Height          =   390
            Left            =   7080
            Picture         =   "FormMainMode.frx":306FE
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label personresetus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   59
            Top             =   360
            Width           =   495
         End
         Begin VB.Label personresetus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   58
            Top             =   360
            Width           =   495
         End
         Begin VB.Label personresetus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   6480
            TabIndex        =   57
            Top             =   360
            Width           =   495
         End
         Begin VB.Label personresetcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   56
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label personresetcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   6960
            TabIndex        =   55
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label personresetcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "重設"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   9600
            TabIndex        =   54
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label PEGFplayertext 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '透明
            Caption         =   "1P"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   8160
            TabIndex        =   53
            Top             =   0
            Width           =   375
         End
         Begin VB.Label PEGFcomtext 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '透明
            Caption         =   "COM"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   2400
            TabIndex        =   52
            Top             =   1395
            Width           =   855
         End
      End
      Begin VB.PictureBox PEGFcardus 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   1
         Left            =   120
         Picture         =   "FormMainMode.frx":30D0E
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   32
         Top             =   600
         Width           =   2535
         Begin VB.Label PEGFusbi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   550
            TabIndex        =   35
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEGFusbi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   34
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFusbi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   33
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardus 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   2
         Left            =   2760
         Picture         =   "FormMainMode.frx":34BB1
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   28
         Top             =   600
         Width           =   2535
         Begin VB.Label PEGFusbi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   550
            TabIndex        =   31
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEGFusbi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   30
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFusbi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   29
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardus 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   3
         Left            =   5400
         Picture         =   "FormMainMode.frx":38A54
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   24
         Top             =   600
         Width           =   2535
         Begin VB.Label PEGFusbi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   550
            TabIndex        =   27
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEGFusbi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   26
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFusbi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   25
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardcom 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   1
         Left            =   3360
         Picture         =   "FormMainMode.frx":3C8F7
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   20
         Top             =   6240
         Width           =   2535
         Begin VB.Label PEGFcardcompi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   23
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   22
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   21
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardcom 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   2
         Left            =   6000
         Picture         =   "FormMainMode.frx":4079A
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   16
         Top             =   6240
         Width           =   2535
         Begin VB.Label PEGFcardcompi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   19
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   18
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   17
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardcom 
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   3
         Left            =   8640
         Picture         =   "FormMainMode.frx":4463D
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   12
         Top             =   6240
         Width           =   2535
         Begin VB.Label PEGFcardcompi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   15
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   14
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   13
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "GameSetting"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   "自由戰鬥模式遊戲引導設定"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   61
         Top             =   195
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   465
         Left            =   0
         Picture         =   "FormMainMode.frx":484E0
         Top             =   0
         Width           =   11400
      End
   End
   Begin VB.PictureBox PEStartForm 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   3480
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   63
      Top             =   960
      Visible         =   0   'False
      Width           =   11340
      Begin VB.Timer tr1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   9720
         Top             =   8400
      End
      Begin VB.Label PEStext1 
         Alignment       =   1  '靠右對齊
         BackStyle       =   0  '透明
         Caption         =   "Now  Loading..."
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   8280
         TabIndex        =   64
         Top             =   9120
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.PictureBox PEMusicForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   1320
      ScaleHeight     =   7935
      ScaleWidth      =   11895
      TabIndex        =   130
      Top             =   1200
      Visible         =   0   'False
      Width           =   11895
      Begin VB.Timer PEMFtr1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   240
         Top             =   1680
      End
      Begin UnlightVBE.ucMusicPlayer cMusicPlayer 
         Height          =   855
         Index           =   0
         Left            =   840
         TabIndex        =   131
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1508
      End
   End
   Begin VB.PictureBox PEAttackingStartForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   3120
      Picture         =   "FormMainMode.frx":4A8CF
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   65
      Top             =   240
      Visible         =   0   'False
      Width           =   11340
      Begin VB.Timer PEASpke 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   240
         Top             =   2880
      End
      Begin VB.Timer start1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   45
         Top             =   4680
      End
      Begin VB.Timer start2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   525
         Top             =   4680
      End
      Begin VB.PictureBox PEAScardcom 
         Appearance      =   0  '平面
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   1
         Left            =   6285
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEAScardcompi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   69
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   68
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   67
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEAScardcom 
         Appearance      =   0  '平面
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   2
         Left            =   7485
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   96
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEAScardcompi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   99
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEAScardcompi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   98
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   97
            Top             =   3240
            Width           =   495
         End
      End
      Begin VB.PictureBox PEAScardus 
         Appearance      =   0  '平面
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   1
         Left            =   2805
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   70
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEASusbi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   550
            TabIndex        =   73
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEASusbi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   72
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEASusbi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   71
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEAScardus 
         Appearance      =   0  '平面
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   2
         Left            =   1485
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   88
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEASusbi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   91
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEASusbi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   90
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEASusbi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   550
            TabIndex        =   89
            Top             =   3240
            Width           =   375
         End
      End
      Begin VB.PictureBox PEAScardcom 
         Appearance      =   0  '平面
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   3
         Left            =   8565
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   100
         Top             =   3960
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEAScardcompi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   103
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEAScardcompi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   102
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   101
            Top             =   3240
            Width           =   495
         End
      End
      Begin VB.PictureBox PEAScardus 
         Appearance      =   0  '平面
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   3
         Left            =   360
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   92
         Top             =   3960
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEASusbi3 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   95
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEASusbi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   94
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEASusbi1 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   550
            TabIndex        =   93
            Top             =   3240
            Width           =   375
         End
      End
      Begin VB.PictureBox downjpg 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   45
         Picture         =   "FormMainMode.frx":79783
         ScaleHeight     =   1455
         ScaleWidth      =   11415
         TabIndex        =   75
         Top             =   8160
         Visible         =   0   'False
         Width           =   11415
         Begin VB.Label cardusname 
            BackStyle       =   0  '透明
            Caption         =   "人物1"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   87
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label cardcomname 
            BackStyle       =   0  '透明
            Caption         =   "人物1"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   6840
            TabIndex        =   86
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label cardusspname 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "稱號1"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   85
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label cardcomspname 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "稱號1"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   1
            Left            =   7920
            TabIndex        =   84
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label cardusname 
            BackStyle       =   0  '透明
            Caption         =   "人物2"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   83
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label cardusname 
            BackStyle       =   0  '透明
            Caption         =   "人物3"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   82
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label cardusspname 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "稱號2"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   81
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label cardusspname 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "稱號3"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   80
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label cardcomname 
            BackStyle       =   0  '透明
            Caption         =   "人物2"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   6840
            TabIndex        =   79
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label cardcomname 
            BackStyle       =   0  '透明
            Caption         =   "人物3"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   6840
            TabIndex        =   78
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label cardcomspname 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "稱號2"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   2
            Left            =   7920
            TabIndex        =   77
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label cardcomspname 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "稱號3"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   3
            Left            =   7920
            TabIndex        =   76
            Top             =   840
            Width           =   3135
         End
      End
      Begin VB.PictureBox upjpg 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   1900
         Left            =   0
         Picture         =   "FormMainMode.frx":820FF
         ScaleHeight     =   1905
         ScaleWidth      =   11415
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   11415
      End
      Begin VB.Timer stup 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   45
         Top             =   1800
      End
      Begin VB.Timer stdown 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   45
         Top             =   6600
      End
      Begin VB.Timer cardustr 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3405
         Top             =   7200
      End
      Begin VB.Timer cardcomtr 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7245
         Top             =   7320
      End
      Begin VB.Timer tr大人物形像_使用者 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1800
         Top             =   7440
      End
      Begin VB.Timer tr大人物形像_電腦 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   9720
         Top             =   7560
      End
      Begin UnlightVBE.uc對話 PEASpersontalk 
         Height          =   1935
         Left            =   0
         TabIndex        =   110
         Top             =   -120
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3413
      End
      Begin UnlightVBE.大人物形像 大人物形像_電腦 
         Height          =   10005
         Left            =   20040
         TabIndex        =   104
         Top             =   -480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   17648
      End
      Begin UnlightVBE.大人物形像 大人物形像_使用者 
         Height          =   10005
         Left            =   -9960
         TabIndex        =   105
         Top             =   -480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   17648
      End
      Begin UnlightVBE.大人物形像 upjpg_2 
         Height          =   1935
         Left            =   0
         TabIndex        =   106
         Top             =   -480
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3413
      End
   End
   Begin VB.PictureBox PEAttackingEndingForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   9120
      Picture         =   "FormMainMode.frx":8D62B
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   107
      Top             =   -1680
      Visible         =   0   'False
      Width           =   11340
      Begin VB.Timer PEAEtr1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5760
         Top             =   8400
      End
      Begin VB.Label bnt 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "結束遊戲"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9480
         TabIndex        =   109
         Top             =   8760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label bnreturnt 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "返回選單"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   108
         Top             =   8760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image bn 
         Height          =   990
         Left            =   9480
         Picture         =   "FormMainMode.frx":B0596
         Top             =   8520
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Image bnreturn 
         Height          =   990
         Left            =   7680
         Picture         =   "FormMainMode.frx":B148B
         Top             =   8520
         Visible         =   0   'False
         Width           =   1470
      End
   End
End
Attribute VB_Name = "FormMainMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub atkingtrcom_Timer()
If 目前數(29) = 1 Then
   目前數(31) = 0
   Formatkingcom.Left = FormMainMode.Left + (FormMainMode.Width - Formatkingcom.Width)
   Formatkingcom.Top = FormMainMode.Top + 380
   atkingtrcom.Enabled = False
   Formatkingcom.t1.Enabled = True
   Formatkingcom.Show 0, Me
Else
   目前數(29) = 目前數(29) + 1
End If
End Sub

Private Sub atkingtrus_Timer()
If 目前數(29) = 1 Then
   目前數(31) = 0
   Formatkingus.Left = FormMainMode.Left
   Formatkingus.Top = FormMainMode.Top + 380
   atkingtrus.Enabled = False
   Formatkingus.t1.Enabled = True
   Formatkingus.Show 0, Me
Else
   目前數(29) = 目前數(29) + 1
End If
End Sub

Private Sub bloodnumus1_Change()
If Val(bloodnumus1.Caption) < 0 Then bloodnumus1.Caption = 0
End Sub

Private Sub bn_Click()
End
End Sub

Private Sub bnreturn_Click()
bnreturnt_Click
End Sub

Sub bnreturnt_Click()
接續讀入表單串 = "PEGF"
一般系統類.主選單_PEStartForm顯示
FormMainMode.PEAttackingEndingForm.Visible = False
End Sub

Private Sub bnt_Click()
End
End Sub

Sub card_CardButtonClickin(Index As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = 戰鬥系統類.CardDeckCollection(5)(CStr(Index))

Call tmpcard.Reverse
一般系統類.音效播放 3
End Sub

Sub card_CardButtonClickout(Index As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = 戰鬥系統類.CardDeckCollection(6)(CStr(Index))

Call tmpcard.Reverse
FormMainMode.card(Index).CardRotationType = tmpcard.CardOnIn
一般系統類.音效播放 3
'===================================================================
If tmpcard.UpperType = a1a Then
   atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) + Val(tmpcard.UpperNum)
   If turnatk = 1 And movecp = 1 And 攻擊防禦骰子總數(3) = 0 Then
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
   End If
   If turnatk = 1 And movecp = 1 Then
       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(tmpcard.UpperNum)
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a5a Then
   atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) + Val(tmpcard.UpperNum)
   If turnatk = 1 And movecp > 1 And 攻擊防禦骰子總數(3) = 0 Then
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
   End If
   If turnatk = 1 And movecp > 1 Then
       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(tmpcard.UpperNum)
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a2a Then
   atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) + Val(tmpcard.UpperNum)
   If turnatk = 2 And 攻擊防禦骰子總數(3) = 0 Then
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + defus(角色人物對戰人數(1, 2))
   End If
   If turnatk = 2 Then
      攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(tmpcard.UpperNum)
      攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a3a Then
   atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) + Val(tmpcard.UpperNum)
End If
If tmpcard.UpperType = a4a Then
   atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) + Val(tmpcard.UpperNum)
End If
'======================================
If tmpcard.LowerType = a1a Then
   atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) - Val(tmpcard.LowerNum)
   If turnatk = 1 And movecp = 1 Then
       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(tmpcard.LowerNum)
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(tmpcard.LowerNum)
   End If
   If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
       攻擊防禦骰子總數(3) = 0
   End If
End If
If tmpcard.LowerType = a5a Then
   atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) - Val(tmpcard.LowerNum)
   If turnatk = 1 And movecp > 1 Then
       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(tmpcard.LowerNum)
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(tmpcard.LowerNum)
   End If
   If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
       攻擊防禦骰子總數(3) = 0
   End If
End If
If tmpcard.LowerType = a2a Then
   atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) - Val(tmpcard.LowerNum)
   If turnatk = 2 Then
       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(tmpcard.LowerNum)
       攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(tmpcard.LowerNum)
   End If
End If
If tmpcard.LowerType = a3a Then
   atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) - Val(tmpcard.LowerNum)
End If
If tmpcard.LowerType = a4a Then
   atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) - Val(tmpcard.LowerNum)
End If
'==============================================
Select Case turnatk
    Case 1
        '===========================執行階段插入點(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
        '============================
    Case 2
        '===========================執行階段插入點(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
        '============================
    Case 3
        '===========================執行階段插入點(44)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 44
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
        '============================
End Select
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1.Enabled = True
End Sub


Sub card_CardClick(Index As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(Index)))(CStr(Index))
'======================以下為專屬事件卡檢查
If tmpcard.UpperType = a7a And turnatk <> 1 And turnatk <> 2 Then
   '=========違反詛咒術事件卡只在攻防階段使用原則
   Exit Sub
End If
'====================================
If tmpcard.Location = 1 And (turnpageonin = 1 Or turnpageoninatking = 1) And tmpcard.Owner = 1 Then
   tmpcard.Location = 2
   If tmpcard.UpperType = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) + Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp = 1 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 1 And movecp = 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(tmpcard.UpperNum)
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) + Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp > 1 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 1 And movecp > 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(tmpcard.UpperNum)
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) + Val(tmpcard.UpperNum)
      If turnatk = 2 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + defus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 2 Then
         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(tmpcard.UpperNum)
         攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) + Val(tmpcard.UpperNum)
   End If
   If tmpcard.UpperType = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) + Val(tmpcard.UpperNum)
   End If
   '=================
   turnpageonin = 0
   card(Index).LocationType = 2
   '===================
   目前數(5) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(5), tmpcard)
   pageqlead(1) = Val(pageqlead(1)) + 1
   pageusglead = Val(pageusglead) - 1
   pageusleadmax(1) = Val(pageusleadmax(1)) + 1
   pageusqlead = Val(pageusqlead) + 1
   目前數(13) = 0
   '===================以下是出牌對齊
   目前數(3) = 0
   使用者出牌_出牌對齊_靠左.Enabled = True
   '=============以下是牌移動(出牌)(使用者)
    戰鬥系統類.座標計算_使用者出牌
    牌移動暫時變數(3) = Index
    tmpcard.XYLeft = card(Index).Left  '指定目前Left(座標)
    tmpcard.XYTop = card(Index).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.卡牌牌堆集合更換 tmpcard, 5, 6
    目前數(15) = 0
    牌移動.Enabled = True
    一般系統類.音效播放 1
   '================以下是手牌對齊
   目前數(4) = 0
   目前數(21) = 1
   使用者出牌_手牌對齊.Enabled = True
   '=================
   If tmpcard.UpperType = a6a Or tmpcard.UpperType = a7a Or tmpcard.UpperType = a8a Or tmpcard.UpperType = a9a Then
        '===================以下是事件卡檢查及啟動
        對齊完成檢查.Enabled = False
        事件卡記錄暫時數(1, 3) = 1
        Select Case tmpcard.UpperType
            Case a6a
                事件卡.機會_使用者 Index, tmpcard.UpperNum
            Case a7a
                事件卡.詛咒術_使用者 Index, tmpcard.UpperNum
            Case a8a
                事件卡.HP回復_使用者 Index, tmpcard.UpperNum
            Case a9a
                事件卡.聖水_使用者 Index, tmpcard.UpperNum
        End Select
        '===================
        Exit Sub
    Else
        對齊完成檢查.Enabled = True
        GoTo vsssystemplay
    End If
End If
'=================================================================
If tmpcard.Location = 2 And (turnpageonin = 1 Or turnpageoninatking = 1) And tmpcard.Owner = 1 Then
   tmpcard.Location = 1
   '===================================
   If tmpcard.UpperType = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) - Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp = 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(tmpcard.UpperNum)
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(tmpcard.UpperNum)
      End If
      If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
          攻擊防禦骰子總數(3) = 0
      End If
   End If
   If tmpcard.UpperType = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) - Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp > 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(tmpcard.UpperNum)
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(tmpcard.UpperNum)
      End If
      If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
          攻擊防禦骰子總數(3) = 0
      End If
   End If
   If tmpcard.UpperType = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) - Val(tmpcard.UpperNum)
      If turnatk = 2 Then
         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(tmpcard.UpperNum)
         攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) - Val(tmpcard.UpperNum)
   End If
   If tmpcard.UpperType = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) - Val(tmpcard.UpperNum)
   End If
   '=============
   turnpageonin = 0
   card(Index).LocationType = 1
   '================
   目前數(5) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(6), tmpcard)
   pageusleadmax(0) = Val(pageusleadmax(0)) + 1
   pageqlead(1) = Val(pageqlead(1)) - 1
   pageusglead = Val(pageusglead) + 1
   pageusqlead = Val(pageusqlead) - 1
   '=============以下是牌移動(回牌)(使用者)
    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(3) = Index
    tmpcard.XYLeft = card(Index).Left  '指定目前Left(座標)
    tmpcard.XYTop = card(Index).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.卡牌牌堆集合更換 tmpcard, 6, 5
    目前數(15) = 0
    牌移動.Enabled = True
    一般系統類.音效播放 1
   '================以下是出牌對齊
   目前數(3) = 0
   使用者出牌_出牌對齊_靠右.Enabled = True
   '=====================
   對齊完成檢查.Enabled = True
   '=====================以下是技能檢查及啟動
    If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
        vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 2 '(階段2)
    End If
    '====================
    GoTo vsssystemplay
End If
'==============================================
Exit Sub
vsssystemplay:
Select Case turnatk
    Case 1
        '===========================執行階段插入點(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
        '============================
    Case 2
        '===========================執行階段插入點(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
        '============================
    Case 3
        '===========================執行階段插入點(44)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 44
        VBEStageNum(1) = -1 '觸發方(1.使用者/2.電腦)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
        '============================
End Select
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1.Enabled = True
End Sub


Private Sub card_CardMouseMove(Index As Integer)

If 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(Index)) = 5 And turnpageonin = 1 Then
    card(Index).CardEventType = True
ElseIf 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(Index)) = 6 And turnpageonin = 1 Then
    card(Index).CardEventType = True
Else
    card(Index).CardEventType = False
End If
End Sub

Sub cnmove_Click()
Dim i As Integer, med As Integer
Dim tmpcard As clsActionCard
'======================
If 電腦方事件卡是否出完選擇數 = True Then
    GoTo 電腦方事件卡先出制度_執行階段結束
End If
'======================
If 角色人物對戰人數(1, 1) > 1 Or 角色人物對戰人數(2, 1) > 1 Then
   顯示列1.人物戰鬥人數 = 3
Else
   顯示列1.人物戰鬥人數 = 1
End If
'======================
movecom = 0
movecheckcom = 0
顯示列1.移動階段選擇值 = 0
電腦方移動階段選擇數 = 0
atkingtrn(1) = 0
atkingtrn(2) = 0
turnatk = 3
pageusqlead.Caption = 0
pagecomqlead.Caption = 0
目前數(6) = 0
目前數(17) = 1
目前數(21) = 1
目前數(25) = 0
階段狀態數 = 3
'=============
If 系統顯示界面紀錄數 = 1 Then
    draw2.Visible = False
    draw1.Visible = True
    move1.Visible = False
    move2.Visible = True
Else
    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\move1-2.gif"
End If
顯示列1.顯示列圖片 = app_path & "gif\system\linemove.png"
cnmove.Visible = False
戰鬥系統類.cleanatkingpagetot
'======================電腦方事件卡先出制度
If 電腦方事件卡是否出完選擇數 = False Then
    GoTo 電腦方事件卡先出制度_執行階段2
End If
'================================
電腦方事件卡先出制度_執行階段結束:
'----------以下為電腦判斷出牌程式碼（移動階段1）
'====================試驗智慧型AI出牌系統
    智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 2, 3, namecom(角色人物對戰人數(2, 2)), movecp, 0
    GoTo 智慧型AI出牌_執行階段結束
'======================
Dim movecomatk1, movecomatk2 As Integer
戰鬥系統類.moveatkin

For i = 1 To 戰鬥系統類.CardDeckCollection(7).Count
    Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(i)
    If tmpcard.ComMark <> 1 Then
        If tmpcard.UpperType = a1a Then movecomatk1 = Val(movecomatk1) + Val(tmpcard.UpperNum)
        If tmpcard.UpperType = a5a Then movecomatk2 = Val(movecomatk2) + Val(tmpcard.UpperNum)
        If tmpcard.LowerType = a1a Then movecomatk1 = Val(movecomatk1) + Val(tmpcard.LowerNum)
        If tmpcard.LowerType = a5a Then movecomatk2 = Val(movecomatk2) + Val(tmpcard.LowerNum)
    End If
Next
'===========================================
麻痺_電腦_執行階段2: '異常狀態-麻痺-電腦-程式跳入點(執行階段2)
'===========================================
If movecomatk1 > movecomatk2 Then
      電腦方移動階段選擇數 = 1
ElseIf movecomatk1 = movecomatk2 Then
      med = Int(Rnd() * 2) + 1
      If med = 1 Then
         電腦方移動階段選擇數 = 1
      Else
         電腦方移動階段選擇數 = 3
      End If
Else
      電腦方移動階段選擇數 = 3
End If
'==============
智慧型AI出牌_執行階段結束:
電腦方事件卡先出制度_執行階段2:
If 電腦方事件卡是否出完選擇數 = False Then
    '==============
    小人物頭像移動方向數(1) = 1
    小人物頭像移動方向數(2) = 1
    小人物頭像移動_使用者.Enabled = True
    小人物頭像移動_電腦.Enabled = True
    '==============
    階段狀態數 = 1
    戰鬥系統類.時間軸_重設
    顯示列1.移動階段圖顯示 = True
    戰鬥系統類.時間軸_顯示
    一般系統類.音效播放 6
    '===========================執行階段插入點(94)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 94, 3
    '============================
End If
'======================電腦方事件卡先出制度_結束後階段2
If 電腦方事件卡是否出完選擇數 = True Then
    電腦出牌.Enabled = True
End If
'===========================
End Sub

Private Sub cnmove2_Click()
turnpageonin = 0
目前數(31) = 0
OK按鈕牌完成移動檢查.Enabled = True
cnmove2.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
一般系統類.離開遊戲提示 Cancel, UnloadMode
End Sub

Private Sub cn1_Click()
turnatk = 4
戰鬥系統類.音量靜音調節設定
'====================
目前數(2) = 1
電腦方事件卡是否出完選擇數 = False
'===========================執行階段插入點(0)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 0, 1
'============================
cn1.Visible = False
目前數(15) = 1
發牌檢查.Enabled = True
End Sub

Private Sub cn2_Click()
If moveturn = 1 Then
  If 系統顯示界面紀錄數 = 1 Then
        move1.Visible = True
        move2.Visible = False
        atkdef1.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\atk2.gif"
  End If
  顯示列1.goi1顯示 = True
  顯示列1.goi2顯示 = True
  顯示列1.移動階段選擇值 = 0
  顯示列1.移動階段圖顯示 = False
Else
  If 系統顯示界面紀錄數 = 1 Then
        atkdef1.Visible = False
        atkdef2.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\atk2.gif"
  End If
End If
'-------------
turnatk = 1
階段狀態數 = 1
If movecp = 1 Then
    顯示列1.顯示列圖片 = app_path & "gif\system\lineusatk1.png"
Else
    顯示列1.顯示列圖片 = app_path & "gif\system\lineusatk2.png"
End If
cn2.Visible = False
FormMainMode.PEAFInterface.BnOKStartListen
'=============
戰鬥系統類.cleanatkingpagetot
'==============
顯示列1.goi1 = 0
顯示列1.goi2 = 0
目前數(6) = 0
目前數(17) = 1
目前數(21) = 1
目前數(15) = 0
攻擊防禦骰子總數(1) = 0
攻擊防禦骰子總數(2) = 0
攻擊防禦骰子總數(3) = 0
攻擊防禦骰子總數(4) = 0
骰數零檢查值(1) = False
骰數零檢查值(2) = False
是否系統公骰 = False
'==============
goicheck(1) = 0
goicheck(2) = 0
chkcomck = 0
atkingtrn(1) = 0
atkingtrn(2) = 0
'=====
If turnatk = 1 Then
 戰鬥系統類.chkdefcom
End If
'======================================
Erase Vss_EventPlayerAllActionOffNum
'===========================執行階段插入點(ATK-17/DEF-37)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 17, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 37, 2
'===========================執行階段插入點(ATK-92/DEF-93)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 92, 4
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 93, 4
'==============
小人物頭像移動方向數(1) = 1
小人物頭像移動方向數(2) = 2
小人物頭像移動_使用者.Enabled = True
小人物頭像移動_電腦.Enabled = True
'==============
一般系統類.音效播放 6
戰鬥系統類.時間軸_重設
trtimeline.Enabled = True
trgoi2.Enabled = True
'==============
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
End Sub



Private Sub cn22_Click()
cn22.Visible = False
OK按鈕牌完成移動檢查.Enabled = True
End Sub

Sub cn3_Click()
'======================
If 電腦方事件卡是否出完選擇數 = True Then
    GoTo 電腦方事件卡先出制度_執行階段結束
End If
'======================
If moveturn = 2 Then
  If 系統顯示界面紀錄數 = 1 Then
        move1.Visible = True
        move2.Visible = False
        atkdef1.Visible = True
        atkdef2.Visible = False
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\def2.gif"
  End If
  顯示列1.goi1顯示 = True
  顯示列1.goi2顯示 = True
  顯示列1.移動階段選擇值 = 0
  顯示列1.移動階段圖顯示 = False
Else
  If 系統顯示界面紀錄數 = 1 Then
        atkdef1.Visible = False
        atkdef2.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\def2.gif"
  End If
End If
turnatk = 2
顯示列1.顯示列圖片 = app_path & "gif\system\lineusdef.png"
戰鬥系統類.cleanatkingpagetot
'===============
顯示列1.goi1 = 0
顯示列1.goi2 = 0
攻擊防禦骰子總數(1) = 0
攻擊防禦骰子總數(2) = 0
攻擊防禦骰子總數(3) = 0
攻擊防禦骰子總數(4) = 0
骰數零檢查值(1) = False
骰數零檢查值(2) = False
是否系統公骰 = False
'=====
目前數(6) = 0
目前數(21) = 1
'===============
goicheck(1) = 0
goicheck(2) = 0
atkingtrn(1) = 0
atkingtrn(2) = 0
If turnatk = 2 Then
 戰鬥系統類.chkdef
End If
'==============
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'============================
Erase Vss_EventPlayerAllActionOffNum
'===========================執行階段插入點(ATK-17/DEF-37)
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 17, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 37, 2
'===========================執行階段插入點(ATK-92/DEF-93)
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 92, 4
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 93, 4
'======================電腦方事件卡先出制度
If 電腦方事件卡是否出完選擇數 = False Then
   GoTo 電腦方事件卡先出制度_執行階段2
End If
'================================
電腦方事件卡先出制度_執行階段結束:
'----------以下為電腦判斷出牌程式碼（攻擊方）
'====================試驗智慧型AI出牌系統
Dim wtyr As Integer '暫時變數
If moveturn = 1 Then wtyr = 1 Else wtyr = 0
智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 2, 1, namecom(角色人物對戰人數(2, 2)), movecp, wtyr
GoTo 智慧型AI出牌_執行階段結束
 '==================
If turnatk = 2 And movecp = 1 Then
   戰鬥系統類.comatk1
ElseIf turnatk = 2 And movecp > 1 Then
   戰鬥系統類.comatk2
End If
'==============================
智慧型AI出牌_執行階段結束:
'==============================
電腦方事件卡先出制度_執行階段2:
If 電腦方事件卡是否出完選擇數 = False Then
    '==========
    cn3.Visible = False
    目前數(6) = 0
    目前數(17) = 1
    目前數(15) = 0
    '==============
    小人物頭像移動方向數(1) = 2
    小人物頭像移動方向數(2) = 1
    小人物頭像移動_使用者.Enabled = True
    小人物頭像移動_電腦.Enabled = True
    '==============
    戰鬥系統類.時間軸_重設
    trtimeline.Enabled = True
ElseIf 電腦方事件卡是否出完選擇數 = True Then  '電腦方事件卡先出制度_結束後階段2
    電腦出牌.Enabled = True
End If
End Sub

Private Sub cn32_Click()
cn32.Visible = False
OK按鈕牌完成移動檢查.Enabled = True
End Sub

Private Sub cn4_Click()
Dim uscomvsn As Integer
cn4.Visible = False
turnatk = 5
If moveturn = 1 Then uscomvsn = 2 Else uscomvsn = 1
'===========================執行階段插入點(50)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 50, 1
'============================
'===========================執行階段插入點(51)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 51, 1
'============================
'===========================執行階段插入點(52)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 52, 1
'============================
HP檢查階段數 = 4
戰鬥系統類.雙方HP檢查
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
For i = 1 To cMusicPlayer.UBound
    Unload cMusicPlayer(i)
Next
End Sub

Private Sub NextTurn_階段2_Timer()
Dim uscomvsn As Integer
Dim i As Integer, j As Integer, k As Integer
goidefus = 0
'======以下為洗牌程式碼
If BattleCardNum < 牌總階段數(1) + 牌總階段數(2) Then
    戰鬥系統類.執行動作_洗牌
End If
'==========================
If moveturn = 1 Then uscomvsn = 2 Else uscomvsn = 1
'===========================執行階段插入點(53)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 53, 1
'============================
'===========================執行階段插入點(54)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 54, 1
'============================
'===========================執行階段插入點(55)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 55, 1
'============================
戰鬥系統類.廣播訊息 BattleTurn & "回合結束。"
'=============
NextTurn_階段2.Enabled = False
'=============
If 戰鬥系統類.雙方HP檢查_結束回合檢查 = True Then
    Exit Sub
End If
'==============將每回合之啟動次數歸零
For i = 1 To 2
   For j = 1 To 3
       For k = 1 To 8
           atkingck(i, j, k, 2) = 0
       Next
    Next
Next
'==============
BattleTurn = BattleTurn + 1
PEAFInterface.turn = BattleTurn
顯示列1.goi1顯示 = False
顯示列1.goi2顯示 = False
顯示列1.goi1 = 0
顯示列1.goi2 = 0
攻擊防禦骰子總數(1) = 0
攻擊防禦骰子總數(2) = 0
'====================
If 系統顯示界面紀錄數 = 1 Then
    move1.Visible = True
    move2.Visible = False
    atkdef1.Visible = False
    atkdef2.Visible = False
    move3.Picture = LoadPicture(app_path & "gif\system\move3.gif")
    move4.Picture = LoadPicture(app_path & "gif\system\move4.gif")
Else
    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\stageblack.gif"
End If
顯示列1.顯示列圖片 = app_path & "gif\system\DRAW.png"
'==============
小人物頭像移動方向數(1) = 2
小人物頭像移動方向數(2) = 2
小人物頭像移動_使用者.Enabled = True
小人物頭像移動_電腦.Enabled = True
'==============
等待時間佇列(2).Add 1
等待時間_2.Enabled = True
End Sub

Private Sub OK按鈕牌完成移動檢查_Timer()
If 使用者出牌_出牌對齊_靠左.Enabled = False And 使用者出牌_出牌對齊_靠右.Enabled = False And 使用者出牌_手牌對齊.Enabled = False And 對齊完成檢查.Enabled = False Then
   OK按鈕牌完成移動檢查.Enabled = False
   turnpageonin = 0
   Select Case turnatk
       Case 1
           攻擊階段_階段初始.Enabled = True
       Case 2
           防禦階段_階段初始.Enabled = True
       Case 3
           移動階段_階段初始.Enabled = True
   End Select
End If
End Sub

Private Sub pagecomglead_Change()
pageglead(2) = Val(pagecomglead.Caption)
End Sub

Private Sub pageusglead_Change()
pageglead(1) = Val(pageusglead.Caption)
End Sub

Private Sub PEAEtr1_Timer()
Select Case PEAEtr1num
    Case 10
         If 戰鬥模式勝敗紀錄數 = 1 Then
             FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(app_path & "gif\system\gamewin.jpg")
         ElseIf 戰鬥模式勝敗紀錄數 = 2 Then
             FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(app_path & "gif\system\gamelose.jpg")
         ElseIf 戰鬥模式勝敗紀錄數 = 3 Then
         
         End If
         FormMainMode.cMusicPlayer(0).MusicPlay
    Case 50
         PEAEtr1.Enabled = False
         '======================
         戰鬥系統類.遊戲對戰結束物件消滅
         '======================
         If Formsetting.chkautocontinuemode.Value = 1 Then
            bnreturnt_Click
         End If
         bnreturn.Visible = True
         bnreturnt.Visible = True
         bn.Visible = True
         bnt.Visible = True
End Select
PEAEtr1num = PEAEtr1num + 1
End Sub

Private Sub PEAFAnimateInterface_AnimateCheckPoint(ByVal uscom As Integer)
Vss_AtkingStartPlayNum(2) = 1 '技能執行中啟動
End Sub

Private Sub PEAFAnimateInterface_AnimateEnd(ByVal uscom As Integer)
Vss_AtkingStartPlayNum(3) = 1
End Sub

Private Sub PEAFInterface_ActiveMouseEnter(ByVal uscom As Integer, ByVal num As Integer)
Dim i As Integer
Dim tmpobj As clsPersonActiveSkill

Select Case uscom
 Case 1
    For i = 1 To 戰鬥系統類.ActionCardTotNum
       card(i).CardEventType = False
    Next
 Case 2
    For i = 1 To 3
      cardcom(i).Visible = False
    Next
End Select
'============================
Set tmpobj = 戰鬥系統類.ActiveSkillObj(uscom, num)
If Not tmpobj Is Nothing Then
    PEAFatkinghelpc.Stage = tmpobj.Stage
    PEAFatkinghelpc.Distance = tmpobj.Distance
    PEAFatkinghelpc.card = tmpobj.card
    PEAFatkinghelpc.Effect = tmpobj.Effect
    PEAFatkinghelpc.Left = atkinghelpxy(uscom, num, 1)
    PEAFatkinghelpc.Top = atkinghelpxy(uscom, num, 2)
    PEAFatkinghelpc.ZOrder
    PEAFatkinghelpc.Visible = True
End If
End Sub

Private Sub PEAFInterface_ActiveMouseExit(ByVal uscom As Integer, ByVal num As Integer)
PEAFatkinghelpc.Visible = False
End Sub

Sub PEAFInterface_BnOKClick()
Dim i As Integer
If turnpageonin = 1 Then
    turnpageonin = 0
    For i = 1 To 戰鬥系統類.ActionCardTotNum
        FormMainMode.card(i).card_MouseExit
    Next
    FormMainMode.PEAFInterface.BnOKStopListen
    戰鬥系統類.時間軸_停止
    Select Case turnatk
        Case 1
            等待時間佇列(1).Add 7
            等待時間.Enabled = True
        Case 2
            等待時間佇列(1).Add 8
            等待時間.Enabled = True
        Case 3
            cnmove2_Click
    End Select
End If
End Sub

Private Sub PEAFInterface_BnOKMouseMove()
Dim i As Integer
For i = 1 To 戰鬥系統類.ActionCardTotNum
   card(i).CardEventType = False
Next
End Sub

Private Sub PEAFInterface_InterfaceMouseMove()
Dim i As Integer
For i = 1 To 戰鬥系統類.ActionCardTotNum
   card(i).CardEventType = False
Next
For i = 1 To 3
  cardcom(i).Visible = False
Next
For i = 1 To 3
  If i <> 角色人物對戰人數(1, 2) Then
     cardus(i).Visible = False
  End If
Next
End Sub

Private Sub PEAFpersoncardcom_MouseEnter(Index As Integer)
cardcom(Index).Left = FormMainMode.PEAFpersoncardcom(Index).Left
cardcom(Index).Top = 480
cardcom(Index).Visible = True
cardcom(Index).ZOrder
Select Case Index
   Case 1
      cardcom(2).Visible = False
      cardcom(3).Visible = False
   Case 2
      cardcom(1).Visible = False
      cardcom(3).Visible = False
   Case 3
      cardcom(2).Visible = False
      cardcom(1).Visible = False
End Select
End Sub

Private Sub PEAFpersoncardus_MouseEnter(Index As Integer)
cardus(Index).Left = FormMainMode.PEAFpersoncardus(Index).Left
cardus(Index).Top = 5760
cardus(Index).Visible = True
cardus(Index).ZOrder
Select Case Index
   Case 1
      If 角色人物對戰人數(1, 2) = 2 Then
          cardus(3).Visible = False
      Else
          cardus(2).Visible = False
      End If
   Case 2
      If 角色人物對戰人數(1, 2) = 1 Then
          cardus(3).Visible = False
      Else
          cardus(1).Visible = False
      End If
   Case 3
      If 角色人物對戰人數(1, 2) = 2 Then
          cardus(1).Visible = False
      Else
          cardus(2).Visible = False
      End If
End Select
End Sub

Private Sub PEASpke_Timer()
If swq = 35 Then
    PEASpke.Enabled = False
    If PEASpersontalk.對話文字 <> "" Then
        PEASpersontalk.對話文字顯示 = True
    End If
ElseIf swq = 10 Then
    PEASpersontalk.Top = -120
    PEASpersontalk.對話文字 = 人物系統類.人物對話選擇
    If PEASpersontalk.對話文字 <> "" Then
        PEASpersontalk.Visible = True
        PEASpersontalk.對話文字顯示 = False
        PEASpersontalk.ZOrder
    End If
    swq = Val(swq) + 1
Else
    swq = Val(swq) + 1
End If

End Sub

Private Sub PEAttackingForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 1 To 戰鬥系統類.ActionCardTotNum
   card(i).CardEventType = False
Next
For i = 1 To 3
  cardcom(i).Visible = False
Next
For i = 1 To 3
  If i <> 角色人物對戰人數(1, 2) Then
     cardus(i).Visible = False
  End If
Next
End Sub

Private Sub tr1_Timer()
Select Case tr1num
    Case 1
'        MsgBox "1-1"
        PEStext1.Visible = True
    Case 3
        If 第一次啟動讀入程序標記 = False Then
'            一般系統類.遊戲初始讀入程序
            第一次啟動讀入程序標記 = True
            接續讀入表單串 = "PEGF"   '====測試階段-直接進入自由戰鬥模式
'            MsgBox "1-3"
        End If
    Case 5
        Select Case 接續讀入表單串
            Case "PEGF"
'                MsgBox "1-5"
                一般系統類.遊戲初始讀入程序
                一般系統類.自由戰鬥模式設定表單讀入程序
                一般系統類.自由戰鬥模式設定表單基本設定程序
        End Select
    Case 7
        Select Case 接續讀入表單串
            Case "PEGF"
'                MsgBox "1-7"
                一般系統類.主選單_PEGameFreeModeSettingForm顯示
        End Select
        tr1.Enabled = False
        PEStartForm.Visible = False
End Select
tr1num = tr1num + 1
End Sub

Private Sub trend_Timer()
If trend暫時變數 = 4 Then
   一般系統類.主選單_PEAttackingEndingForm顯示
   PEAttackingForm.Visible = False
   PEAEtr1num = 0
   PEAEtr1.Enabled = True
   trend.Enabled = False
ElseIf trend暫時變數 = 2 Then
   FormMainMode.cMusicPlayer(0).MusicStop
   FormMainMode.cMusicPlayer(0).IsLoop = False
   FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulse15.mp3"
   trend暫時變數 = trend暫時變數 + 1
Else
   trend暫時變數 = trend暫時變數 + 1
End If
End Sub

Sub trgoi1_Timer()
'=========更新骰子總數量表示
If 攻擊防禦骰子總數(1) < 0 Then
   顯示列1.goi1 = 0
Else
   顯示列1.goi1 = 攻擊防禦骰子總數(1)
End If
FormMainMode.trgoi1.Enabled = False
'=====================
End Sub

Sub trgoi2_Timer()
'=========更新骰子總數量表示
If 攻擊防禦骰子總數(2) < 0 Then
   顯示列1.goi2 = 0
Else
   顯示列1.goi2 = 攻擊防禦骰子總數(2)
End If
trgoi2.Enabled = False

End Sub

Private Sub trnextend_Timer()
Select Case Val(擲骰表單溝通暫時變數(3))
   Case 1
      傷害執行_使用者 (Val(擲骰表單溝通暫時變數(2)))
   Case 2
      傷害執行_電腦 (Val(擲骰表單溝通暫時變數(2)))
End Select
'=============
等待時間佇列(2).Add 21
等待時間_2.Enabled = True
trnextend.Enabled = False
End Sub

Private Sub trtimeline_Timer()
Dim i As Integer

timelineout1.X1 = timelineout1.X1 + 2
timelineout2.X2 = timelineout2.X2 - 2
For i = 1 To 3
   時間軸顏色變化紀錄暫時變數(2, i) = 時間軸顏色變化紀錄暫時變數(2, i) + 2
Next
Select Case timelineout1.X1
   Case Is <= 2624
       If 時間軸顏色變化紀錄暫時變數(2, 1) >= 時間軸顏色變化紀錄暫時變數(1, 1) Then
           時間軸顏色變化紀錄暫時變數(2, 1) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1) + 1, 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3))
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1) + 1, 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3))
           時間軸顏色變化紀錄暫時變數(3, 1) = 時間軸顏色變化紀錄暫時變數(3, 1) + 1
       End If
       If 時間軸顏色變化紀錄暫時變數(2, 2) >= 時間軸顏色變化紀錄暫時變數(1, 2) Then
           時間軸顏色變化紀錄暫時變數(2, 2) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2) - 1, 時間軸顏色變化紀錄暫時變數(3, 3))
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2) - 1, 時間軸顏色變化紀錄暫時變數(3, 3))
           時間軸顏色變化紀錄暫時變數(3, 2) = 時間軸顏色變化紀錄暫時變數(3, 2) - 1
       End If
       If timelineout1.X1 >= 2624 Then
            時間軸顏色變化紀錄暫時變數(1, 1) = 34
            時間軸顏色變化紀錄暫時變數(1, 2) = 13
            時間軸顏色變化紀錄暫時變數(1, 3) = 60
            時間軸顏色變化紀錄暫時變數(2, 1) = 0
            時間軸顏色變化紀錄暫時變數(2, 2) = 0
            時間軸顏色變化紀錄暫時變數(2, 3) = 0
            時間軸顏色變化紀錄暫時變數(3, 1) = 217
            時間軸顏色變化紀錄暫時變數(3, 2) = 217
            時間軸顏色變化紀錄暫時變數(3, 3) = 50
            timelineout1.BorderColor = RGB(217, 217, 50)
            timelineout2.BorderColor = RGB(217, 217, 50)
        End If
   Case Is <= 3936
        If 時間軸顏色變化紀錄暫時變數(2, 1) >= 時間軸顏色變化紀錄暫時變數(1, 1) Then
           時間軸顏色變化紀錄暫時變數(2, 1) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1) + 1, 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3))
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1) + 1, 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3))
           時間軸顏色變化紀錄暫時變數(3, 1) = 時間軸顏色變化紀錄暫時變數(3, 1) + 1
       End If
       If 時間軸顏色變化紀錄暫時變數(2, 2) >= 時間軸顏色變化紀錄暫時變數(1, 2) Then
           時間軸顏色變化紀錄暫時變數(2, 2) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2) - 1, 時間軸顏色變化紀錄暫時變數(3, 3))
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2) - 1, 時間軸顏色變化紀錄暫時變數(3, 3))
           時間軸顏色變化紀錄暫時變數(3, 2) = 時間軸顏色變化紀錄暫時變數(3, 2) - 1
       End If
       If 時間軸顏色變化紀錄暫時變數(2, 3) >= 時間軸顏色變化紀錄暫時變數(1, 3) Then
           時間軸顏色變化紀錄暫時變數(2, 3) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3) - 1)
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3) - 1)
           時間軸顏色變化紀錄暫時變數(3, 3) = 時間軸顏色變化紀錄暫時變數(3, 3) - 1
       End If
       If timelineout1.X1 >= 3936 Then
            時間軸顏色變化紀錄暫時變數(1, 1) = 0
            時間軸顏色變化紀錄暫時變數(1, 2) = 11
            時間軸顏色變化紀錄暫時變數(1, 3) = 47
            時間軸顏色變化紀錄暫時變數(2, 1) = 0
            時間軸顏色變化紀錄暫時變數(2, 2) = 0
            時間軸顏色變化紀錄暫時變數(2, 3) = 0
            時間軸顏色變化紀錄暫時變數(3, 1) = 255
            時間軸顏色變化紀錄暫時變數(3, 2) = 118
            時間軸顏色變化紀錄暫時變數(3, 3) = 28
            timelineout1.BorderColor = RGB(255, 118, 28)
            timelineout2.BorderColor = RGB(255, 118, 28)
            '=========時間軸(外)
            時間軸顏色變化紀錄暫時變數(4, 1) = 1
            時間軸顏色變化紀錄暫時變數(4, 2) = 0
            時間軸顏色變化紀錄暫時變數(4, 3) = 0
            timelinein1.BorderColor = RGB(0, 0, 0)
            timelinein2.BorderColor = RGB(0, 0, 0)
        End If
    Case Is > 3936
       If 時間軸顏色變化紀錄暫時變數(2, 2) >= 時間軸顏色變化紀錄暫時變數(1, 2) Then
           時間軸顏色變化紀錄暫時變數(2, 2) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2) - 1, 時間軸顏色變化紀錄暫時變數(3, 3))
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2) - 1, 時間軸顏色變化紀錄暫時變數(3, 3))
           時間軸顏色變化紀錄暫時變數(3, 2) = 時間軸顏色變化紀錄暫時變數(3, 2) - 1
       End If
       If 時間軸顏色變化紀錄暫時變數(2, 3) >= 時間軸顏色變化紀錄暫時變數(1, 3) Then
           時間軸顏色變化紀錄暫時變數(2, 3) = 0
           timelineout1.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3) - 1)
           timelineout2.BorderColor = RGB(時間軸顏色變化紀錄暫時變數(3, 1), 時間軸顏色變化紀錄暫時變數(3, 2), 時間軸顏色變化紀錄暫時變數(3, 3) - 1)
           時間軸顏色變化紀錄暫時變數(3, 3) = 時間軸顏色變化紀錄暫時變數(3, 3) - 1
       End If
       '===================時間軸(外)
       Select Case 時間軸顏色變化紀錄暫時變數(4, 1)
           Case 1
                    If 255 - Val(時間軸顏色變化紀錄暫時變數(4, 3)) < 9 Then
                       timelinein1.BorderColor = RGB(255, 0, 0)
                       timelinein2.BorderColor = RGB(255, 0, 0)
                       時間軸顏色變化紀錄暫時變數(4, 3) = 255
                    Else
                       timelinein1.BorderColor = RGB(Val(時間軸顏色變化紀錄暫時變數(4, 3)) + 9, 0, 0)
                       timelinein2.BorderColor = RGB(Val(時間軸顏色變化紀錄暫時變數(4, 3)) + 9, 0, 0)
                       時間軸顏色變化紀錄暫時變數(4, 3) = Val(時間軸顏色變化紀錄暫時變數(4, 3)) + 9
                    End If
                If 時間軸顏色變化紀錄暫時變數(4, 3) = 255 Then
                    時間軸顏色變化紀錄暫時變數(4, 1) = 2
                End If
           Case 2
               If 時間軸顏色變化紀錄暫時變數(4, 3) < 9 Then
                   timelinein1.BorderColor = RGB(0, 0, 0)
                   timelinein2.BorderColor = RGB(0, 0, 0)
                   時間軸顏色變化紀錄暫時變數(4, 3) = 0
                Else
                   timelinein1.BorderColor = RGB(Val(時間軸顏色變化紀錄暫時變數(4, 3)) - 9, 0, 0)
                   timelinein2.BorderColor = RGB(Val(時間軸顏色變化紀錄暫時變數(4, 3)) - 9, 0, 0)
                   時間軸顏色變化紀錄暫時變數(4, 3) = Val(時間軸顏色變化紀錄暫時變數(4, 3)) - 9
                End If
                If 時間軸顏色變化紀錄暫時變數(4, 3) = 0 Then
                    時間軸顏色變化紀錄暫時變數(4, 1) = 1
                End If
       End Select
End Select
If timelineout1.X1 >= timelineout1.X2 Then
    戰鬥系統類.時間軸_停止
    turnpageonin = 0
    FormMainMode.PEAFInterface.BnOKStopListen
    等待時間佇列(2).Add 4
    等待時間_2.Enabled = True
End If
End Sub

Private Sub tr使用者_棄牌_Timer()
戰鬥系統類.執行動作_使用者_棄牌 目前數(20)
tr使用者_棄牌.Enabled = False
End Sub

Private Sub tr使用者牌_偷牌_Timer()
戰鬥系統類.執行動作_使用者牌_偷牌_電腦 目前數(20)
tr使用者牌_偷牌.Enabled = False
End Sub

Private Sub tr牌組_回牌_使用者_Timer()
card(目前數(16)).Left = 240
card(目前數(16)).Top = 960
card(目前數(16)).Visible = True
戰鬥系統類.執行動作_牌組_回牌_使用者 目前數(16)
tr牌組_回牌_使用者.Enabled = False
End Sub

Sub tr牌組_回牌_電腦_Timer()
card(目前數(16)).Left = 240
card(目前數(16)).Top = 960
card(目前數(16)).Visible = True
戰鬥系統類.執行動作_牌組_回牌_電腦 目前數(16)
tr牌組_回牌_電腦.Enabled = False
End Sub


Private Sub tr牌組_抽牌_使用者_Timer()
tr牌組_抽牌_使用者.Enabled = False
If BattleCardNum > 0 Then
    戰鬥系統類.執行動作_抽牌_公用牌 1
End If
End Sub

Private Sub tr牌組_抽牌_電腦_Timer()
tr牌組_抽牌_電腦.Enabled = False
If BattleCardNum > 0 Then
    戰鬥系統類.執行動作_抽牌_公用牌 2
End If
End Sub

Private Sub tr電腦牌_偷牌_Timer()
戰鬥系統類.執行動作_電腦牌_偷牌_使用者 目前數(16)
tr電腦牌_偷牌.Enabled = False
End Sub

Private Sub tr電腦牌_棄牌_Timer()
戰鬥系統類.執行動作_電腦_棄牌 目前數(16)
tr電腦牌_棄牌.Enabled = False
End Sub

Private Sub tr電腦牌_翻牌_Timer()
    戰鬥系統類.執行動作_翻牌 目前數(16)
    tr電腦牌_翻牌.Enabled = False
    If 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards") <> 0 Then
        vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards")) = 2 '(階段2)
    End If
    If 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards") <> 0 Then
        vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards")) = 2 '(階段2)
    End If
    If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
        vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 2 '(階段2)
    End If
   '=======================以下是事件卡檢查及啟動
   If 事件卡記錄暫時數(1, 5) = 2 And 事件卡記錄暫時數(1, 6) = 1 Then
        事件卡記錄暫時數(1, 3) = 4
        事件卡.詛咒術_使用者 0, 0 '==事件卡執行_詛咒術_使用者(階段4)
   End If
End Sub

Private Sub 人物消失檢查_Timer()
If 人物消失檢查暫時變數(1) = 10 Then
    If 人物消失檢查暫時變數(2) = 1 Then
        personusminijpg.小人物消失 = True
    End If
    If 人物消失檢查暫時變數(3) = 1 Then
        personcomminijpg.小人物消失 = True
    End If
    人物消失檢查暫時變數(1) = Val(人物消失檢查暫時變數(1)) + 1
ElseIf Val(人物消失檢查暫時變數(1)) > 10 And personcomminijpg.小人物消失 = False And personusminijpg.小人物消失 = False Then
    人物消失檢查.Enabled = False
    FormMainMode.等待時間.Enabled = True
Else
    人物消失檢查暫時變數(1) = Val(人物消失檢查暫時變數(1)) + 1
End If
End Sub

Private Sub 小人物頭像移動_使用者_Timer()
Dim pnm As Integer
If 顯示列1.使用者方小人物圖片width > 1440 Then
    pnm = 0
Else
    pnm = 1440 - 顯示列1.使用者方小人物圖片width
End If
Select Case 小人物頭像移動方向數(1)
    Case 1
        If 顯示列1.使用者方小人物圖片left >= pnm Then
           顯示列1.使用者方小人物圖片left = pnm
           戰鬥系統類.小人物頭像執行完判斷_使用者
           小人物頭像移動_使用者.Enabled = False
           Exit Sub
        End If
           顯示列1.使用者方小人物圖片left = 顯示列1.使用者方小人物圖片left + 100
        If 顯示列1.使用者方小人物圖片left >= pnm Then
           顯示列1.使用者方小人物圖片left = pnm
           小人物頭像移動_使用者.Enabled = False
           戰鬥系統類.小人物頭像執行完判斷_使用者
        End If
    Case 2
        If 顯示列1.使用者方小人物圖片left <= -顯示列1.使用者方小人物圖片width Then
           顯示列1.使用者方小人物圖片left = -顯示列1.使用者方小人物圖片width
           小人物頭像移動_使用者.Enabled = False
           Exit Sub
        End If
           顯示列1.使用者方小人物圖片left = 顯示列1.使用者方小人物圖片left - 100
        If 顯示列1.使用者方小人物圖片left <= -顯示列1.使用者方小人物圖片width Then
           顯示列1.使用者方小人物圖片left = -顯示列1.使用者方小人物圖片width
           小人物頭像移動_使用者.Enabled = False
        End If
End Select
End Sub

Private Sub 小人物頭像移動_電腦_Timer()
Dim pnm As Integer
If 顯示列1.電腦方小人物圖片width > 1440 Then
    pnm = FormMainMode.ScaleWidth - 顯示列1.電腦方小人物圖片width
Else
    pnm = FormMainMode.ScaleWidth - 1440
End If
Select Case 小人物頭像移動方向數(2)
    Case 1
        If 顯示列1.電腦方小人物圖片left <= pnm Then
           顯示列1.電腦方小人物圖片left = pnm
           戰鬥系統類.小人物頭像執行完判斷_電腦
           小人物頭像移動_電腦.Enabled = False
           Exit Sub
        End If
           顯示列1.電腦方小人物圖片left = 顯示列1.電腦方小人物圖片left - 100
        If 顯示列1.電腦方小人物圖片left <= pnm Then
           顯示列1.電腦方小人物圖片left = pnm
           小人物頭像移動_電腦.Enabled = False
           戰鬥系統類.小人物頭像執行完判斷_電腦
        End If
    Case 2
        If 顯示列1.電腦方小人物圖片left >= FormMainMode.ScaleWidth Then
           顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
           小人物頭像移動_電腦.Enabled = False
           Exit Sub
        End If
           顯示列1.電腦方小人物圖片left = 顯示列1.電腦方小人物圖片left + 100
        If 顯示列1.電腦方小人物圖片left >= FormMainMode.ScaleWidth Then
           顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
           小人物頭像移動_電腦.Enabled = False
        End If
End Select
End Sub

Private Sub 收牌階段_計算_Timer()
Select Case 目前數(10)
    Case 1
       戰鬥系統類.收牌計算距離單位_使用者
       收牌階段_計算.Enabled = False
       目前數(11) = 0
       目前數(12) = pageqlead(目前數(10)) - 1
       牌移動_收牌.Enabled = True
    Case 2
       戰鬥系統類.收牌計算距離單位_電腦
       收牌階段_計算.Enabled = False
       目前數(11) = 0
       目前數(12) = pageqlead(目前數(10)) - 1
       牌移動_收牌.Enabled = True
    Case 3
       收牌階段_計算.Enabled = False
       Select Case turnatk
          Case 1
             戰鬥系統類.雙方HP檢查
          Case 2
             戰鬥系統類.雙方HP檢查
          Case 3
             '===========================執行階段插入點(8)
              執行階段系統類.執行階段系統總主要程序_執行階段開始 moveturn, 8, 1
            '============================
             HP檢查階段數 = 1
             戰鬥系統類.雙方HP檢查
       End Select
End Select
End Sub

Private Sub 血量載入動畫_Timer()
If 血量計數器動畫暫時變數(1, 2) = 0 Then
    If bloodlineout1.Width >= 5295 Then
        血量計數器動畫暫時變數(1, 2) = 1
    ElseIf 5295 - bloodlineout1.Width <= 106 Then
        血量計數器動畫暫時變數(1, 1) = 5295 - bloodlineout1.Width
        bloodlineout1.Width = bloodlineout1.Width + 血量計數器動畫暫時變數(1, 1)
        血量計數器動畫暫時變數(1, 2) = 1
    Else
       bloodlineout1.Width = bloodlineout1.Width + 血量計數器動畫暫時變數(1, 1)
    End If
End If
If 血量計數器動畫暫時變數(2, 2) = 0 Then
    If bloodlineout2.Left <= 6060 Then
        血量計數器動畫暫時變數(2, 2) = 1
    ElseIf bloodlineout2.Left - 6060 <= 106 Then
        血量計數器動畫暫時變數(2, 1) = bloodlineout2.Left - 6060
        bloodlineout2.Left = bloodlineout2.Left - 血量計數器動畫暫時變數(2, 1)
        血量計數器動畫暫時變數(2, 2) = 1
    Else
        bloodlineout2.Left = bloodlineout2.Left - 血量計數器動畫暫時變數(2, 1)
    End If
End If
If 血量計數器動畫暫時變數(1, 2) = 1 And 血量計數器動畫暫時變數(2, 2) = 1 Then
   血量載入動畫.Enabled = False
   等待時間佇列(2).Add 1
   等待時間_2.Enabled = True
End If
End Sub


Private Sub 攻擊階段_階段1_Timer()
'======================
If 電腦方事件卡是否出完選擇數 = True Then
    GoTo 電腦方事件卡先出制度_執行階段結束
End If
'======================電腦方事件卡先出制度
If 電腦方事件卡是否出完選擇數 = False Then
    GoTo 電腦方事件卡先出制度_執行階段2
End If
'================================
電腦方事件卡先出制度_執行階段結束:
'====================試驗智慧型AI出牌系統
Dim wtyr As Integer '暫時變數
If moveturn = 2 Then wtyr = 1 Else wtyr = 0
智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 2, 2, namecom(角色人物對戰人數(2, 2)), movecp, wtyr
'================
電腦方事件卡先出制度_執行階段2:
'================
攻擊階段_階段1.Enabled = False
If 電腦方事件卡是否出完選擇數 = False Then
    目前數(6) = 0
    目前數(17) = 1
    目前數(15) = 0
    小人物頭像移動方向數(1) = 2
    小人物頭像移動方向數(2) = 1
    小人物頭像移動_使用者.Enabled = True
    小人物頭像移動_電腦.Enabled = True
End If
'======================電腦方事件卡先出制度_結束後階段2
If 電腦方事件卡是否出完選擇數 = True Then
    電腦出牌.Enabled = True
End If
'===========================
End Sub

Private Sub 攻擊階段_階段2_Timer()
'----------以下為攻擊模式程序
擲骰表單溝通暫時變數(2) = 0
擲骰表單溝通暫時變數(3) = 0
擲骰表單溝通暫時變數(5) = 0
擲骰表單溝通暫時變數(6) = 0
擲骰表單溝通暫時變數(7) = 0
擲骰表單溝通暫時變數(8) = 0
'==============================
HP檢查變數 = False
'==============================
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'===========================執行階段插入點(ATK-10/DEF-30)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 10, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 30, 2
'============================
'===========================執行階段插入點(ATK-11/DEF-31)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 11, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 31, 2
'============================
'=================
If 攻擊防禦骰子總數(1) <= 0 Then
  戰鬥系統類.廣播訊息 "沒有攻擊。"
  戰鬥系統類.廣播訊息 "您取消了攻擊。"
  骰數零檢查值(1) = True
Else
  戰鬥系統類.廣播訊息 "決定攻擊力" & 攻擊防禦骰子總數(1) & "點。"
End If
If 攻擊防禦骰子總數(2) <= 0 Then
   骰數零檢查值(2) = True
End If
'===========================執行階段插入點(ATK-12/DEF-32)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 12, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 32, 2
'============================
階段狀態數 = 2
攻擊階段_階段2.Enabled = False
'============================
HP檢查變數 = True
HP檢查階段數 = 2
目前數(10) = 1
收牌階段_計算.Enabled = True
End Sub

Private Sub 攻擊階段_階段初始_Timer()
戰鬥系統類.時間軸_重設
trtimeline.Enabled = True
電腦方事件卡是否出完選擇數 = False
'==============================
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'==============================
攻擊階段_階段初始.Enabled = False
攻擊階段_階段1.Enabled = True
End Sub

Private Sub 防禦階段_階段初始_Timer()
'----------以下為防禦模式程序
擲骰表單溝通暫時變數(2) = 0
擲骰表單溝通暫時變數(3) = 0
擲骰表單溝通暫時變數(5) = 0
擲骰表單溝通暫時變數(6) = 0
擲骰表單溝通暫時變數(7) = 0
擲骰表單溝通暫時變數(8) = 0
'====================
HP檢查變數 = False
'==============================
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'===========================執行階段插入點(ATK-10/DEF-30)
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 10, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 30, 2
'============================
'===========================執行階段插入點(ATK-11/DEF-31)
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 11, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 31, 2
'============================
If 攻擊防禦骰子總數(2) <= 0 Then
  戰鬥系統類.廣播訊息 "沒有攻擊。"
  戰鬥系統類.廣播訊息 "您的對手取消了攻擊。"
  骰數零檢查值(2) = True
Else
  戰鬥系統類.廣播訊息 "決定攻擊力" & 攻擊防禦骰子總數(2) & "點。"
End If
If 攻擊防禦骰子總數(1) <= 0 Then
   骰數零檢查值(1) = True
End If
'===========================執行階段插入點(ATK-12/DEF-32)
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 12, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 32, 2
'============================
階段狀態數 = 4
防禦階段_階段初始.Enabled = False
'============================
HP檢查變數 = True
HP檢查階段數 = 2
目前數(10) = 1
收牌階段_計算.Enabled = True
End Sub

Sub 使用者出牌_AI出牌控制_Timer()
Dim tmpcard As clsActionCard
If turnpageonin = 1 And 牌移動.Enabled = False Then
    If 戰鬥系統類.CardDeckCollection(5).Count > 0 Then
        Set tmpcard = 戰鬥系統類.CardDeckCollection(5)(目前數(32))
        If tmpcard.ComMark = 3 Then
            目前數(32) = 目前數(32) - 1
            FormMainMode.card_CardClick tmpcard.CardNum
        End If
        目前數(32) = 目前數(32) + 1
    End If
    If 目前數(32) > 戰鬥系統類.CardDeckCollection(5).Count Then
        使用者出牌_AI出牌控制.Enabled = False
        等待時間佇列(1).Add 37
        等待時間.Enabled = True
    End If
End If
End Sub

Sub 使用者出牌_AI出牌控制_事件卡_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

If turnpageonin = 1 And 牌移動.Enabled = False Then
    If 戰鬥系統類.CardDeckCollection(5).Count > 0 Then
        For i = 1 To 戰鬥系統類.CardDeckCollection(5).Count
            Set tmpcard = 戰鬥系統類.CardDeckCollection(5)(i)
            If tmpcard.CardType = 2 Then
                If tmpcard.UpperType = a6a Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
                If tmpcard.UpperType = a7a And (turnatk = 1 Or turnatk = 2) Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
                If tmpcard.UpperType = a8a Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
                If tmpcard.UpperType = a9a Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
            End If
        Next
    End If
    
    使用者出牌_AI出牌控制_事件卡.Enabled = False
    等待時間佇列(2).Add 46
    等待時間_2.Enabled = True
End If
End Sub


Private Sub 使用者出牌_手牌對齊_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To 戰鬥系統類.CardDeckCollection(5).Count
    If i >= 目前數(5) Then
        Set tmpcard = 戰鬥系統類.CardDeckCollection(5)(i)
        If 目前數(13) = 0 Then
            If card(tmpcard.CardNum).Left = 2640 And card(tmpcard.CardNum).Top = 7980 Then  '指定第2列第1張牌
                目前數(13) = tmpcard.CardNum
                tmpcard.XYLeft = card(目前數(13)).Left  '指定目前Left(座標)
                tmpcard.XYTop = card(目前數(13)).Top  '指定目前Top(座標)
                '==========戰鬥系統類.計算牌移動距離單位
                距離單位_收牌暫時數(1, 1) = (9840 - tmpcard.XYLeft) \ 10 '計算Left
                距離單位_收牌暫時數(1, 2) = -((tmpcard.XYTop - 6700) \ 10)  '計算Top
            End If
        End If
        If 目前數(13) = tmpcard.CardNum Then
           card(目前數(13)).Left = card(目前數(13)).Left + 距離單位_收牌暫時數(1, 1)
           card(目前數(13)).Top = card(目前數(13)).Top + 距離單位_收牌暫時數(1, 2)
        Else
           card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (900 / 10)
        End If
  End If
Next
目前數(4) = 目前數(4) + (900 / 10)
If 目前數(4) >= 900 Then
    使用者出牌_手牌對齊.Enabled = False
    Select Case 目前數(21)
        Case 1
            '======結束動作
        Case 2
             If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 2 '(階段2)
            End If
       Case 3
           '===========事件卡執行_詛咒術_電腦(階段3)
            事件卡記錄暫時數(2, 3) = 3
            事件卡.詛咒術_電腦 0, 0
       Case 4
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards")) = 2 '(階段2)
            End If
       Case 5
             If 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards")) = 2 '(階段2)
            End If
        Case 11
            等待時間佇列(2).Add 38
            等待時間_2.Enabled = True
    End Select
End If
End Sub

Private Sub 使用者出牌_出牌對齊_靠右_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To 戰鬥系統類.CardDeckCollection(6).Count
    Set tmpcard = 戰鬥系統類.CardDeckCollection(6)(i)
    If i < 目前數(5) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left + (480 / 10)
    End If
    If i >= 目前數(5) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (500 / 10)
    End If
Next
目前數(3) = 目前數(3) + (480 / 10)
If 目前數(3) >= 480 Then
    使用者出牌_出牌對齊_靠右.Enabled = False
End If
End Sub

Private Sub 使用者出牌_出牌對齊_靠左_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To (戰鬥系統類.CardDeckCollection(6).Count - 1)
    Set tmpcard = 戰鬥系統類.CardDeckCollection(6)(i)
    card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (480 / 10)
Next
目前數(3) = 目前數(3) + (480 / 10)
If 目前數(3) >= 480 Then
    使用者出牌_出牌對齊_靠左.Enabled = False
End If
End Sub

Private Sub 移動階段_階段初始_Timer()
If 目前數(31) = 0 Then
    Dim movecpn As Integer, mfd As Integer
    movecpn = movecp
    '===============
    movecom = atkingpagetot(2, 3)
    moveus = atkingpagetot(1, 3)
    Erase Vss_PersonMoveActionChangeNum
    Erase Vss_PersonMoveControlNum
    Vss_PersonAttackFirstControlNum = 0
    '===========================執行階段插入點(2)
    戰鬥系統類.移動階段移動前執行階段呼叫 2
    '===========================執行階段插入點(3)
    戰鬥系統類.移動階段移動前執行階段呼叫 3
    '===========================執行階段插入點(4)
    戰鬥系統類.移動階段移動前執行階段呼叫 4
    '===========================執行階段插入點(70)
    戰鬥系統類.移動階段移動前執行階段呼叫 70
    '============================
    If Vss_PersonMoveControlNum(1, 2) = 0 Then
        moveus = moveus + Vss_PersonMoveControlNum(1, 1)
    Else
        moveus = Vss_PersonMoveControlNum(1, 1)
    End If
    If Vss_PersonMoveControlNum(2, 2) = 0 Then
        movecom = movecom + Vss_PersonMoveControlNum(2, 1)
    Else
        movecom = Vss_PersonMoveControlNum(2, 1)
    End If
    '==================================
    If moveus < 0 Then moveus = 0
    If movecom < 0 Then movecom = 0
    '==================================
    movecheckus = moveus
    movecheckcom = movecom
    顯示列1.電腦方移動值 = movecheckcom
    '----------以下為電腦判斷出牌程式碼（移動階段2）
    If movecheckcom <= 0 Then
       電腦方移動階段選擇數 = 2
    End If
    '==================================
    If Vss_PersonMoveActionChangeNum(1, 1) = 1 Then
        顯示列1.移動階段選擇值 = Vss_PersonMoveActionChangeNum(1, 2)
    End If
    If Vss_PersonMoveActionChangeNum(2, 1) = 1 Then
        電腦方移動階段選擇數 = Vss_PersonMoveActionChangeNum(2, 2)
    End If
    '===============
    If Vss_EventPlayerAllActionOffNum(1) = 1 Then 顯示列1.移動階段選擇值 = 0
    If Vss_EventPlayerAllActionOffNum(2) = 1 Then 電腦方移動階段選擇數 = 0
    '==================================
    ReDim VBEStageNum(0 To 4) As Integer
    VBEStageNum(0) = 71
    VBEStageNum(1) = moveus '使用者方總移動數
    VBEStageNum(2) = movecom '電腦方總移動數
    VBEStageNum(3) = 顯示列1.移動階段選擇值 '使用者方目前移動階段行動選擇
    VBEStageNum(4) = 電腦方移動階段選擇數 '電腦方目前移動階段行動選擇
    '===========================執行階段插入點(71)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 71, 1
    '============================
    If 顯示列1.移動階段選擇值 = 1 Or 顯示列1.移動階段選擇值 = 3 Then
       If 顯示列1.移動階段選擇值 = 3 Then
          moveus = -Val(moveus)
          顯示列1.使用者方移動內外 = 1
       ElseIf 顯示列1.移動階段選擇值 = 1 Then
          顯示列1.使用者方移動內外 = 2
       End If
     顯示列1.使用者方移動值 = movecheckus
    End If
    '========
    If 電腦方移動階段選擇數 = 1 Or 電腦方移動階段選擇數 = 3 Then
       If 電腦方移動階段選擇數 = 3 Then
          movecom = -Val(movecom)
          顯示列1.電腦方移動內外 = 1
       ElseIf 電腦方移動階段選擇數 = 1 Then
          顯示列1.電腦方移動內外 = 2
       End If
       顯示列1.電腦方移動值 = movecheckcom
    ElseIf 電腦方移動階段選擇數 = 2 Then
        If livecom(角色人物對戰人數(2, 2)) < livecommax(角色人物對戰人數(2, 2)) Then
            回復執行_電腦 1, 1, 0, True, True
        End If
        顯示列1.電腦方移動值 = 0
    ElseIf 電腦方移動階段選擇數 = 4 Then
        顯示列1.電腦方移動值 = 0
        交換角色紀錄暫時變數(2) = 1
    ElseIf 電腦方移動階段選擇數 = 0 Then
        顯示列1.電腦方移動值 = 0
    End If
    '==============================
    If 顯示列1.移動階段選擇值 = 2 Then
         回復執行_使用者 1, 1, 0, True, True
         顯示列1.使用者方移動值 = 0
    ElseIf 顯示列1.移動階段選擇值 = 0 Then
      顯示列1.使用者方移動值 = 0
    ElseIf 顯示列1.移動階段選擇值 = 4 Then
      顯示列1.使用者方移動值 = 0
      交換角色紀錄暫時變數(1) = 1
    End If
    '==============================
    If (顯示列1.移動階段選擇值 = 1 Or 顯示列1.移動階段選擇值 = 3) Then
        movecpn = Val(moveus) + Val(movecpn)
    End If
    If (電腦方移動階段選擇數 = 1 Or 電腦方移動階段選擇數 = 3) Then
        movecpn = Val(movecom) + Val(movecpn)
    End If
    '==============================
    
    If movecpn < 1 Then
       movecpn = 1
    ElseIf movecpn > 3 Then
       movecpn = 3
    End If
    
    執行動作_距離變更 movecpn, True, True
    
    If Vss_PersonAttackFirstControlNum = 1 Then
        戰鬥系統類.movetnus
    ElseIf Vss_PersonAttackFirstControlNum = 2 Then
        戰鬥系統類.movetncom
    Else
        If Val(movecheckus) > Val(movecheckcom) Then
            戰鬥系統類.movetnus
        ElseIf Val(movecheckus) < Val(movecheckcom) Then
            戰鬥系統類.movetncom
        Else
            Randomize
            mfd = Int(Rnd() * 2) + 1
            If mfd = 1 Then 戰鬥系統類.movetnus
            If mfd = 2 Then 戰鬥系統類.movetncom
        End If
    End If
    擲骰表單溝通暫時變數(4) = moveturn
    HP檢查變數 = False
    等待時間佇列(2).Add 23
    FormMainMode.等待時間_2.Enabled = True
Else
    '===========================執行階段插入點(5)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 moveturn, 5, 1
    '============================
    '===========================執行階段插入點(6)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 moveturn, 6, 1
    '============================
    '===========================執行階段插入點(7)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 moveturn, 7, 1
    '============================
    目前數(6) = 0
    目前數(10) = 1
    階段狀態數 = 2
    電腦出牌_亮牌.Enabled = True
End If
移動階段_階段初始.Enabled = False
End Sub

Private Sub 移動圖片完成檢查_Timer()
If 顯示列1.移動方向圖片顯示 = False Then
   收牌階段_計算.Enabled = True
   移動圖片完成檢查.Enabled = False
   FormMainMode.PEAFInterface.BnOKVisable False
End If
End Sub

Private Sub 牌移動_Timer()
Dim i As Integer

card(牌移動暫時變數(3)).Left = card(牌移動暫時變數(3)).Left + 距離單位(2, 1, 1)
card(牌移動暫時變數(3)).Top = card(牌移動暫時變數(3)).Top + 距離單位(2, 1, 2)
If Abs(牌移動暫時變數(1) - card(牌移動暫時變數(3)).Left) <= 50 Or Abs(牌移動暫時變數(2) - card(牌移動暫時變數(3)).Top) <= 50 Then
   card(牌移動暫時變數(3)).Left = 牌移動暫時變數(1)
   card(牌移動暫時變數(3)).Top = 牌移動暫時變數(2)
   card(牌移動暫時變數(3)).ZOrder
   For i = 1 To 3
       FormMainMode.PEAFpersoncardcom(i).ZOrder
   Next
   FormMainMode.PEAFAnimateInterface.ZOrder
   牌移動.Enabled = False
   Select Case 目前數(15)
        Case 1
            發牌檢查.Enabled = True
        Case 2
            目前數(8) = 0
            電腦出牌_手牌對齊.Enabled = True
        Case 3
            'Nothing
        Case 4
            card(目前數(20)).Visible = False
            目前數(4) = 0
            目前數(13) = 0
            使用者出牌_手牌對齊.Enabled = True
        Case 5
            card(目前數(16)).Visible = False
            目前數(8) = 0
            電腦出牌_手牌對齊.Enabled = True
        Case 6
            '===========事件卡執行_機會_使用者(階段2)
            card(事件卡記錄暫時數(1, 4)).Visible = False
            等待時間佇列(2).Add 6
            等待時間_2.Enabled = True
        Case 7
             '===========事件卡執行_機會_使用者(階段1)
            等待時間佇列(2).Add 5
            等待時間_2.Enabled = True
        Case 8
            '===========事件卡執行_機會_使用者(階段3)
            事件卡記錄暫時數(1, 3) = 3
            事件卡.機會_使用者 0, 0
        Case 9
             '===========事件卡執行_機會_電腦(階段1)
            等待時間佇列(2).Add 7
            等待時間_2.Enabled = True
        Case 10
            '===========事件卡執行_機會_電腦(階段3)
            card(事件卡記錄暫時數(2, 4)).Visible = False
            等待時間佇列(2).Add 8
            等待時間_2.Enabled = True
        Case 11
            '===========事件卡執行_機會_電腦(階段4)
            事件卡記錄暫時數(2, 3) = 4
            事件卡.機會_電腦 0, 0
        Case 12
            '===========事件卡執行_詛咒術_使用者(階段1)
            等待時間佇列(2).Add 11
            等待時間_2.Enabled = True
        Case 13
            '===========事件卡執行_詛咒術_使用者(階段6)
            card(事件卡記錄暫時數(1, 4)).Visible = False
            事件卡記錄暫時數(1, 3) = 6
            事件卡.詛咒術_使用者 0, 0
        Case 14
            '===========事件卡執行_詛咒術_電腦(階段1)
            等待時間佇列(2).Add 13
            等待時間_2.Enabled = True
        Case 15
            '===========事件卡執行_詛咒術_電腦(階段5>6)
            card(事件卡記錄暫時數(2, 4)).Visible = False
            事件卡記錄暫時數(2, 3) = 6
            事件卡.詛咒術_電腦 0, 0
        Case 16
            '===========事件卡執行_HP回復_使用者(階段1)
            等待時間佇列(2).Add 16
            等待時間_2.Enabled = True
            turnpageonin = 0
            FormMainMode.PEAFInterface.BnOKEnabled False
        Case 17
            '===========事件卡執行_HP回復_使用者(階段4)
            card(事件卡記錄暫時數(1, 4)).Visible = False
            事件卡記錄暫時數(1, 3) = 4
            事件卡.HP回復_使用者 0, 0
        Case 18
            '===========事件卡執行_HP回復_電腦(階段1)
            等待時間佇列(2).Add 18
            等待時間_2.Enabled = True
        Case 19
            '===========事件卡執行_HP回復_電腦(階段4>5)
            card(事件卡記錄暫時數(2, 4)).Visible = False
            事件卡記錄暫時數(2, 3) = 5
            事件卡.HP回復_電腦 0, 0
        Case 20
            目前數(4) = 0
            目前數(13) = 0
            使用者出牌_手牌對齊.Enabled = True
        Case 21
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDrawCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDrawCards")) = 2 '(階段2)
            End If
        Case 22
           If 執行階段系統_搜尋正在執行之執行階段("AtkingGetUsedCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGetUsedCards")) = 2 '(階段2)
            End If
        Case 23
            If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 3 '(階段3)
            End If
        Case 40
            等待時間佇列(2).Add 37
            等待時間_2.Enabled = True
        Case 41
            '===========事件卡執行_聖水_使用者(階段1)
            等待時間佇列(2).Add 39
            等待時間_2.Enabled = True
            turnpageonin = 0
            FormMainMode.PEAFInterface.BnOKEnabled False
        Case 42
            '===========事件卡執行_聖水_使用者(階段4>5)
            card(事件卡記錄暫時數(1, 4)).Visible = False
            事件卡記錄暫時數(1, 3) = 4
            事件卡.聖水_使用者 0, 0
        Case 43
            '===========事件卡執行_聖水_電腦(階段1)
            等待時間佇列(2).Add 41
            等待時間_2.Enabled = True
        Case 44
            '===========事件卡執行_聖水_電腦(階段4>5)
            card(事件卡記錄暫時數(2, 4)).Visible = False
            事件卡記錄暫時數(2, 3) = 5
            事件卡.聖水_電腦 0, 0
   End Select
End If
End Sub


Private Sub 牌移動_收牌_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

If 目前數(11) = pageqlead(目前數(10)) Then
    戰鬥系統類.checkpage
    牌移動_收牌.Enabled = False
    目前數(10) = 目前數(10) + 1
    收牌階段_計算.Enabled = True
    Exit Sub
End If
For i = 1 + 目前數(11) To pageqlead(目前數(10)) - 目前數(12)
    If Abs(240 - card(距離單位_收牌暫時數(i, 3)).Left) <= 10 Or Abs(960 - card(距離單位_收牌暫時數(i, 3)).Top) <= 10 Then
        card(距離單位_收牌暫時數(i, 3)).Left = 240
        card(距離單位_收牌暫時數(i, 3)).Top = 960
        card(距離單位_收牌暫時數(i, 3)).Visible = False
        
        Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(距離單位_收牌暫時數(i, 3))))(CStr(距離單位_收牌暫時數(i, 3)))
        tmpcard.Location = 3
        Select Case tmpcard.CardType
            Case 1 '公用牌
                戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(距離單位_收牌暫時數(i, 3))), 2
            Case 2 '事件卡
                戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(距離單位_收牌暫時數(i, 3))), 9
        End Select
        
        目前數(11) = 目前數(11) + 1
    End If
    card(距離單位_收牌暫時數(i, 3)).Left = card(距離單位_收牌暫時數(i, 3)).Left + 距離單位_收牌暫時數(i, 1)
    card(距離單位_收牌暫時數(i, 3)).Top = card(距離單位_收牌暫時數(i, 3)).Top + 距離單位_收牌暫時數(i, 2)
    If 目前數(12) > 0 Then
        目前數(12) = 目前數(12) - 1
    End If
Next

End Sub

Private Sub 發牌_使用者階段_Timer()
發牌_使用者階段.Enabled = False
目前數(2) = 2

If Val(pageusglead) < 牌總階段數(1) Then
    戰鬥系統類.執行動作_抽牌_公用牌 1
Else
    發牌檢查.Enabled = True
End If
End Sub

Private Sub 發牌_電腦階段_Timer()
發牌_電腦階段.Enabled = False
目前數(2) = 3

If Val(pagecomglead) < 牌總階段數(2) Then
    戰鬥系統類.執行動作_抽牌_公用牌 2
Else
    發牌檢查.Enabled = True
End If
End Sub

Private Sub 發牌檢查_Timer()
If (Val(pageusglead) >= 牌總階段數(1) And Val(pagecomglead) >= 牌總階段數(2)) Or BattleCardNum <= 0 Then
   發牌檢查.Enabled = False
   目前數(15) = 0
   等待時間佇列(1).Add 3
   等待時間.Enabled = True
Else
   Select Case 目前數(2)
       Case 1
           發牌_使用者階段.Enabled = True
           發牌檢查.Enabled = False
       Case 2
           發牌_電腦階段.Enabled = True
           發牌檢查.Enabled = False
        Case 3
           目前數(2) = 1
    End Select
End If

End Sub

Private Sub 等待時間_2_Timer()
Select Case 目前數(14)
   Case 0
      目前數(14) = 目前數(14) + 1
   Case 1
      If 等待時間佇列(2).Count <= 1 Then
          目前數(14) = 0
          等待時間_2.Enabled = False
      End If
      If 等待時間佇列(2).Count = 0 Then Exit Sub
      Select Case 等待時間佇列(2).item(1)
          Case 1
              '========開始初始階段1
                顯示列1.Visible = True
                顯示列1.移動階段圖顯示 = False
                顯示列1.移動方向圖片顯示 = False
                一般系統類.音效播放 6
                If 系統顯示界面紀錄數 = 1 Then
                    draw1.Visible = False
                    draw2.Visible = True
                Else
                    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\draw2.gif"
                End If
                等待時間佇列(1).Add 2
                等待時間.Enabled = True
          Case 2
              cn22_Click
              FormMainMode.PEAFInterface.BnOKVisable False
           Case 3
              cn32_Click
              FormMainMode.PEAFInterface.BnOKVisable False
           Case 4
              Select Case turnatk
                    Case 1
                        等待時間佇列(1).Add 7
                        等待時間.Enabled = True
                    Case 2
                        等待時間佇列(1).Add 8
                        等待時間.Enabled = True
                    Case 3
                        cnmove2_Click
                End Select
           Case 5
                '===========事件卡執行_機會_使用者(階段1>2)
                事件卡記錄暫時數(1, 3) = 2
                事件卡.機會_使用者 0, 0
           Case 6
                '===========事件卡執行_機會_使用者(階段2>3)
                事件卡記錄暫時數(1, 3) = 3
                事件卡.機會_使用者 0, 0
           Case 7
                '===========事件卡執行_機會_電腦(階段1>2)
                事件卡記錄暫時數(2, 3) = 2
                事件卡.機會_電腦 0, 0
           Case 8
                '===========事件卡執行_機會_電腦(階段3>4)
                事件卡記錄暫時數(2, 3) = 4
                事件卡.機會_電腦 0, 0
            Case 9
                '===========事件卡執行_機會_電腦(階段2>3)
                事件卡記錄暫時數(2, 3) = 3
                事件卡.機會_電腦 0, 0
            Case 10
                電腦出牌.Enabled = True
            Case 11
                '===========事件卡執行_詛咒術_使用者(階段1>2)
                事件卡記錄暫時數(1, 3) = 2
                事件卡.詛咒術_使用者 0, 0
            Case 12
                '===========事件卡執行_詛咒術_使用者(階段>5)
                事件卡記錄暫時數(1, 3) = 5
                事件卡.詛咒術_使用者 0, 0
            Case 13
                '===========事件卡執行_詛咒術_電腦(階段1>2)
                事件卡記錄暫時數(2, 3) = 2
                事件卡.詛咒術_電腦 0, 0
            Case 14
                '===========事件卡執行_詛咒術_電腦(階段>4)
                事件卡記錄暫時數(2, 3) = 4
                事件卡.詛咒術_電腦 0, 0
            Case 15
                '===========事件卡執行_詛咒術_電腦(階段4>5)
                事件卡記錄暫時數(2, 3) = 5
                事件卡.詛咒術_電腦 0, 0
            Case 16
                '===========事件卡執行_HP回復_使用者(階段1>2)
                事件卡記錄暫時數(1, 3) = 2
                事件卡.HP回復_使用者 0, 0
            Case 17
                '===========事件卡執行_HP回復_使用者(階段2>3)
                事件卡記錄暫時數(1, 3) = 3
                事件卡.HP回復_使用者 0, 0
            Case 18
                '===========事件卡執行_HP回復_電腦(階段1>2)
                事件卡記錄暫時數(2, 3) = 2
                事件卡.HP回復_電腦 0, 0
            Case 19
                '===========事件卡執行_HP回復_電腦(階段2>3)
                事件卡記錄暫時數(2, 3) = 3
                事件卡.HP回復_電腦 0, 0
            Case 20
                '===========事件卡執行_HP回復_電腦(階段3>4)
                事件卡記錄暫時數(2, 3) = 4
                事件卡.HP回復_電腦 0, 0
            Case 21
                Select Case turnatk
                   Case 1
                       戰鬥系統類.執行動作_攻擊階段結束時技能啟動
                   Case 2
                       戰鬥系統類.執行動作_防禦階段結束時技能啟動
               End Select
            Case 22
               FormMainMode.骰子執行完啟動.Enabled = True
            Case 23
                目前數(31) = 1
                FormMainMode.移動階段_階段初始.Enabled = True
            Case 24
                If FormMainMode.PEAFDiceInterface.DiceStop = True Or 骰數零檢查值(1) = True Or 骰數零檢查值(2) = True Then
                    If 執行階段系統_搜尋正在執行之執行階段("BattleStartDice") <> 0 Then
                        vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("BattleStartDice")) = 2 '(階段2)
                    End If
                Else
                    等待時間佇列(2).Add 24
                    等待時間_2.Enabled = True
                End If
            Case 25
                If FormMainMode.PEAFDiceInterface.DiceStop = True Or 骰數零檢查值(1) = True Or 骰數零檢查值(2) = True Then
                    戰鬥系統類.擲骰後續判斷
                Else
                    等待時間佇列(2).Add 25
                    等待時間_2.Enabled = True
                End If
            Case 30
                If 電腦出牌_亮牌.Enabled = False Then
                    顯示列1.移動方向圖片顯示 = True
                    移動圖片完成檢查.Enabled = True
                Else
                    等待時間佇列(2).Add 30
                    等待時間_2.Enabled = True
                End If
            Case 39
                '===========事件卡執行_聖水_使用者(階段1>2)
                事件卡記錄暫時數(1, 3) = 2
                事件卡.聖水_使用者 0, 0
            Case 40
                '===========事件卡執行_聖水_使用者(階段2>3)
                事件卡記錄暫時數(1, 3) = 3
                事件卡.聖水_使用者 0, 0
            Case 41
                '===========事件卡執行_聖水_電腦(階段1>2)
                事件卡記錄暫時數(2, 3) = 2
                事件卡.聖水_電腦 0, 0
            Case 42
                '===========事件卡執行_聖水_電腦(階段2>3)
                事件卡記錄暫時數(2, 3) = 3
                事件卡.聖水_電腦 0, 0
            Case 43
                '===========事件卡執行_聖水_電腦(階段3>4)
                事件卡記錄暫時數(2, 3) = 4
                事件卡.聖水_電腦 0, 0
            Case 45
                目前數(32) = 1
                FormMainMode.使用者出牌_AI出牌控制_事件卡.Enabled = True
            Case 46
                '====================試驗智慧型AI出牌系統
                Dim wtyr As Integer '暫時變數
                If (moveturn = 1 And turnatk = 2) Or (moveturn = 2 And turnatk = 1) Then wtyr = 1 Else wtyr = 0
                智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 1, turnatk, nameus(角色人物對戰人數(1, 2)), movecp, wtyr
                智慧型AI系統類.智慧型AI系統_使用者出牌階段判斷反轉
                目前數(32) = 1
                FormMainMode.使用者出牌_AI出牌控制.Enabled = True
            Case 48
                執行動作_電腦方各階段出牌完畢後行動 turnatk
      End Select
      等待時間佇列(2).Remove 1
End Select
End Sub

Private Sub 等待時間_Timer()
Select Case 目前數(22)
    Case 0
       目前數(22) = 目前數(22) + 1
    Case 1
        If 等待時間佇列(1).Count <= 1 Then
            目前數(22) = 0
            等待時間.Enabled = False
        End If
        If 等待時間佇列(1).Count = 0 Then Exit Sub
        Select Case 等待時間佇列(1).item(1)
            Case 2   '========開始初始階段2
                等待時間佇列(1).Add 5
                等待時間.Enabled = True
            Case 3
                等待時間佇列(1).Add 22
                等待時間.Enabled = True
            Case 4
                戰鬥系統類.廣播訊息 "現在的距離" & movecp & "。"
                交換角色紀錄暫時變數(4) = 1
                戰鬥系統類.執行動作_移動階段選擇執行
            Case 5
                cn1_Click
            Case 6
                Erase Vss_EventPlayerAllActionOffNum
                '===========================執行階段插入點(1)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 1, 1
                '============================
                cnmove_Click
            Case 7
                等待時間佇列(2).Add 2
                等待時間_2.Enabled = True
            Case 8
                等待時間佇列(2).Add 3
                等待時間_2.Enabled = True
            Case 9
                cn2_Click
                顯示列1.Visible = True
                戰鬥系統類.時間軸_顯示
            Case 10
                電腦方事件卡是否出完選擇數 = False
                cn3_Click
                顯示列1.Visible = True
                戰鬥系統類.時間軸_顯示
            Case 11
                戰鬥系統類.時間軸_隱藏
                顯示列1.Visible = False
                等待時間佇列(1).Add 12
                等待時間.Enabled = True
            Case 12
                If Val(擲骰表單溝通暫時變數(4)) = 1 Then
                   Select Case Val(擲骰表單溝通暫時變數(1))
                    Case 1
                        '===========================執行階段插入點(ATK-13/DEF-33)
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 13, 2
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 33, 2
                        '============================
                    Case 2
                       '===========================執行階段插入點(ATK-13/DEF-33)
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 13, 2
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 33, 2
                        '============================
                    End Select
                Else
                   Select Case Val(擲骰表單溝通暫時變數(1))
                    Case 1
                       '===========================執行階段插入點(ATK-13/DEF-33)
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 13, 2
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 33, 2
                        '============================
                    Case 2
                       '===========================執行階段插入點(ATK-13/DEF-33)
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 13, 2
                        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 33, 2
                        '============================
                    End Select
                End If
                擲骰表單溝通暫時變數(9) = 攻擊防禦骰子總數(1)
                擲骰表單溝通暫時變數(10) = 攻擊防禦骰子總數(2)
                是否系統公骰 = True
                戰鬥系統類.擲骰表單顯示
                等待時間佇列(2).Add 25
                FormMainMode.等待時間_2.Enabled = True
            Case 13
                等待時間佇列(1).Add 9
                等待時間.Enabled = True
            Case 14
                等待時間佇列(1).Add 10
                等待時間.Enabled = True
            Case 15
                cn4_Click
            Case 17
                '===========================執行階段插入點(9)
                 執行階段系統類.執行階段系統總主要程序_執行階段開始 moveturn, 9, 1
                '============================
                Select Case moveturn
                  Case 1
                     cn2_Click
                  Case 2
                     電腦方事件卡是否出完選擇數 = False
                     cn3_Click
                End Select
            Case 18
                 戰鬥系統類.執行動作_交換人物角色_電腦_初始
            Case 19
                 戰鬥系統類.執行動作_交換人物角色_電腦_交換
            Case 20
                 戰鬥系統類.時間軸_隱藏
                 顯示列1.Visible = False
                 cn4_Click
            Case 21
                交換角色紀錄暫時變數(4) = 2
                執行動作_人物死亡交換階段選擇執行
            Case 22
                戰鬥系統類.事件卡處理_分派_使用者方
                戰鬥系統類.事件卡處理_分派_電腦方
                等待時間佇列(1).Add 6
                等待時間.Enabled = True
            Case 30
                電腦出牌.Enabled = True
            Case 36
                FormMainMode.trend.Enabled = True
            Case 37
                Dim ckl As Integer
                '=============使用者方選擇行動
                If turnatk = 3 Then
                    顯示列1.移動階段選擇值 = 目前數(33)
                End If
                '====================
                FormMainMode.PEAFInterface.BnOKStartListen
                FormMainMode.PEAFInterface_BnOKClick
                For ckl = 1 To 戰鬥系統類.ActionCardTotNum
                    FormMainMode.card(ckl).CardEnabledType = True
                Next
        End Select
        等待時間佇列(1).Remove 1
End Select
End Sub

Private Sub 電腦出牌_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

電腦出牌.Enabled = False
If 電腦方事件卡是否出完選擇數 = False Then
     '=========================專屬事件卡出牌階段
    For i = 1 To 戰鬥系統類.CardDeckCollection(7).Count
        Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(i)
        If tmpcard.CardType = 2 Then
            If tmpcard.UpperType = a6a Then
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a6a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            End If
            If tmpcard.UpperType = a7a And (turnatk = 1 Or turnatk = 2) Then
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a7a And (turnatk = 1 Or turnatk = 2) Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            End If
            If tmpcard.UpperType = a8a Then
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a8a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            End If
            If tmpcard.UpperType = a9a Then
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a9a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
                Exit Sub
            End If
        End If
    Next
    '==============================事件卡均已出牌完畢
    電腦方事件卡是否出完選擇數 = True
    Select Case turnatk
        Case 1
             攻擊階段_階段1.Enabled = True
        Case 2
             cn3_Click
        Case 3
             cnmove_Click
    End Select
    Exit Sub
End If
'===========================================
If 電腦方事件卡是否出完選擇數 = True Then
    Do
        目前數(6) = 目前數(6) + 1
        If 目前數(6) > 戰鬥系統類.CardDeckCollection(7).Count Then
            電腦方事件卡是否出完選擇數 = False
            Select Case turnatk
               Case 1
                  目前數(6) = 0
                  目前數(10) = 1
                  戰鬥系統類.時間軸_停止
                  電腦出牌_亮牌.Enabled = True
                  trgoi2_Timer
               Case 2
                  目前數(6) = 0
                  目前數(10) = 1
                  戰鬥系統類.時間軸_停止
                  電腦出牌_亮牌.Enabled = True
                  trgoi2_Timer
                  trgoi1_Timer
               Case 3
                    執行動作_電腦方各階段出牌完畢後行動 3
            End Select
            Exit Sub
        End If
        Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(目前數(6))
        If tmpcard.ComMark = 1 Then
            目前數(6) = 目前數(6) - 1
            戰鬥系統類.電腦牌_模擬按牌 tmpcard.CardNum
            Exit Sub
        End If
    Loop
End If
End Sub

Private Sub 電腦出牌_手牌對齊_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

If 目前數(8) < 240 Then
    For i = 1 To 戰鬥系統類.CardDeckCollection(7).Count
        Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(i)
        If i >= 目前數(9) Then
            card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left + (240 / 10)
        End If
    Next
    目前數(8) = 目前數(8) + (240 / 10)
Else
    電腦出牌_手牌對齊.Enabled = False
    Select Case 目前數(17)
        Case 1
            If 牌移動.Enabled = False Then
                電腦出牌.Enabled = True
            Else
                電腦出牌_手牌對齊.Enabled = True  '等待牌移動完畢
            End If
        Case 2
            '======結束動作
        Case 3
            If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 3 '(階段3)
            End If
        Case 4
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards")) = 3 '(階段3)
            End If
        Case 5
            If 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards")) = 3 '(階段3)
            End If
        Case 6
           '===========事件卡執行_詛咒術_使用者(階段3)
            事件卡記錄暫時數(1, 3) = 3
            事件卡.詛咒術_使用者 0, 0
    End Select
    
End If
End Sub


Private Sub 電腦出牌_出牌對齊_靠右_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To 戰鬥系統類.CardDeckCollection(8).Count
    Set tmpcard = 戰鬥系統類.CardDeckCollection(8)(i)
    If i < 目前數(9) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left + (480 / 10)
    End If
    If i >= 目前數(9) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (500 / 10)
    End If
Next
目前數(7) = 目前數(7) + (480 / 10)
If 目前數(7) >= 480 Then
    電腦出牌_出牌對齊_靠右.Enabled = False
End If
End Sub

Private Sub 電腦出牌_出牌對齊_靠左_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To (戰鬥系統類.CardDeckCollection(8).Count - 1)
    Set tmpcard = 戰鬥系統類.CardDeckCollection(8)(i)
    card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (480 / 10)
Next
目前數(7) = 目前數(7) + (480 / 10)
If 目前數(7) >= 480 Then
    電腦出牌_出牌對齊_靠左.Enabled = False
End If
End Sub


Private Sub 電腦出牌_亮牌_Timer()
Dim tmpcard As clsActionCard

目前數(6) = 目前數(6) + 1
If 目前數(6) > 戰鬥系統類.CardDeckCollection(8).Count Then
    電腦出牌_亮牌.Enabled = False
    Select Case turnatk
       Case 1, 2
            執行動作_電腦方各階段出牌完畢後行動 turnatk
       Case 3
            等待時間佇列(2).Add 30
            等待時間_2.Enabled = True
    End Select
    Exit Sub
End If

Set tmpcard = 戰鬥系統類.CardDeckCollection(8)(目前數(6))
戰鬥系統類.公用牌回復正面 tmpcard.CardNum
一般系統類.音效播放 4
End Sub

Private Sub 對齊完成檢查_Timer()
If 使用者出牌_出牌對齊_靠左.Enabled = False And 使用者出牌_出牌對齊_靠右.Enabled = False And 使用者出牌_手牌對齊.Enabled = False And 牌移動.Enabled = False Then
   turnpageonin = 1
   對齊完成檢查.Enabled = False
End If
End Sub

Private Sub 骰子執行完啟動_Timer()
Dim uscomvsn As Integer
骰子執行完啟動.Enabled = False
'===========================
If Val(擲骰表單溝通暫時變數(4)) = 1 Then
   Select Case Val(擲骰表單溝通暫時變數(1))
    Case 1
        uscomvsn = 1
    Case 2
        uscomvsn = 2
    End Select
Else
   Select Case Val(擲骰表單溝通暫時變數(1))
    Case 1
       uscomvsn = 2
    Case 2
       uscomvsn = 1
    End Select
End If
'===========================執行階段插入點(20)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 20, 1
'============================
'===========================執行階段插入點(21)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 21, 1
'============================
'===========================執行階段插入點(22)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 22, 1
'============================
'===========================執行階段插入點(23)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 23, 1
'============================
'===========================執行階段插入點(24)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 24, 1
'============================
'===========================執行階段插入點(25)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 25, 1
'============================
'===========================執行階段插入點(26)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 26, 1
'============================
'===========================執行階段插入點(27)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 27, 1
'============================
'===========================執行階段插入點(28)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 28, 1
'============================
'===========================執行階段插入點(29)
執行階段系統類.執行階段系統總主要程序_執行階段開始 uscomvsn, 29, 1
'============================
trnextend.Enabled = True
End Sub

Private Sub 影子設定_Click()
FormDevSetting.smallleftus.Caption = personusminijpg.小人物影子Left
FormDevSetting.smalltopus.Caption = personusminijpg.小人物影子top差
FormDevSetting.smallleftcom.Caption = personcomminijpg.小人物影子Left
FormDevSetting.smalltopcom.Caption = personcomminijpg.小人物影子top差
FormDevSetting.smallpnleftus.Caption = personusminijpg.Left
FormDevSetting.smallpntopus.Caption = personusminijpg.Top
FormDevSetting.smallpnleftcom.Caption = personcomminijpg.Left
FormDevSetting.smallpntopcom.Caption = personcomminijpg.Top
FormDevSetting.personfus.Caption = 顯示列1.使用者方小人物圖片left
FormDevSetting.personfcom.Caption = 顯示列1.電腦方小人物圖片left
If Formsetting.checktest.Value = 1 Then
    FormDevSetting.Height = 6825
ElseIf Formsetting.checktestpersondown.Value = 1 Then
    FormDevSetting.Height = 3075
End If
'=================
戰鬥系統類.時間軸_停止
'=================
FormDevSetting.Show 1
End Sub
Private Sub bnabout_Click()
FormAbout.Show 1
一般系統類.音效播放 11
End Sub

Private Sub bnconfig_Click()
Formsetting.Left = FormMainMode.Left + 915
Formsetting.Top = FormMainMode.Top + 300
一般系統類.自由戰鬥模式設定表單各式設定讀入程序
一般系統類.音效播放 11
Formsetting.Show 1
End Sub

Sub PEGFbnstart_Click()
PEGameFreeModeSettingForm.Enabled = False
一般系統類.開始遊戲進行程序
End Sub

Private Sub Form_Load()
'============
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'==============
Dim cmstr() As String
Dim i As Integer

cmstr = Split(Command$, "/")
If UBound(cmstr) > 0 Then
    For i = 0 To UBound(cmstr)
        If cmstr(i) = "wine" Then 一般系統類.ProgramIsOnWine = True
    Next
End If
一般系統類.判斷字型_FormMainMode
一般系統類.主選單_PEStartForm顯示
End Sub
Private Sub personreadifus_Click()
cdgpersonus.ShowOpen
人物系統類.卡片人物資訊讀入_初階段 cdgpersonus.filename
End Sub
Private Sub personlevelcom_Click(Index As Integer)
人物系統類.清除角色人物資訊變數 2, Index
'人物系統類.卡片人物資訊讀入_三階段_電腦 personnamecom(Index).Text, personlevelcom(Index).Text, Index, 2
人物系統類.卡片人物資訊讀入_三階段 personnamecom(Index).Text, personlevelcom(Index).Text, Index, 2
'人物系統類.卡片人物資訊讀入_四階段_電腦 personnamecom(Index).Text, Index   '基於Unlight官方無電腦方對話原則
'人物系統類.卡片人物資訊讀入_四階段 personnamecom(Index).Text, Index, 2 '基於Unlight官方無電腦方對話原則
人物系統類.卡片人物資訊顯示_電腦 Index
End Sub

Private Sub personlevelus_Click(Index As Integer)
人物系統類.清除角色人物資訊變數 1, Index
'人物系統類.卡片人物資訊讀入_三階段_使用者 personnameus(Index).Text, personlevelus(Index).Text, Index, 1
人物系統類.卡片人物資訊讀入_三階段 personnameus(Index).Text, personlevelus(Index).Text, Index, 1
'人物系統類.卡片人物資訊讀入_四階段_使用者 personnameus(Index).Text, Index
人物系統類.卡片人物資訊讀入_四階段 personnameus(Index).Text, Index, 1
人物系統類.卡片人物資訊顯示_使用者 Index
End Sub

Private Sub personnamecom_Click(Index As Integer)
If 選單電腦事件 = True Then
    更新人物清單_電腦方_變更 Index
    If personnamecom(Index).Text = "" Or personnamecom(Index).Text = "《隨機》" Then
       personlevelcom(Index).Clear
        人物系統類.角色隨機_電腦 Index
        人物系統類.卡片人物資訊顯示_電腦 Index
    Else
       卡片人物資訊讀入_二階段_電腦 personnamecom(Index).Text, Index
    End If
    personlevelcom(Index).ListIndex = personlevelcom(Index).ListCount - 1
End If
End Sub

Private Sub personnameus_Click(Index As Integer)
If 選單使用者事件 = True Then
    更新人物清單_使用者方_變更 Index
    If personnameus(Index).Text = "" Or personnameus(Index).Text = "《隨機》" Then
        personlevelus(Index).Clear
        人物系統類.角色隨機_使用者 Index
        人物系統類.卡片人物資訊顯示_使用者 Index
    Else
        卡片人物資訊讀入_二階段_使用者 personnameus(Index).Text, Index
    End If
    personlevelus(Index).ListIndex = personlevelus(Index).ListCount - 1
End If
End Sub

Private Sub personresetcom_Click(Index As Integer)
personnamecom(Index).ListIndex = -1
personlevelcom(Index).Clear
End Sub

Private Sub personresetus_Click(Index As Integer)
personnameus(Index).ListIndex = -1
personlevelus(Index).Clear
End Sub
Private Sub start1_Timer()
Dim i As Integer
If st > 200 Then
   stup.Enabled = True
   stdown.Enabled = True
   start1.Enabled = False
   start2.Enabled = True
   For i = 1 To 3
      If PEASusbi1(i).Caption = "0" Then
         PEAScardus(i).Visible = False
         cardusname(i).Visible = False
         cardusspname(i).Visible = False
         Formchangeperson.card(i - 1).Visible = False
         Formchangeperson.bnok(i - 1).Visible = False
      Else
         PEAScardus(i).Visible = True
      End If
      If PEAScardcompi1(i).Caption = "0" Then
         PEAScardcom(i).Visible = False
         cardcomname(i).Visible = False
         cardcomspname(i).Visible = False
      Else
         PEAScardcom(i).Visible = True
      End If
   Next
   If Formsetting.chkpersonvsmode.Value = 1 Then
       For i = 2 To 3
           PEAScardcompi1(i).Caption = "?"
           PEAScardcompi2(i).Caption = "?"
           PEAScardcompi3(i).Caption = "?"
           PEAScardcom(i).Picture = LoadPicture(app_path & "gif\system\personunknown.jpg")
           cardcomname(i).Caption = "UnKnown"
           cardcomspname(i).Visible = False
        Next
    End If
    '==============
   downjpg.Visible = True
   upjpg_2.Visible = True
   開始卡片移動動畫完成數(1, 4) = 角色人物對戰人數(1, 1)
   開始卡片移動動畫完成數(2, 4) = 角色人物對戰人數(2, 1)
Else
  st = Val(st) + 1
End If
End Sub

Private Sub start2_Timer()
If sq = 401 Then
   tr大人物形像_使用者.Enabled = True
   tr大人物形像_電腦.Enabled = True
   sq = Val(sq) + 1
ElseIf sq = 500 Then
   一般系統類.主選單_PEAttackingForm顯示
   PEAttackingStartForm.Visible = False
   start2.Enabled = False
   FormMainMode.血量載入動畫.Enabled = True
Else
   sq = Val(sq) + 1
End If
   
End Sub

Private Sub stdown_Timer()
If sq <= 400 Then
   If downjpg.Top <= 8400 Then
      downjpg.Top = 8400
      stdown.Enabled = False
      cardustr.Enabled = True
      cardcomtr.Enabled = True
   Else
      downjpg.Top = Val(downjpg.Top) - 60
   End If
Else
   If downjpg.Top >= Val(FormMainMode.Height) Then
      downjpg.Top = Val(FormMainMode.Height)
      stdown.Enabled = False
   Else
      downjpg.Top = Val(downjpg.Top) + 60
   End If
End If
End Sub

Private Sub stup_Timer()
If sq <= 400 Then
   If upjpg_2.Top >= 0 Then
      upjpg_2.Top = 0
      stup.Enabled = False
   Else
      upjpg_2.Top = Val(upjpg_2.Top) + 60
   End If
Else
   If upjpg_2.Top <= -Val(upjpg_2.Height) Then
      upjpg_2.Top = -Val(upjpg_2.Height)
      PEASpersontalk.Top = -Val(PEASpersontalk.Height)
      stup.Enabled = False
   Else
      upjpg_2.Top = Val(upjpg_2.Top) - 60
      PEASpersontalk.Top = Val(PEASpersontalk.Top) - 60
   End If
End If
End Sub

Private Sub tr大人物形像_使用者_Timer()
Dim bigall As Integer
Dim bigw As Integer
Dim kp As Integer

bigw = 大人物形像_使用者.大人物圖片width / 2
If 2580 - bigw < 0 Or Val(VBEPerson(1, 1, 2, 2, 5)) = 1 Then
    bigall = 0
Else
    bigall = 2580 - bigw
End If

kp = (大人物形像_使用者.大人物圖片width + bigall) / 30
If sq <= 400 Then
   If 大人物形像_使用者.Left >= bigall Then
       大人物形像_使用者.Left = bigall
       tr大人物形像_使用者.Enabled = False
       swq = 0
       PEASpke.Enabled = True
   Else
       If Abs(大人物形像_使用者.Left - bigall) < kp Then
          大人物形像_使用者.Left = 大人物形像_使用者.Left + Abs(大人物形像_使用者.Left - bigall)
       Else
          大人物形像_使用者.Left = 大人物形像_使用者.Left + kp
       End If
   End If
Else
   If 大人物形像_使用者.Left <= -大人物形像_使用者.大人物圖片width Then
       大人物形像_使用者.Left = -大人物形像_使用者.大人物圖片width
       tr大人物形像_使用者.Enabled = False
       stup.Enabled = True
       stdown.Enabled = True
   Else
       大人物形像_使用者.Left = 大人物形像_使用者.Left - kp
   End If
End If
End Sub

Private Sub tr大人物形像_電腦_Timer()
Dim kr As Integer, kn As Integer

kn = 大人物形像_電腦.大人物圖片width
Dim bigwn, bigall As Integer
bigwn = (大人物形像_電腦.大人物圖片width / 2)
If 8760 - bigwn > Val(FormMainMode.ScaleWidth) - Val(大人物形像_電腦.大人物圖片width) Or Val(VBEPerson(2, 1, 2, 2, 5)) = 1 Then
    bigall = Val(FormMainMode.ScaleWidth) - Val(大人物形像_電腦.大人物圖片width)
Else
    bigall = 8760 - bigwn
End If
kr = (Val(FormMainMode.ScaleWidth) - bigall) / 30
If sq <= 400 Then
   If 大人物形像_電腦.Left <= bigall Then
       大人物形像_電腦.Left = bigall
       tr大人物形像_電腦.Enabled = False
   Else
       If 大人物形像_電腦.Left - bigall < kr Then
           大人物形像_電腦.Left = 大人物形像_電腦.Left - (大人物形像_電腦.Left - bigall)
       Else
           大人物形像_電腦.Left = 大人物形像_電腦.Left - kr
       End If
   End If
Else
   If 大人物形像_電腦.Left >= FormMainMode.ScaleWidth Then
       大人物形像_電腦.Left = FormMainMode.ScaleWidth
       tr大人物形像_電腦.Enabled = False
   Else
       大人物形像_電腦.Left = 大人物形像_電腦.Left + kr
   End If
End If
End Sub

Private Sub cardcomtr_Timer()
Dim i As Integer
If sq <= 400 Then
  For i = 3 To 1 Step -1
     If PEAScardcom(i).Visible = True Then
        If i < 3 Then
           If PEAScardcom(i + 1).Visible = True And PEAScardcom(i + 1).Top - PEAScardcom(i).Top >= 4000 Then
               Select Case i
                  Case 2
                     If PEAScardcom(i).Top <= 4000 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                     End If
                     If PEAScardcom(i).Top >= 4080 Then
                         PEAScardcom(i).Top = 4080
                     End If
                  Case 1
                     If PEAScardcom(i).Top <= 3700 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                     End If
                End Select
           ElseIf PEAScardcom(i + 1).Visible = False Or PEAScardcom(i + 1).Top >= 3000 Then
               Select Case i
                  Case 2
                      If PEAScardcom(i).Top <= 4000 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                     End If
                     If PEAScardcom(i).Top >= 4080 Then
                         PEAScardcom(i).Top = 4080
                     End If
                  Case 1
                      If PEAScardcom(i).Top <= 3700 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                      End If
                End Select
'               PEAScardcom(i).Top = PEAScardcom(i).Top + 200
           End If
        Else
           If PEAScardcom(i).Top <= 4400 Then
               PEAScardcom(i).Top = PEAScardcom(i).Top + 200
           End If
           If PEAScardcom(i).Top >= 4440 Then
                PEAScardcom(i).Top = 4440
           End If
        End If
    End If
  Next
  If PEAScardcom(1).Top >= 3720 Then
    PEAScardcom(1).Top = 3720
    cardcomtr.Enabled = False
    tr大人物形像_電腦.Enabled = True
  End If
End If
End Sub

Private Sub cardustr_Timer()
Dim i As Integer
If sq <= 400 Then
  For i = 3 To 1 Step -1
     If PEAScardus(i).Visible = True Then
        If i < 3 Then
           If PEAScardus(i + 1).Visible = True And PEAScardus(i + 1).Top - PEAScardus(i).Top >= 4000 Then
               Select Case i
                  Case 2
                     If PEAScardus(i).Top <= 4000 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                     End If
                     If PEAScardus(i).Top >= 4080 Then
                         PEAScardus(i).Top = 4080
                     End If
                  Case 1
                     If PEAScardus(i).Top <= 3700 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                     End If
                End Select
           ElseIf PEAScardus(i + 1).Visible = False Or PEAScardus(i + 1).Top >= 3000 Then
               Select Case i
                  Case 2
                      If PEAScardus(i).Top <= 4000 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                     End If
                     If PEAScardus(i).Top >= 4080 Then
                         PEAScardus(i).Top = 4080
                     End If
                  Case 1
                      If PEAScardus(i).Top <= 3700 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                      End If
                End Select
'               cardus(i).Top = cardus(i).Top + 200
           End If
        Else
           If PEAScardus(i).Top <= 4400 Then
               PEAScardus(i).Top = PEAScardus(i).Top + 200
           End If
           If PEAScardus(i).Top >= 4440 Then
                PEAScardus(i).Top = 4440
           End If
        End If
    End If
  Next
  If PEAScardus(1).Top >= 3720 Then
    PEAScardus(1).Top = 3720
    cardustr.Enabled = False
    tr大人物形像_使用者.Enabled = True
  End If
End If
End Sub
