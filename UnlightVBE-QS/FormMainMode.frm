VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
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
   Tag             =   "UnlightVBE-QS OB-Test"
   Begin VB.PictureBox PEAttackingForm 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9910
      Left            =   8160
      Picture         =   "FormMainMode.frx":0CCA
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   11340
      Begin UnlightVBE.uc角色卡片介面 cardcom 
         Height          =   3615
         Index           =   0
         Left            =   1080
         TabIndex        =   207
         Top             =   6240
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   2355
         _ExtentY        =   3625
      End
      Begin UnlightVBE.uc角色卡片介面 cardus 
         Height          =   3615
         Index           =   0
         Left            =   0
         TabIndex        =   206
         Top             =   6240
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   2355
         _ExtentY        =   3625
      End
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
         TabIndex        =   203
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
         TabIndex        =   202
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
         TabIndex        =   201
         Top             =   9360
         Width           =   1215
      End
      Begin VB.PictureBox uspiin 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5040
         Picture         =   "FormMainMode.frx":22BA9
         ScaleHeight     =   495
         ScaleWidth      =   2520
         TabIndex        =   183
         Top             =   9360
         Width           =   2520
         Begin VB.Label uspi5 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   189
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspidef 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   188
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspiatk 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   187
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   186
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspi4 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "2"
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
            Left            =   2040
            TabIndex        =   185
            Top             =   -30
            Width           =   495
         End
         Begin VB.Label uspi1 
            BackStyle       =   0  '透明
            Caption         =   "測試3"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   184
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.PictureBox PEAFtoolbox 
         Height          =   2175
         Left            =   5160
         ScaleHeight     =   2115
         ScaleWidth      =   6075
         TabIndex        =   174
         Top             =   6600
         Visible         =   0   'False
         Width           =   6135
         Begin VB.Timer 攻擊階段_階段初始 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   4680
            Top             =   240
         End
         Begin VB.Timer 智慧型AI_使用者出牌 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   240
            Top             =   840
         End
         Begin VB.Timer 移動階段_階段前啟動 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   1920
            Top             =   960
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
            TabIndex        =   182
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
            TabIndex        =   181
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
            TabIndex        =   180
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
            TabIndex        =   179
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
            TabIndex        =   178
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
            TabIndex        =   177
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
            TabIndex        =   176
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
            TabIndex        =   175
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
      Begin VB.Frame atkinghelpc 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   2775
         Left            =   7680
         TabIndex        =   36
         Top             =   3360
         Width           =   2205
         Begin VB.Label atkinghelpt4 
            BackColor       =   &H00000000&
            Caption         =   "む這裡是技能效果區めめめめめめめめめめめめめめめめめめ"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   2295
            Left            =   120
            TabIndex        =   42
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label atkinghelpt3 
            BackColor       =   &H00000000&
            BackStyle       =   0  '透明
            Caption         =   "む卡片めめめめめめめめめ"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   720
            TabIndex        =   37
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label atkinghelpt2 
            BackColor       =   &H00000000&
            Caption         =   "む距離め"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label atkinghelpi1 
            BackColor       =   &H00000000&
            Caption         =   "「條件」"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   0
            Width           =   975
         End
         Begin VB.Label atkinghelpi5 
            BackColor       =   &H00000000&
            Caption         =   "「效果」"
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label atkinghelpi3 
            BackColor       =   &H00000000&
            Caption         =   "距離："
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   735
         End
         Begin VB.Label atkinghelpi4 
            BackColor       =   &H00000000&
            Caption         =   "卡片："
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   735
         End
         Begin VB.Label atkinghelpt1 
            BackColor       =   &H00000000&
            Caption         =   "む階段め"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   39
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label atkinghelpi2 
            BackColor       =   &H00000000&
            Caption         =   "階段："
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
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
      Begin VB.Timer tr牌組_翻牌 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1080
         Top             =   3600
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
      Begin VB.PictureBox uspiin 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   1
         Left            =   0
         Picture         =   "FormMainMode.frx":22DF1
         ScaleHeight     =   495
         ScaleWidth      =   2520
         TabIndex        =   29
         Top             =   9360
         Width           =   2520
         Begin VB.Label uspi1 
            BackStyle       =   0  '透明
            Caption         =   "測試1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label uspi4 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "2"
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
            Left            =   2040
            TabIndex        =   34
            Top             =   -30
            Width           =   495
         End
         Begin VB.Label uspi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspiatk 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   32
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspidef 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   31
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspi5 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   30
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.PictureBox uspiin 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2520
         Picture         =   "FormMainMode.frx":23039
         ScaleHeight     =   495
         ScaleWidth      =   2520
         TabIndex        =   22
         Top             =   9360
         Width           =   2520
         Begin VB.Label uspi5 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   28
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspidef 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   27
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspiatk 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   26
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   255
         End
         Begin VB.Label uspi4 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "2"
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
            Left            =   2040
            TabIndex        =   24
            Top             =   -30
            Width           =   495
         End
         Begin VB.Label uspi1 
            BackStyle       =   0  '透明
            Caption         =   "測試2"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.PictureBox compiin 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5040
         Picture         =   "FormMainMode.frx":23281
         ScaleHeight     =   495
         ScaleWidth      =   2520
         TabIndex        =   15
         Top             =   0
         Width           =   2520
         Begin VB.Label compi1 
            BackStyle       =   0  '透明
            Caption         =   "測試3"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label compi4 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "2"
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
            Left            =   2040
            TabIndex        =   20
            Top             =   -30
            Width           =   495
         End
         Begin VB.Label compi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compiatk 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   18
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compidef 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   17
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compi5 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   16
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.PictureBox compiin 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2520
         Picture         =   "FormMainMode.frx":234C9
         ScaleHeight     =   495
         ScaleWidth      =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   2520
         Begin VB.Label compi1 
            BackStyle       =   0  '透明
            Caption         =   "艾伯李斯特"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label compi4 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "2"
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
            Left            =   2040
            TabIndex        =   13
            Top             =   -30
            Width           =   495
         End
         Begin VB.Label compi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compiatk 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compidef 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   10
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compi5 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.PictureBox compiin 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         Picture         =   "FormMainMode.frx":23711
         ScaleHeight     =   495
         ScaleWidth      =   2520
         TabIndex        =   1
         Top             =   0
         Width           =   2520
         Begin VB.Label compi1 
            BackStyle       =   0  '透明
            Caption         =   "艾伯李斯特"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label compi4 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "2"
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
            Left            =   2040
            TabIndex        =   6
            Top             =   -30
            Width           =   495
         End
         Begin VB.Label compi2 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compiatk 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compidef 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   3
            Top             =   360
            Width           =   255
         End
         Begin VB.Label compi5 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   2
            Top             =   360
            Width           =   255
         End
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
         TabIndex        =   47
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
         TabIndex        =   48
         Top             =   5600
         Width           =   375
      End
      Begin VB.Image PEAFbloodbackimage2 
         Height          =   690
         Left            =   10080
         Picture         =   "FormMainMode.frx":23959
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
         TabIndex        =   49
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
         TabIndex        =   50
         Top             =   5600
         Width           =   450
      End
      Begin VB.Image PEAFbloodbackimage1 
         Height          =   690
         Left            =   0
         Picture         =   "FormMainMode.frx":240A0
         Top             =   5440
         Width           =   1290
      End
      Begin UnlightVBE.uc擲骰介面 PEAFDiceInterface 
         Height          =   9910
         Left            =   0
         TabIndex        =   205
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
         TabIndex        =   200
         Top             =   6600
         Width           =   135
      End
      Begin VB.Image bnok 
         Height          =   1050
         Left            =   8040
         Picture         =   "FormMainMode.frx":24830
         Top             =   8280
         Visible         =   0   'False
         Width           =   1500
      End
      Begin UnlightVBE.ucCard card 
         Height          =   1335
         Index           =   0
         Left            =   5280
         TabIndex        =   199
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
         TabIndex        =   198
         Top             =   720
         Width           =   135
      End
      Begin VB.Label comaiatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   1
         Left            =   6960
         TabIndex        =   197
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label comaiatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   196
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label comaiatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   195
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label comaiatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   194
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label personatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   4
         Left            =   9140
         TabIndex        =   193
         Top             =   6240
         Width           =   2205
      End
      Begin VB.Label personatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   3
         Left            =   6930
         TabIndex        =   192
         Top             =   6240
         Width           =   2205
      End
      Begin VB.Label personatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   2
         Left            =   4730
         TabIndex        =   191
         Top             =   6240
         Width           =   2205
      End
      Begin VB.Label personatk 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         Caption         =   "人物技能"
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro M"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   190
         Top             =   6240
         Width           =   2205
      End
      Begin UnlightVBE.顯示列 顯示列1 
         Height          =   1215
         Left            =   0
         TabIndex        =   46
         Top             =   3520
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2143
      End
      Begin VB.Image PEAFtest 
         Height          =   1080
         Index           =   0
         Left            =   3120
         Picture         =   "FormMainMode.frx":2873F
         Tag             =   "movejpg"
         Top             =   2160
         Visible         =   0   'False
         Width           =   5490
      End
      Begin VB.Image atkdef2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":2A0B2
         Top             =   1860
         Width           =   2280
      End
      Begin VB.Image atkdef1 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":2A7FC
         Top             =   1590
         Width           =   2280
      End
      Begin VB.Image draw2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":2AF52
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Image move2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":2B5B9
         Top             =   1320
         Width           =   2280
      End
      Begin VB.Image PEAFtest 
         Height          =   2865
         Index           =   1
         Left            =   2040
         Picture         =   "FormMainMode.frx":2BCDB
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   2160
         Width           =   135
      End
      Begin VB.Image Image2 
         Height          =   120
         Left            =   5280
         Picture         =   "FormMainMode.frx":2D809
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
         Picture         =   "FormMainMode.frx":2D89C
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
      Begin VB.Line 小人物角色基準線 
         BorderColor     =   &H000000FF&
         X1              =   1080
         X2              =   10320
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line 小人物距離基準線 
         BorderColor     =   &H000000FF&
         Index           =   2
         X1              =   2640
         X2              =   2640
         Y1              =   5880
         Y2              =   6120
      End
      Begin VB.Line 小人物距離基準線 
         BorderColor     =   &H000000FF&
         Index           =   1
         X1              =   1040
         X2              =   1040
         Y1              =   5880
         Y2              =   6120
      End
      Begin VB.Line 小人物距離基準線 
         BorderColor     =   &H000000FF&
         Index           =   3
         X1              =   4320
         X2              =   4320
         Y1              =   5880
         Y2              =   6120
      End
      Begin VB.Line 小人物距離基準線 
         BorderColor     =   &H000000FF&
         Index           =   4
         X1              =   7080
         X2              =   7080
         Y1              =   5880
         Y2              =   6120
      End
      Begin VB.Line 小人物距離基準線 
         BorderColor     =   &H000000FF&
         Index           =   5
         X1              =   8680
         X2              =   8680
         Y1              =   5880
         Y2              =   6120
      End
      Begin VB.Line 小人物距離基準線 
         BorderColor     =   &H000000FF&
         Index           =   6
         X1              =   10320
         X2              =   10320
         Y1              =   5880
         Y2              =   6120
      End
      Begin VB.Image PEAFtest 
         Height          =   2880
         Index           =   2
         Left            =   9960
         Picture         =   "FormMainMode.frx":2D908
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
         Picture         =   "FormMainMode.frx":2F332
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Image move1 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":2F4B1
         Top             =   1340
         Width           =   2040
      End
      Begin VB.Image move3 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":2F64D
         Top             =   1610
         Width           =   2040
      End
      Begin VB.Image move4 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":2F8D6
         Top             =   1880
         Width           =   2040
      End
      Begin UnlightVBE.小人物形象 personusminijpg 
         Height          =   4935
         Left            =   0
         TabIndex        =   54
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8705
      End
      Begin UnlightVBE.小人物形象 personcomminijpg 
         Height          =   4935
         Left            =   5520
         TabIndex        =   55
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8705
      End
      Begin UnlightVBE.小人物形象 movejpg 
         Height          =   2535
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4471
      End
      Begin VB.Image cardpagejpg 
         Height          =   915
         Left            =   0
         Picture         =   "FormMainMode.frx":2FB53
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
         TabIndex        =   51
         Top             =   1100
         Width           =   855
      End
      Begin UnlightVBE.uc戰鬥系統牌型介面 PEAFInterface 
         Height          =   9910
         Left            =   0
         TabIndex        =   204
         Top             =   0
         Width           =   11340
         _ExtentX        =   2143
         _ExtentY        =   2778
      End
   End
   Begin VB.PictureBox PEGameFreeModeSettingForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   1080
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   57
      Top             =   480
      Visible         =   0   'False
      Width           =   11340
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1785
         ScaleWidth      =   11385
         TabIndex        =   82
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            TabIndex        =   89
            Top             =   1440
            Width           =   855
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
            Index           =   1
            Left            =   4320
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton opnpersonvs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3v3"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   84
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton personreadifus 
            Caption         =   "讀入..."
            Height          =   495
            Left            =   2040
            TabIndex        =   83
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
            TabIndex        =   85
            Top             =   720
            Value           =   -1  'True
            Width           =   855
         End
         Begin UnlightVBE.大人物形像 personfus 
            Height          =   1215
            Left            =   0
            TabIndex        =   113
            Top             =   0
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2143
         End
         Begin VB.Image bnstart 
            Height          =   510
            Left            =   9600
            Picture         =   "FormMainMode.frx":303B6
            Top             =   600
            Width           =   1440
         End
         Begin VB.Image bnabout 
            Height          =   390
            Left            =   8280
            Picture         =   "FormMainMode.frx":30F0C
            Top             =   720
            Width           =   1320
         End
         Begin VB.Image bnconfig 
            Height          =   390
            Left            =   7080
            Picture         =   "FormMainMode.frx":31577
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "VS"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   20.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   5400
            TabIndex        =   112
            Top             =   600
            Width           =   735
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   109
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
            TabIndex        =   108
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
            TabIndex        =   107
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
            TabIndex        =   106
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label personsettingus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "人物設定"
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
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label personsettingus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "人物設定"
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
            Left            =   2760
            TabIndex        =   104
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label personsettingus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "人物設定"
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
            Left            =   5400
            TabIndex        =   103
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label personsettingcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "人物設定"
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
            Left            =   3360
            TabIndex        =   102
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label personsettingcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "人物設定"
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
            Left            =   6000
            TabIndex        =   101
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label personsettingcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "人物設定"
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
            Left            =   8640
            TabIndex        =   100
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '透明
            Caption         =   "1P"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   8040
            TabIndex        =   99
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '透明
            Caption         =   "COM"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   2400
            TabIndex        =   98
            Top             =   1400
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
         Picture         =   "FormMainMode.frx":31B87
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   78
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   79
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
         Picture         =   "FormMainMode.frx":35A2A
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   74
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
         Picture         =   "FormMainMode.frx":398CD
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   70
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
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
         Picture         =   "FormMainMode.frx":3D770
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   66
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
         Picture         =   "FormMainMode.frx":41613
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   62
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   63
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
         Picture         =   "FormMainMode.frx":454B6
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   58
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   59
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
         TabIndex        =   115
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
         TabIndex        =   114
         Top             =   195
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   465
         Left            =   0
         Picture         =   "FormMainMode.frx":49359
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
      Left            =   -1200
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   126
      Top             =   2520
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
         TabIndex        =   127
         Top             =   9120
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.PictureBox PEAttackingStartForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   3120
      Picture         =   "FormMainMode.frx":4B748
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   128
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
         TabIndex        =   129
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
            TabIndex        =   132
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
            TabIndex        =   131
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
            TabIndex        =   130
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
         TabIndex        =   159
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
            TabIndex        =   162
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
            TabIndex        =   161
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
            TabIndex        =   160
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
         TabIndex        =   133
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
            TabIndex        =   136
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
            TabIndex        =   135
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
            TabIndex        =   134
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
         TabIndex        =   151
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
            TabIndex        =   154
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
            TabIndex        =   153
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
            TabIndex        =   152
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
         TabIndex        =   163
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
            TabIndex        =   166
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
            TabIndex        =   165
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
            TabIndex        =   164
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
         TabIndex        =   155
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
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
         Picture         =   "FormMainMode.frx":7A5FC
         ScaleHeight     =   1455
         ScaleWidth      =   11415
         TabIndex        =   138
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
            TabIndex        =   150
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
            TabIndex        =   149
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
            TabIndex        =   148
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
            TabIndex        =   147
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
            TabIndex        =   146
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
            TabIndex        =   145
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
            TabIndex        =   144
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   139
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
         Picture         =   "FormMainMode.frx":82F78
         ScaleHeight     =   1905
         ScaleWidth      =   11415
         TabIndex        =   137
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
         TabIndex        =   173
         Top             =   -120
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3413
      End
      Begin UnlightVBE.大人物形像 大人物形像_電腦 
         Height          =   10005
         Left            =   20040
         TabIndex        =   167
         Top             =   -480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   17648
      End
      Begin UnlightVBE.大人物形像 大人物形像_使用者 
         Height          =   10005
         Left            =   -9960
         TabIndex        =   168
         Top             =   -480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   17648
      End
      Begin UnlightVBE.大人物形像 upjpg_2 
         Height          =   1935
         Left            =   0
         TabIndex        =   169
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
      Picture         =   "FormMainMode.frx":8E4A4
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   170
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
         TabIndex        =   172
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
         TabIndex        =   171
         Top             =   8760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image bn 
         Height          =   990
         Left            =   9480
         Picture         =   "FormMainMode.frx":B140F
         Top             =   8520
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Image bnreturn 
         Height          =   990
         Left            =   7680
         Picture         =   "FormMainMode.frx":B2304
         Top             =   8520
         Visible         =   0   'False
         Width           =   1470
      End
   End
   Begin VB.PictureBox PEMusicForm 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   2040
      ScaleHeight     =   7935
      ScaleWidth      =   9615
      TabIndex        =   116
      Top             =   840
      Visible         =   0   'False
      Width           =   9615
      Begin VB.Timer PEMtr1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   240
         Top             =   1680
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmp 
         Height          =   960
         Left            =   840
         TabIndex        =   125
         Top             =   480
         Width           =   1080
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1905
         _cy             =   1693
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse1 
         Height          =   600
         Left            =   3840
         TabIndex        =   124
         Top             =   240
         Width           =   2160
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   3810
         _cy             =   1058
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse2 
         Height          =   720
         Left            =   3840
         TabIndex        =   123
         Top             =   1200
         Width           =   2280
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4022
         _cy             =   1270
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse3 
         Height          =   600
         Left            =   3840
         TabIndex        =   122
         Top             =   2160
         Width           =   2400
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4233
         _cy             =   1058
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse4 
         Height          =   600
         Left            =   3840
         TabIndex        =   121
         Top             =   2880
         Width           =   2400
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4233
         _cy             =   1058
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse5 
         Height          =   600
         Left            =   3840
         TabIndex        =   120
         Top             =   3840
         Width           =   2280
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4022
         _cy             =   1058
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse6 
         Height          =   600
         Left            =   6840
         TabIndex        =   119
         Top             =   360
         Width           =   2400
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4233
         _cy             =   1058
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse7 
         Height          =   600
         Left            =   6840
         TabIndex        =   118
         Top             =   1200
         Width           =   2400
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4233
         _cy             =   1058
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmpse8 
         Height          =   720
         Left            =   6840
         TabIndex        =   117
         Top             =   2160
         Width           =   2400
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "invisible"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4233
         _cy             =   1270
      End
   End
End
Attribute VB_Name = "FormMainMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub atkinghelpc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpi1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpi2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpi3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpi4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpi5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpt1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpt2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpt3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub


Private Sub atkinghelpt4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
atkinghelpc.Visible = False
End Sub

Private Sub atkingtrcom_Timer()
If 目前數(29) = 1 Then
   目前數(31) = 0
'   Formatkingcom.Left = FormMainMode.Left + 5295
   Formatkingcom.Left = FormMainMode.Left + (FormMainMode.Width - Formatkingcom.Width)
   Formatkingcom.Top = FormMainMode.Top + 380
   atkingtrcom.Enabled = False
   Formatkingcom.t1.Enabled = True
'   技能動畫顯示_電腦tr.Enabled = True
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
'   技能動畫顯示_使用者tr.Enabled = True
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

Sub bnok_Click()
If turnpageonin = 1 Then
    turnpageonin = 0
    For i = 1 To 公用牌實體卡片分隔紀錄數(1)
        FormMainMode.card(i).card_MouseExit
    Next
    bnok.Picture = LoadPicture(app_path & "gif\system\ok_3.jpg")
    戰鬥系統類.時間軸_停止
    Select Case turnatk
        Case 1
            目前數(22) = 7
            等待時間.Enabled = True
        Case 2
            目前數(22) = 8
            等待時間.Enabled = True
        Case 3
            cnmove2_Click
    End Select
End If
End Sub

Private Sub bnok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If turnpageonin = 1 Then
    bnok.Picture = LoadPicture(app_path & "gif\system\ok_2.jpg")
End If
For i = 1 To 公用牌實體卡片分隔紀錄數(1)
   card(i).CardEventType = False
Next
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
Dim uspce As String, uspme As String
uspce = pagecardnum(Index, 1)
uspme = pagecardnum(Index, 2)
pagecardnum(Index, 1) = pagecardnum(Index, 3)
pagecardnum(Index, 2) = pagecardnum(Index, 4)
pagecardnum(Index, 3) = uspce
pagecardnum(Index, 4) = uspme
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
pageonin(Index) = Val(card(Index).CardRotationType)
End Sub

Sub card_CardButtonClickout(Index As Integer)
Dim uspce As String, uspme As String
uspce = pagecardnum(Index, 1)
uspme = pagecardnum(Index, 2)
pagecardnum(Index, 1) = pagecardnum(Index, 3)
pagecardnum(Index, 2) = pagecardnum(Index, 4)
pagecardnum(Index, 3) = uspce
pagecardnum(Index, 4) = uspme
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
pageonin(Index) = Val(card(Index).CardRotationType)
'===================================================================
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 And movecp = 1 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 1 And movecp = 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 And movecp > 1 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 1 And movecp > 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + defus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 2 Then
         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2))
         攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) + Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) + Val(pagecardnum(Index, 2))
   End If
'======================================
   If pagecardnum(Index, 3) = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) - Val(pagecardnum(Index, 4))
      If turnatk = 1 And movecp = 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 4))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(pagecardnum(Index, 4))
      End If
      If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
          攻擊防禦骰子總數(3) = 0
      End If
   End If
   If pagecardnum(Index, 3) = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) - Val(pagecardnum(Index, 4))
      If turnatk = 1 And movecp > 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 4))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(pagecardnum(Index, 4))
      End If
      If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
          攻擊防禦骰子總數(3) = 0
      End If
   End If
   If pagecardnum(Index, 3) = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) - Val(pagecardnum(Index, 4))
      If turnatk = 2 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 4))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) - Val(pagecardnum(Index, 4))
   End If
   If pagecardnum(Index, 3) = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) - Val(pagecardnum(Index, 4))
   End If
    '============以下是技能檢查及啟動

'==============================================
Select Case turnatk
    Case 1
        '===========================執行階段插入點(ATK-42/DEF-43)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
        '============================
    Case 2
        '===========================執行階段插入點(ATK-42/DEF-43)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
        '============================
    Case 3
        '===========================執行階段插入點(44)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
        '============================
End Select
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1.Enabled = True
End Sub


Sub card_CardClick(Index As Integer)
'======================以下為專屬事件卡檢查
If pagecardnum(Index, 1) = a7a And turnatk <> 1 And turnatk <> 2 Then
   '=========違反詛咒術事件卡只在攻防階段使用原則
   Exit Sub
End If
'====================================
If pagecardnum(Index, 6) = 1 And (turnpageonin = 1 Or turnpageoninatking = 1) And pagecardnum(Index, 5) = 1 Then
   pagecardnum(Index, 6) = 2
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 And movecp = 1 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 1 And movecp = 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 And movecp > 1 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 1 And movecp > 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And 攻擊防禦骰子總數(3) = 0 Then
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + defus(角色人物對戰人數(1, 2))
      End If
      If turnatk = 2 Then
         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2))
         攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) + Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) + Val(pagecardnum(Index, 2))
   End If
   '=================
   turnpageonin = 0
   card(Index).LocationType = 2
   '===================
   目前數(5) = pagecardnum(Index, 7)
   pageqlead(1) = Val(pageqlead(1)) + 1
   pageusglead = Val(pageusglead) - 1
   pagecardnum(Index, 7) = Val(pageusleadmax(1)) + 1
   pageusleadmax(1) = Val(pageusleadmax(1)) + 1
   pageusqlead = Val(pageusqlead) + 1
   目前數(13) = 0
   '===================以下是出牌對齊
   目前數(3) = 0
   戰鬥系統類.出牌順序計算_使用者_出牌
   使用者出牌_出牌對齊_靠左.Enabled = True
    '============以下是技能檢查及啟動

   '=============以下是牌移動(出牌)(使用者)
    戰鬥系統類.座標計算_使用者出牌
    牌移動暫時變數(3) = Index
    pagecardnum(Index, 9) = card(Index).Left  '指定目前Left(座標)
    pagecardnum(Index, 10) = card(Index).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    目前數(15) = 0
    牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
   '================以下是手牌對齊
   目前數(4) = 0
   目前數(21) = 1
   戰鬥系統類.出牌順序計算_使用者_手牌
   使用者出牌_手牌對齊.Enabled = True
   '=================
   對齊完成檢查.Enabled = True
   '===================以下是事件卡檢查及啟動
   If pagecardnum(Index, 1) = a6a Then
       事件卡記錄暫時數(1, 3) = 1
       事件卡.機會_使用者 Index, pagecardnum(Index, 2)
   End If
   If turnatk = 1 Or turnatk = 2 Then
        If pagecardnum(Index, 1) = a7a Then
            事件卡記錄暫時數(1, 3) = 1
            事件卡.詛咒術_使用者 Index, pagecardnum(Index, 2)
        End If
   End If
   If pagecardnum(Index, 1) = a8a Then
       事件卡記錄暫時數(1, 3) = 1
       事件卡.HP回復_使用者 Index, pagecardnum(Index, 2)
   End If
   If pagecardnum(Index, 1) = a9a Then
       事件卡記錄暫時數(1, 3) = 1
       事件卡.聖水_使用者 Index, pagecardnum(Index, 2)
   End If
   '===================
   GoTo vsssystemplay
End If
'=================================================================
If pagecardnum(Index, 6) = 2 And (turnpageonin = 1 Or turnpageoninatking = 1) And pagecardnum(Index, 5) = 1 Then
   pagecardnum(Index, 6) = 1
   '===================================
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) - Val(pagecardnum(Index, 2))
      If turnatk = 1 And movecp = 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(pagecardnum(Index, 2))
      End If
      If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
          攻擊防禦骰子總數(3) = 0
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) - Val(pagecardnum(Index, 2))
      If turnatk = 1 And movecp > 1 Then
          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(pagecardnum(Index, 2))
      End If
      If 攻擊防禦骰子總數(3) = atkus(角色人物對戰人數(1, 2)) Then
          攻擊防禦骰子總數(3) = 0
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 Then
         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 2))
         攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) - Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) - Val(pagecardnum(Index, 2))
   End If
   '=============
   turnpageonin = 0
   card(Index).LocationType = 1
'   Erase atkingckdice
   '================
   目前數(5) = pagecardnum(Index, 7)
   pagecardnum(Index, 7) = Val(pageusleadmax(0)) + 1
   pageusleadmax(0) = Val(pageusleadmax(0)) + 1
   pageqlead(1) = Val(pageqlead(1)) - 1
   pageusglead = Val(pageusglead) + 1
   pageusqlead = Val(pageusqlead) - 1
'   目前數(1) = pageusglead
   '============以下是技能檢查及啟動

   '=============以下是牌移動(回牌)(使用者)
    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(3) = Index
    pagecardnum(Index, 9) = card(Index).Left  '指定目前Left(座標)
    pagecardnum(Index, 10) = card(Index).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    目前數(15) = 0
    牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
   '================以下是出牌對齊
   目前數(3) = 0
   戰鬥系統類.出牌順序計算_使用者_出牌
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
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
        '============================
    Case 2
        '===========================執行階段插入點(ATK-42/DEF-43)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
        '============================
    Case 3
        '===========================執行階段插入點(44)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
        '============================
End Select
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1.Enabled = True
End Sub


Private Sub card_CardMouseMove(Index As Integer)
If pagecardnum(Index, 6) = 1 And pagecardnum(Index, 5) = 1 And turnpageonin = 1 Then
    card(Index).CardEventType = True
ElseIf pagecardnum(Index, 6) = 2 And pagecardnum(Index, 5) = 1 And turnpageonin = 1 Then
    card(Index).CardEventType = True
Else
    card(Index).CardEventType = False
End If
End Sub

Sub cnmove_Click()
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
livecom(角色人物對戰人數(2, 2)) = Val(compi4(角色人物對戰人數(2, 2)).Caption)
liveus(角色人物對戰人數(1, 2)) = Val(uspi4(角色人物對戰人數(1, 2)).Caption)
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
    FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\move1-2.gif"
End If
顯示列1.顯示列圖片 = app_path & "gif\system\linemove.png"
cnmove.Visible = False
'cnmove2.Visible = True
'turnpageonin = 1
戰鬥系統類.cleanatkingpagetot
'==============
'For i = 1 To UBound(atkingck)
'     atkingck(i, 1) = 1
'     atkingck(i, 2) = 0
'Next
'For i = 1 To UBound(atkingckai)
'     atkingckai(i, 1) = 1
'     atkingckai(i, 2) = 0
'Next
'======================電腦方事件卡先出制度
If 電腦方事件卡是否出完選擇數 = False Then
    GoTo 電腦方事件卡先出制度_執行階段2
End If
'================================
電腦方事件卡先出制度_執行階段結束:
'----------以下為電腦判斷出牌程式碼（移動階段1）
'====================試驗智慧型AI出牌系統
'If 智慧型AI系統_目前可執行之人物判斷(namecom(角色人物對戰人數(2, 2))) = True Then
    智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 2, 3, namecom(角色人物對戰人數(2, 2)), movecp, 0
    GoTo 智慧型AI出牌_執行階段結束
'End If
'=========以下為技能檢查及啟動
'   If turnatk = 3 Then
'      AI技能.雪莉_巨大黑犬 '(階段1)
'   End If
''============以下是異常狀態檢查及啟動
'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 17) = True Then
'      異常狀態檢查數(17, 1) = 1
'      異常狀態.麻痺_電腦  '(階段1)
'      電腦方移動階段選擇數 = 2
'      GoTo 麻痺_電腦_執行階段2
'End If
''======================

Dim movecomatk1, movecomatk2 As Integer
戰鬥系統類.moveatkin

For i = 1 To 公用牌實體卡片分隔紀錄數(1)
   If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 11)) <> 1 Then
       If pagecardnum(i, 1) = a1a Then movecomatk1 = Val(movecomatk1) + Val(pagecardnum(i, 2))
       If pagecardnum(i, 1) = a5a Then movecomatk2 = Val(movecomatk2) + Val(pagecardnum(i, 2))
       If pagecardnum(i, 3) = a1a Then movecomatk1 = Val(movecomatk1) + Val(pagecardnum(i, 4))
       If pagecardnum(i, 3) = a5a Then movecomatk2 = Val(movecomatk2) + Val(pagecardnum(i, 4))
   End If
Next
麻痺_電腦_執行階段2: '異常狀態-麻痺-電腦-程式跳入點(執行階段2)
''=====================以下是異常狀態檢查及啟動
'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 21) = True And livecom(角色人物對戰人數(2, 2)) <= 1 Then
'      異常狀態檢查數(21, 1) = 2
'      異常狀態.中毒_電腦  '(階段2)
'End If
'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 3) = True Then
'      異常狀態檢查數(3, 1) = 1
'      異常狀態.MOV加_電腦  '(階段1)
'End If
'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 6) = True Then
'      異常狀態檢查數(6, 1) = 1
'      異常狀態.MOV減_電腦  '(階段1)
'End If
''=========
'AI人物.史塔夏 2
'AI人物.全人物通用 1   '===異常狀態-MOV減-有效移動值判斷處理
''==============
'AI人物.阿奇波爾多 1
'AI人物.艾伯李斯特 2
'AI人物.CC 2
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
    FormMainMode.wmpse6.Controls.play
    一般系統類.檢查音樂播放 6
End If
'======================電腦方事件卡先出制度_結束後階段2
If 電腦方事件卡是否出完選擇數 = True Then
    電腦出牌.Enabled = True
End If
'===========================
End Sub

Private Sub cnmove2_Click()
turnpageonin = 0
OK按鈕牌完成移動檢查.Enabled = True
cnmove2.Visible = False
End Sub

Private Sub comaiatk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 3
  cardcom(i).Visible = False
Next
'Select Case Index
'   Case 1
'       Select Case comaiatk(Index).Caption
'            Case "自殺傾向"
'              akhpnm = 1
'            Case "猛擊"
'              akhpnm = 4
'            Case "輪旋曲-琉璃色的微風"
'              akhpnm = 10
'            Case "Ex輪旋曲-琉璃色的微風"
'              akhpnm = 16
'            Case "重壓"
'              akhpnm = 23
'            Case "冰結之翼"
'              akhpnm = 24
'        End Select
'   Case 2
'       Select Case comaiatk(Index).Caption
'            Case "異質者"
'              akhpnm = 7
'            Case "超再生"
'              akhpnm = 22
'            Case "煉獄之翼"
'              akhpnm = 25
'        End Select
'   Case 3
'       Select Case comaiatk(Index).Caption
'            Case "巨大黑犬"
'              akhpnm = 3
'            Case "混沌之翼"
'              akhpnm = 26
'        End Select
'   Case 4
'       Select Case comaiatk(Index).Caption
'            Case "飛刃雨"
'              akhpnm = 2
'            Case "終曲-無盡輪迴的終結"
'              akhpnm = 12
'       End Select
'End Select
'技能說明載入 akhpnm
    戰鬥系統類.技能說明載入_電腦 Index
    
    atkinghelpc.Left = atkinghelpxy(2, Index, 1)
    atkinghelpc.Top = atkinghelpxy(2, Index, 2)
    atkinghelpc.ZOrder
    atkinghelpc.Visible = True
End Sub




Private Sub compi1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cardcom(Index).Left = compiin(Index).Left
cardcom(Index).Top = 480
cardcom(Index).ZOrder
If 人物卡面背面編號紀錄數(1) = 2 And 人物卡面背面編號紀錄數(2) = Index Then
'    PEAFcardback(1).Visible = True
    cardcom(Index).Visible = True
'    PEAFcardback(1).ZOrder
Else
    cardcom(Index).Visible = True
'    PEAFcardback(1).Visible = False
End If
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
atkinghelpc.Visible = False
End Sub


Private Sub compi4_Change(Index As Integer)
  If Val(compi4(Index).Caption) = Val(livecommax(Index)) Then
   compi4(Index).ForeColor = RGB(255, 255, 255)
'   cardcompi1(Index).ForeColor = RGB(255, 255, 255)
'   cardbackcom(Index).Visible = False
 End If
 If Val(compi4(Index).Caption) < Val(livecommax(Index)) Then
   compi4(Index).ForeColor = RGB(255, 255, 128)
'   cardcompi1(Index).ForeColor = RGB(255, 255, 128)
'   cardbackcom(Index).Visible = False
 End If
 If Val(compi4(Index).Caption) <= Val(livecom41(Index)) Then
   compi4(Index).ForeColor = RGB(255, 0, 0)
'   cardcompi1(Index).ForeColor = RGB(255, 0, 0)
'   cardbackcom(Index).Visible = False
 End If
 If Val(compi4(Index).Caption) = 0 And compi1(Index).Caption = "" Then
   compi4(Index).ForeColor = RGB(255, 255, 255)
'   cardcompi1(Index).ForeColor = RGB(255, 255, 255)
'   cardbackcom(Index).Visible = False
 End If
 If Val(compi4(Index).Caption) <= 0 Then
    If compi1(Index).Caption <> "" Then
        compi4(Index).Caption = 0
'        cardcompi1(Index).Caption = 0
'        cardbackcom(Index).Visible = True
    Else
'        cardbackcom(Index).Visible = False
    End If
 End If
End Sub

Private Sub compi4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cardcom(Index).Left = compiin(Index).Left
cardcom(Index).Top = 480
cardcom(Index).ZOrder
If 人物卡面背面編號紀錄數(1) = 2 And 人物卡面背面編號紀錄數(2) = Index Then
'    PEAFcardback(1).Visible = True
    cardcom(Index).Visible = True
'    PEAFcardback(1).ZOrder
Else
    cardcom(Index).Visible = True
'    PEAFcardback(1).Visible = False
End If
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
atkinghelpc.Visible = False
End Sub


Private Sub compiin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cardcom(Index).Left = compiin(Index).Left
cardcom(Index).Top = 480
cardcom(Index).ZOrder
If 人物卡面背面編號紀錄數(1) = 2 And 人物卡面背面編號紀錄數(2) = Index Then
'    PEAFcardback(1).Visible = True
    cardcom(Index).Visible = True
'    PEAFcardback(1).ZOrder
Else
    cardcom(Index).Visible = True
'    PEAFcardback(1).Visible = False
End If
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
atkinghelpc.Visible = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
  YesNo = MsgBox("確定離開遊戲?", 36, "UnlightVBE-系統提示")
  If YesNo = 6 Then
    End
  Else
    Cancel = 1
  End If
End If
End Sub


Private Sub cn1_Click()
'draw1.Visible = False
'draw2.Visible = True
turnatk = 4
戰鬥系統類.音量靜音調節設定
'====================
'目前數(1) = 1
目前數(2) = 1

'If 牌總階段數(3) > 9 Then 牌總階段數(3) = 9 '現行限制(總擁有牌最大數)
'Erase atkingck
'===========================執行階段插入點(0)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 0, 1
'============================
cn1.Visible = False
目前數(15) = 1
發牌檢查.Enabled = True
'FormMainMode.wmpse6.Controls.play
End Sub

Private Sub cn2_Click()
If moveturn = 1 Then
  If 系統顯示界面紀錄數 = 1 Then
        move1.Visible = True
        move2.Visible = False
        atkdef1.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\atk2.gif"
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
        FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\atk2.gif"
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
'cn22.Visible = True
bnok.Picture = LoadPicture(app_path & "gif\system\ok_1.jpg")
bnok.Visible = True
'=============
livecom(角色人物對戰人數(2, 2)) = Val(compi4(角色人物對戰人數(2, 2)).Caption)
liveus(角色人物對戰人數(1, 2)) = Val(uspi4(角色人物對戰人數(1, 2)).Caption)
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
'Erase 異常狀態_混沌紀錄數
'Erase 異常狀態_AI_混沌紀錄數
''=====
'If atkingck(49, 2) = 1 And atking_尤莉卡_超載目前階段紀錄數(3) = 2 Then
'   atkingck(49, 1) = 5
'   技能.尤莉卡_超載 '(階段5)
'End If
'If atkingckai(139, 2) = 1 And atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2 Then
'   atkingckai(139, 1) = 5
'   AI技能.尤莉卡_超載 '(階段5)
'End If
'=====
If turnatk = 1 Then
 戰鬥系統類.chkdefcom
End If
'==============以下是技能檢查及啟動
'If uspi1(角色人物對戰人數(1, 2)).Caption = "音音夢" Then
'    If atking_音音夢_成長模式狀態數(2) = 1 And turnatk = 1 Then
'       atking_音音夢_成長模式狀態數(1) = 2
'       戰鬥系統類.特殊_音音夢_成長狀態_使用者 '(階段2)
'    End If
'End If
'If turnatk = 1 Then
'    atkingckai(44, 1) = 1
'    AI技能.庫勒尼西_沙漠中的海市蜃樓 '(階段1)
'End If
'==============
小人物頭像移動方向數(1) = 1
小人物頭像移動方向數(2) = 2
小人物頭像移動_使用者.Enabled = True
小人物頭像移動_電腦.Enabled = True
'==============
FormMainMode.wmpse6.Controls.play
一般系統類.檢查音樂播放 6
戰鬥系統類.時間軸_重設
trtimeline.Enabled = True
trgoi2.Enabled = True
'==============
戰鬥系統類.骰量更新顯示
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'======================================
Erase Vss_EventPlayerAllActionOffNum
'===========================執行階段插入點(ATK-17/DEF-37)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 17, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 37, 2
'============================
If Vss_EventPlayerAllActionOffNum(1) = 1 Then
    For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
        FormMainMode.card(ckl).CardEnabledType = False
    Next
    FormMainMode.bnok.Enabled = False
    目前數(24) = 47
    FormMainMode.等待時間_2.Enabled = True
ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
    For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
        FormMainMode.card(ckl).CardEnabledType = False
    Next
    目前數(24) = 45
    等待時間_2.Enabled = True
End If
End Sub



Private Sub cn22_Click()
'turnpageonin = 0
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
        FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\def2.gif"
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
        FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\def2.gif"
  End If
End If
turnatk = 2
顯示列1.顯示列圖片 = app_path & "gif\system\lineusdef.png"
戰鬥系統類.cleanatkingpagetot
livecom(角色人物對戰人數(2, 2)) = Val(compi4(角色人物對戰人數(2, 2)).Caption)
liveus(角色人物對戰人數(1, 2)) = Val(uspi4(角色人物對戰人數(1, 2)).Caption)
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
Erase atkingckdice
'Erase 異常狀態_混沌紀錄數
'Erase 異常狀態_AI_混沌紀錄數
''=====
'If atkingck(49, 2) = 1 And atking_尤莉卡_超載目前階段紀錄數(3) = 2 Then
'   atkingck(49, 1) = 5
'   技能.尤莉卡_超載 '(階段5)
'End If
'If atkingckai(139, 2) = 1 And atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2 Then
'   atkingckai(139, 1) = 5
'   AI技能.尤莉卡_超載 '(階段5)
'End If
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
'======================電腦方事件卡先出制度
If 電腦方事件卡是否出完選擇數 = False Then
   GoTo 電腦方事件卡先出制度_執行階段2
End If
'================================
電腦方事件卡先出制度_執行階段結束:
'----------以下為電腦判斷出牌程式碼（攻擊方）
'====================試驗智慧型AI出牌系統
'If 智慧型AI系統_目前可執行之人物判斷(namecom(角色人物對戰人數(2, 2))) = True Then
    Dim wtyr As Integer '暫時變數
    If moveturn = 1 Then wtyr = 1 Else wtyr = 0
    智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 2, 1, namecom(角色人物對戰人數(2, 2)), movecp, wtyr
    GoTo 智慧型AI出牌_執行階段結束
'End If
   '============以下是技能檢查及啟動
'    If turnatk = 2 Then
'        atkingckai(1, 1) = 1
'       AI技能.雪莉_自殺傾向 (0)  '(階段1)
'    End If
'    If turnatk = 2 And movecp = 3 Then
'       atkingckai(5, 1) = 1
'       AI技能.雪莉_飛刃雨 '(階段1)
'    End If
'    If turnatk = 2 Then
'       atkingckai(48, 1) = 5
'       AI技能.傑多_因果之幻 '(階段5)
'    End If
'    If turnatk = 2 And movecp < 3 Then
'        atkingckai(11, 1) = 1
'       AI技能.蕾_終曲_無盡輪迴的終結  '(階段1)
'    End If
'    '==========
'    AI人物.CC 1
'    AI人物.史塔夏 1
'    AI人物.庫勒尼西 1
     '==================
If turnatk = 2 And movecp = 1 Then
   戰鬥系統類.comatk1
ElseIf turnatk = 2 And movecp > 1 Then
   戰鬥系統類.comatk2
End If
'==============
'AI人物.艾依查庫 1
'AI人物.艾伯李斯特 1
'AI人物.利恩 1
'AI人物.蕾格烈芙 1
'AI人物.阿奇波爾多 2
'AI人物.瑪格莉特 1
'AI人物.多妮妲 1
''==========
'If moveturn = 1 Then
'    AI人物.全人物通用 2
'End If
'==============================
智慧型AI出牌_執行階段結束:
'==============================
'If compi1(角色人物對戰人數(2, 2)).Caption = "音音夢" Then
'    If atking_AI_音音夢_成長模式狀態數(2) = 1 And turnatk = 2 Then
'       atking_AI_音音夢_成長模式狀態數(1) = 2
'       戰鬥系統類.特殊_音音夢_成長狀態_電腦 '(階段2)
'    End If
'End If
'If atkingckai(5, 2) = 1 Then
'    atkingckai(5, 1) = 2
'    AI技能.雪莉_飛刃雨 '(階段2)
'ElseIf atkingckai(5, 2) = 0 Then
'    atkingckai(5, 1) = 3  '(目標階段3)
'End If
'If turnatk = 2 Then
'    atkingck(128, 1) = 1
'    技能.庫勒尼西_沙漠中的海市蜃樓 '(階段1)
'End If
'=========
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
'turnpageonin = 0
cn32.Visible = False
OK按鈕牌完成移動檢查.Enabled = True
End Sub

Private Sub cn4_Click()
Dim uscomvsn As Integer
cn4.Visible = False
turnatk = 5
'atkingtrn(1) = 0
'atkingtrn(2) = 0
'=================以下是技能檢查及啟動(回合結束階段1)
'If turnatk = 5 And atkingck(34, 2) = 1 Then
'    atkingck(34, 1) = 2
'    技能.CC_白銀戰機 '(階段2)
'End If
'If turnatk = 5 And atkingck(33, 2) = 1 Then
'    atkingck(33, 1) = 2
'    技能.CC_滅菌空間 '(階段2)
'End If
'If turnatk = 5 And atkingckai(103, 2) = 1 Then
'    atkingckai(103, 1) = 2
'    AI技能.CC_滅菌空間 '(階段2)
'End If
'If turnatk = 5 And atkingckai(33, 2) = 1 Then
'    atkingckai(33, 1) = 2
'    AI技能.CC_白銀戰機 '(階段2)
'End If
''=================
'技能動畫顯示階段數 = 7
'戰鬥系統類.技能啟動數量檢查
''=================以下是技能檢查及啟動(回合結束階段2)
''==================
'If turnatk = 5 And atkingck(34, 2) = 1 Then
'    atkingck(34, 1) = 3
'    技能.CC_白銀戰機 '(階段3)
'End If
'If turnatk = 5 And atkingck(33, 2) = 1 Then
'    atkingck(33, 1) = 3
'    技能.CC_滅菌空間 '(階段3)
'End If
''===================
'If turnatk = 5 And atkingckai(103, 2) = 1 Then
'    atkingckai(103, 1) = 3
'    AI技能.CC_滅菌空間 '(階段3)
'End If
'If turnatk = 5 And atkingckai(33, 2) = 1 Then
'    atkingckai(33, 1) = 3
'    AI技能.CC_白銀戰機 '(階段3)
'End If
''=================
'atkingtrtot.Interval = 600
'atkingtrtot.Enabled = True
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

Private Sub Command2_Click()
'MsgBox "劍：" & atkingpagetot(2, 1) & Space(5) & "防：" & atkingpagetot(2, 2) & Space(5) & "移：" & atkingpagetot(2, 3) & Space(5) & "特：" & atkingpagetot(2, 4) & Space(5) & "槍：" & atkingpagetot(2, 5)
'MsgBox livecom & ";" & livecommax
'MsgBox "3:" & 攻擊防禦骰子總數(3) & "     " & "4:" & 攻擊防禦骰子總數(4)
'智慧型AI系統類.智慧型AI系統計算_引導程序_試驗1 1, 1, "艾伯李斯特", 1
'智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 1, 3, "艾伯李斯特", 3, 0
End Sub



Private Sub Form_Unload(Cancel As Integer)
wmp.Close
wmpse1.Close
wmpse2.Close
wmpse3.Close
wmpse4.Close
wmpse5.Close
wmpse6.Close
wmpse7.Close
wmpse8.Close
End Sub

Private Sub messageus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'For i = 1 To 公用牌實體卡片分隔紀錄數(1)
'   card(i).card_MouseExit
'Next
'For i = 1 To 3
'    PEAFcardusbackclick(i).Visible = False
'Next
'PEAFcardback(1).Visible = False
'For i = 2 To 3
'     cardus(角色待機人物紀錄數(1, i)).Visible = False
'Next
'If turnpageonin = 1 Then
'    bnok.Picture = LoadPicture(app_path & "gif\system\ok_1.jpg")
'End If
End Sub
Private Sub NextTurn_階段2_Timer()
Dim uscomvsn As Integer
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
'liveus(角色人物對戰人數(1, 2)) = Val(usbi1(角色人物對戰人數(1, 2)).Caption)
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
    FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\stageblack.gif"
End If
顯示列1.顯示列圖片 = app_path & "gif\system\DRAW.png"
'cn4.Visible = False
'cn1.Visible = True
'==============
小人物頭像移動方向數(1) = 2
小人物頭像移動方向數(2) = 2
小人物頭像移動_使用者.Enabled = True
小人物頭像移動_電腦.Enabled = True
'==============
目前數(24) = 1
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
           移動階段_階段前啟動.Enabled = True
   End Select
End If
End Sub

Private Sub pagecomglead_Change()
pageglead(2) = Val(pagecomglead.Caption)
End Sub

Sub pagecomqlead_Change()
'atkingckai(26, 1) = 1
'atkingckai(98, 1) = 1
'atkingckai(12, 1) = 2
'atkingckai(82, 1) = 2
''============以下是技能檢查
'If turnatk = 2 And 階段狀態數 = 3 Then
'    AI技能.雪莉_飛刃雨 '(階段3/4)
'    AI技能.蕾_終曲_無盡輪迴的終結 '(階段1)
'    AI技能.古魯瓦爾多_猛擊  '(階段1)
'    AI技能.蕾_輪旋曲_琉璃色的微風 '(階段1)
'    AI技能.南瓜王_重壓 '(階段1)
'    AI技能.妖精王妃_冰結之翼 '(階段1)
'    AI技能.蕾_EX_輪旋曲_琉璃色的微風 '(階段1)
'    AI技能.吸血姬蕾米雅_吸血  '(階段1)
'    AI技能.吸血姬蕾米雅_高貴的晚餐 '(階段1)
'    AI技能.艾伯李斯特_精密射擊  '(階段1)
'    AI技能.史塔夏_愚者之手  '(階段1)
'    AI技能.史塔夏_命運的鐵門  '(階段1)
'    AI技能.阿貝爾_霸王閃擊  '(階段1)
'    AI技能.阿貝爾_幻影劍舞  '(階段1)
'    AI技能.布勞_時間爆彈  '(階段1)
'    AI技能.艾依查庫_連射  '(階段1)
'    AI技能.艾依查庫_神速之劍 (0) '(階段1)
'    AI技能.梅倫_Gamble '(階段1)
'    AI技能.羅莎琳_染血之刃 '(階段1)
'    AI技能.CC_白銀戰機 '(階段1)
'    AI技能.帕茉_戰慄的狼牙 '(階段1)
'    AI技能.帕茉_慈悲的藍眼 '(階段1)
'    AI技能.帕茉_靜謐之背 '(階段1)
'    AI技能.艾茵_十三隻眼 '(階段1)
'    AI技能.夏洛特_冬之夢 '(階段1)
'    AI技能.泰瑞爾_Chr_799 '(階段1)
'    AI技能.泰瑞爾_Rud_913 '(階段1)
'    AI技能.泰瑞爾_Wil_846 '(階段1)
'    AI技能.瑪格莉特_地獄獵心獸 '(階段1)
'    AI技能.傑多_因果之幻 '(階段1)
'    AI技能.CC_高頻電磁手術刀 '(階段1)
'    AI技能.伊芙琳_紅蓮車輪 '(階段1)
'    AI技能.多妮妲_律死擊 '(階段1)
'    AI技能.多妮妲_殘虐傾向 '(階段1)
'    AI技能.庫勒尼西_深淵 '(階段1)
'    AI技能.羅莎琳_黑霧的纏繞 '(階段1)
'    AI技能.梅倫_Lowball '(階段1)
'    AI技能.艾伯李斯特_雷擊 '(階段1)
'    AI技能.艾依查庫_憤怒一擊 '(階段1)
'    AI技能.阿貝爾_閃電旋風刺 '(階段1)
'    AI技能.利恩_劫影攻擊 '(階段1)
'    AI技能.利恩_毒牙 '(階段1)
'    AI技能.利恩_背刺 '(階段1)
'    AI技能.瑪格莉特_月光 '(階段1)
'    AI技能.蕾格烈芙_CTL '(階段1)
'    AI技能.蕾格烈芙_BPA '(階段1)
'    AI技能.阿奇波爾多_致命槍擊 '(階段1)
'    AI技能.阿奇波爾多_劫影攻擊 '(階段1)
'    AI技能.洛洛妮_砲擊壓制 '(階段1)
'    AI技能.洛洛妮_貪婪之刃與嗜血之槍  '(階段1)
'    AI技能.艾蕾可_聖王威光  '(階段1)
'    AI技能.露緹亞_腐朽之靈  '(階段1)
'    AI技能.露緹亞_渦騎劍閃 (0) '(階段1)
'    AI技能.梅莉_夢幻魔杖 '(階段1)
'    AI技能.梅莉_夢境搖籃 '(階段1)
'    AI技能.貝琳達_裂地冰牙 '(階段1)
'    AI技能.貝琳達_溶魂之雨 '(階段1)
'    AI技能.蕾_EX_終曲_無盡輪迴的終結 '(階段1)
'    AI技能.克頓_隱蔽射擊 '(階段1)
'    AI技能.尤莉卡_奸佞的鐵鎚 '(階段1)
'    AI技能.羅莎琳_EX_染血之刃 '(階段1)
'ElseIf turnatk = 3 And 階段狀態數 = 3 Then
'    AI技能.雪莉_巨大黑犬 '(階段2)
'    AI技能.南瓜王_超再生 '(階段1)
'    AI技能.妖精王妃_混沌之翼 '(階段1)
'    AI技能.音音夢_美味牛奶 '(階段1)
'    AI技能.艾伯李斯特_智略 '(階段1)
'    AI技能.史塔夏_殺戮器官 '(階段1)
'    AI技能.阿奇波爾多_大地崩壞 '(階段1)
'    AI技能.艾蕾可_救濟天使  '(階段1)
'    AI技能.CC_滅菌空間  '(階段1)
'    AI技能.梅莉_綿羊幻夢  '(階段1)
'    AI技能.古魯瓦爾多_必殺架勢  '(階段1)
'    AI技能.古魯瓦爾多_精神力吸收  '(階段1)
'    AI技能.帕茉_憤怒之爪  '(階段1)
'    AI技能.伊芙琳_怠惰的墓表 '(階段1)
'    AI技能.伊芙琳_赤紅石榴 '(階段1)
'    AI技能.布勞_發條機構 '(階段1)
'    AI技能.布勞_夜幕時分 '(階段1)
'    AI技能.阿貝爾_抽刀斷水計 '(階段1)
'    AI技能.夏洛特_夜未央 '(階段1)
'    AI技能.夏洛特_幸福的理由 '(階段1)
'    AI技能.瑪格莉特_末日幻影 '(階段1)
'    AI技能.蕾格烈芙_SSS '(階段1)
'    AI技能.多妮妲_超級女主角 '(階段1)
'    AI技能.傑多_因果之線 '(階段1)
'    AI技能.貝琳達_雪光 '(階段1)
'    AI技能.洛洛妮_逆轉戰局的槍響 '(階段1)
'    AI技能.克頓_惡意情報 '(階段1)
'    AI技能.艾茵_一顆心 '(階段1)
'    AI技能.尤莉卡_超載 '(階段1)
'ElseIf turnatk = 1 And 階段狀態數 = 3 Then
'    AI技能.雪莉_異質者 '(階段2)
'    AI技能.妖精王妃_煉獄之翼 '(階段1)
'    AI技能.吸血姬蕾米雅_消失 '(階段1)
'    AI技能.艾依查庫_不屈之心 '(階段1)
'    AI技能.音音夢_溫柔注射 '(階段1)
'    AI技能.梅倫_Jackpot '(階段1)
'    AI技能.艾茵_兩個身體 '(階段1)
'    AI技能.瑪格莉特_恍惚 '(階段1)
'    AI技能.庫勒尼西_黑暗漩渦 '(階段1)
'    AI技能.蕾格烈芙_LAR '(階段1)
'    AI技能.蕾_協奏曲_加百烈的守護 '(階段1)
'    AI技能.史塔夏_時間種子 '(階段1)
'    AI技能.艾茵_九個靈魂 '(階段1)
'    AI技能.CC_原子之心 '(階段1)
'    AI技能.阿奇波爾多_防護射擊 '(階段1)
'    AI技能.蕾_EX_協奏曲_加百烈的守護 '(階段1)
'    AI技能.羅莎琳_咀咒的刻印  '(階段1)
'    AI技能.伊芙琳_慟哭之歌  '(階段1)
'    AI技能.古魯瓦爾多_血之恩賜  '(階段1)
'    AI技能.蕾_EX_安魂曲_死神的鎮魂歌 '(階段1)
'    AI技能.梅倫_High_hand '(階段1)
'    AI技能.艾伯李斯特_茨林 '(階段1)
'    AI技能.布勞_時間追獵 '(階段1)
'    AI技能.利恩_反擊的狼煙 '(階段1)
'    AI技能.泰瑞爾_Von_541 '(階段1)
'    AI技能.庫勒尼西_瘋狂眼窩 '(階段1)
'    AI技能.多妮妲_異質者 '(階段2)
'    AI技能.洛洛妮_風暴感知 '(階段1)
'    AI技能.夏洛特_大聖堂 '(階段1)
'    AI技能.艾蕾可_王座之炎 '(階段1/2)
'    AI技能.艾蕾可_白百合 '(階段1)
'    AI技能.露緹亞_朦朧之暗 '(階段1)
'    AI技能.露緹亞_暗影之翼 '(階段1)
'    AI技能.梅莉_徬徨夢羽  '(階段1)
'    AI技能.音音夢_秘密苦藥  '(階段1)
'    AI技能.傑多_因果之輪 '(階段1)
'    AI技能.傑多_因果之刻 '(階段1)
'    AI技能.貝琳達_水晶幻鏡 '(階段1)
'    AI技能.蕾_安魂曲_死神的鎮魂歌 '(階段1)
'    AI技能.羅莎琳_黑霧幻影 '(階段1)
'    AI技能.羅莎琳_EX_黑霧幻影 '(階段1)
'    AI技能.克頓_竊取資料 '(階段1)
'    AI技能.克頓_逃亡計畫 '(階段1)
'    AI技能.尤莉卡_不善的信仰 '(階段1)
'    AI技能.尤莉卡_曲惡的安寧 '(階段1)
'End If
'==================
'===========================執行階段插入點(43-2)
'執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 3
''============================
'戰鬥系統類.骰量更新顯示
End Sub

Private Sub pageusglead_Change()
pageglead(1) = Val(pageusglead.Caption)
End Sub

Private Sub pageusqlead_Change()
'atkingck(79, 1) = 1
'atkingck(101, 1) = 1
''============以下是技能檢查
'If turnatk = 1 And 階段狀態數 = 1 Then
'   技能.雪莉_飛刃雨 '(階段1/2)
'   技能.雪莉_VBE_飛刃雨 '(階段1/2)
'   技能.古魯瓦爾多_猛擊 '(階段1)
'   技能.帕茉_慈悲的藍眼 '(階段1)
'   技能.帕茉_靜謐之背 '(階段1)
'   技能.蕾_輪旋曲_琉璃色的微風 '(階段1)
'   技能.蕾_EX_輪旋曲_琉璃色的微風 '(階段1)
'   技能.蕾_終曲_無盡輪迴的終結 '(階段1)
'   技能.艾茵_十三隻眼 '(階段1)
'   技能.帕茉_戰慄的狼牙 '(階段1)
'   技能.史塔夏_愚者之手 '(階段1)
'   技能.史塔夏_命運的鐵門 '(階段1)
'   技能.CC_白銀戰機 '(階段1)
'   技能.CC_高頻電磁手術刀 '(階段1)
'   技能.羅莎琳_染血之刃 '(階段1)
'   技能.羅莎琳_黑霧的纏繞 '(階段1)
'   技能.伊芙琳_紅蓮車輪 '(階段1)
'   技能.梅倫_Lowball '(階段1)
'   技能.梅倫_Gamble '(階段1)
'   技能.艾伯李斯特_精密射擊 '(階段1)
'   技能.艾伯李斯特_雷擊 '(階段1)
'   技能.艾依查庫_連射 '(階段1)
'   技能.艾依查庫_神速之劍 (0) '(階段1)
'   技能.艾依查庫_憤怒一擊 '(階段1)
'   技能.布勞_時間爆彈 '(階段1)
'   技能.阿貝爾_霸王閃擊 '(階段1)
'   技能.阿貝爾_閃電旋風刺  '(階段1)
'   技能.阿貝爾_幻影劍舞  '(階段1)
'   技能.利恩_劫影攻擊  '(階段1)
'   技能.利恩_毒牙  '(階段1)
'   技能.利恩_背刺  '(階段1)
'   技能.夏洛特_冬之夢  '(階段1)
'   技能.泰瑞爾_Rud_913  '(階段1)
'   技能.泰瑞爾_Chr_799  '(階段1)
'   技能.泰瑞爾_Wil_846  '(階段1)
'   技能.瑪格莉特_月光  '(階段1)
'   技能.瑪格莉特_地獄獵心獸  '(階段1)
'   技能.庫勒尼西_深淵  '(階段1)
'   技能.蕾格烈芙_CTL  '(階段1)
'   技能.蕾格烈芙_BPA  '(階段1)
'   技能.多妮妲_殘虐傾向  '(階段1)
'   技能.多妮妲_律死擊 '(階段1)
'   技能.傑多_因果之幻 '(階段1)
'   技能.阿奇波爾多_致命槍擊 '(階段1)
'   技能.阿奇波爾多_劫影攻擊 '(階段1)
'   技能.洛洛妮_砲擊壓制 '(階段1)
'   技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段1)
'   技能.克頓_隱蔽射擊 '(階段1)
'   技能.露緹亞_腐朽之靈  '(階段1)
'   技能.露緹亞_渦騎劍閃 (0) '(階段1)
'   技能.艾蕾可_聖王威光  '(階段1)
'   技能.梅莉_夢幻魔杖  '(階段1)
'   技能.梅莉_夢境搖籃  '(階段1)
'   技能.貝琳達_裂地冰牙 '(階段1)
'   技能.貝琳達_溶魂之雨 '(階段1)
'   技能.蕾_EX_終曲_無盡輪迴的終結 '(階段1)
'   技能.尤莉卡_奸佞的鐵鎚 '(階段1)
'   技能.羅莎琳_EX_染血之刃 '(階段1)
'ElseIf turnatk = 3 And 階段狀態數 = 1 Then
'   技能.雪莉_巨大黑犬 '(階段1)
'   技能.雪莉_VBE_巨大黑犬 '(階段1)
'   技能.帕茉_憤怒之爪 '(階段1)
'   技能.古魯瓦爾多_必殺架勢 '(階段1)
'   技能.史塔夏_殺戮器官 '(階段1)
'   技能.CC_滅菌空間 '(階段1)
'   技能.艾茵_一顆心 '(階段1)
'   技能.伊芙琳_怠惰的墓表 '(階段1)
'   技能.伊芙琳_赤紅石榴 '(階段1)
'   技能.古魯瓦爾多_精神力吸收 '(階段1)
'   技能.音音夢_美味牛奶 '(階段1)
'   技能.艾伯李斯特_智略 '(階段1)
'   技能.布勞_發條機構 '(階段1)
'   技能.布勞_夜幕時分 '(階段1)
'   技能.阿貝爾_抽刀斷水計 '(階段1)
'   技能.夏洛特_夜未央 '(階段1)
'   技能.夏洛特_幸福的理由 '(階段1)
'   技能.瑪格莉特_末日幻影 '(階段1)
'   技能.蕾格烈芙_SSS '(階段1)
'   技能.多妮妲_超級女主角 '(階段1)
'   技能.傑多_因果之線 '(階段1)
'   技能.阿奇波爾多_大地崩壞 '(階段1)
'   技能.洛洛妮_逆轉戰局的槍響 '(階段1)
'   技能.克頓_惡意情報 '(階段1)
'   技能.艾蕾可_救濟天使 '(階段1)
'   技能.梅莉_綿羊幻夢 '(階段1)
'   技能.貝琳達_雪光 '(階段1)
'   技能.尤莉卡_超載 '(階段1)
'ElseIf turnatk = 2 And 階段狀態數 = 1 Then
'   技能.雪莉_異質者 '(階段1)
'   技能.雪莉_VBE_異質者 '(階段1)
'   技能.蕾_協奏曲_加百烈的守護 '(階段1)
'   技能.蕾_安魂曲_死神的鎮魂歌 '(階段1)
'   技能.蕾_EX_安魂曲_死神的鎮魂歌 '(階段1)
'   技能.史塔夏_時間種子 '(階段1)
'   技能.艾茵_九個靈魂 '(階段1)
'   技能.艾茵_兩個身體 '(階段1)
'   技能.CC_原子之心 '(階段1)
'   技能.蕾_EX_協奏曲_加百烈的守護 '(階段1)
'   技能.羅莎琳_咀咒的刻印 '(階段1)
'   技能.羅莎琳_黑霧幻影 '(階段1)
'   技能.羅莎琳_EX_黑霧幻影 '(階段1)
'   技能.伊芙琳_慟哭之歌 '(階段1)
'   技能.古魯瓦爾多_血之恩賜 '(階段1)
'   技能.梅倫_High_hand '(階段1)
'   技能.梅倫_Jackpot '(階段1)
'   技能.音音夢_溫柔注射 '(階段1)
'   技能.音音夢_秘密苦藥 '(階段1)
'   技能.艾伯李斯特_茨林 '(階段1)
'   技能.艾依查庫_不屈之心 '(階段1)
'   技能.布勞_時間追獵 '(階段1)
'   技能.利恩_反擊的狼煙  '(階段1)
'   技能.夏洛特_大聖堂  '(階段1)
'   技能.泰瑞爾_Von_541  '(階段1)
'   技能.瑪格莉特_恍惚  '(階段1)
'   技能.庫勒尼西_瘋狂眼窩 '(階段1)
'   技能.庫勒尼西_黑暗漩渦 '(階段1)
'   技能.蕾格烈芙_LAR '(階段1)
'   技能.多妮妲_異質者 '(階段1)
'   技能.傑多_因果之輪 '(階段1)
'   技能.傑多_因果之刻 '(階段1)
'   技能.阿奇波爾多_防護射擊 '(階段1)
'   技能.洛洛妮_風暴感知 '(階段1)
'   技能.克頓_竊取資料 '(階段1)
'   技能.克頓_逃亡計畫 '(階段1)
'   技能.露緹亞_朦朧之暗 '(階段1)
'   技能.露緹亞_暗影之翼 '(階段1)
'   技能.艾蕾可_王座之炎 '(階段1/2)
'   技能.艾蕾可_白百合  '(階段1)
'   技能.梅莉_徬徨夢羽  '(階段1)
'   技能.貝琳達_水晶幻鏡  '(階段1)
'   技能.尤莉卡_不善的信仰 '(階段1)
'   技能.尤莉卡_曲惡的安寧 '(階段1)
'End If
'==================
'Select Case turnatk
'    Case 1
'        '===========================執行階段插入點(ATK-42/DEF-43)
'        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
'        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
'        '============================
'    Case 2
'        '===========================執行階段插入點(ATK-42/DEF-43)
'        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
'        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
'        '============================
'    Case 3
'        '===========================執行階段插入點(44)
'        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 3
'        '============================
'End Select
'戰鬥系統類.骰量更新顯示

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
         FormMainMode.wmp.Controls.play
    Case 50
         bnreturn.Visible = True
         bnreturnt.Visible = True
         bn.Visible = True
         bnt.Visible = True
         PEAEtr1.Enabled = False
         '======================
         戰鬥系統類.遊戲對戰結束物件消滅
         '======================
End Select
PEAEtr1num = PEAEtr1num + 1
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
For i = 1 To 公用牌實體卡片分隔紀錄數(1)
   card(i).CardEventType = False
Next
For i = 1 To 3
  cardcom(i).Visible = False
'  PEAFcardusbackclick(i).Visible = False
'  PEAFcardcombackclick(i).Visible = False
Next
For i = 1 To 3
  If i <> 角色人物對戰人數(1, 2) Then
     cardus(i).Visible = False
  End If
Next
atkinghelpc.Visible = False
'PEAFcardback(1).Visible = False
If turnpageonin = 1 Then
    bnok.Picture = LoadPicture(app_path & "gif\system\ok_1.jpg")
End If
End Sub

Sub PEMtr1_Timer()
Select Case 音樂檢查播放目標數
     Case 0
         If Left(FormMainMode.wmp.Status, 2) = "就緒" Then
             wmp.Controls.play
         ElseIf Left(FormMainMode.wmp.Status, 2) = "播放" Then
             PEMtr1.Enabled = False
         End If
     Case 1
         If Left(FormMainMode.wmpse1.Status, 2) = "就緒" Then
             wmpse1.Controls.play
         ElseIf Left(FormMainMode.wmpse1.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse1.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 2
         If Left(FormMainMode.wmpse2.Status, 2) = "就緒" Then
             wmpse2.Controls.play
         ElseIf Left(FormMainMode.wmpse2.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse2.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 3
         If Left(FormMainMode.wmpse3.Status, 2) = "就緒" Then
             wmpse3.Controls.play
         ElseIf Left(FormMainMode.wmpse3.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse3.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 4
         If Left(FormMainMode.wmpse4.Status, 2) = "就緒" Then
             wmpse4.Controls.play
         ElseIf Left(FormMainMode.wmpse4.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse4.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 5
         If Left(FormMainMode.wmpse5.Status, 2) = "就緒" Then
             wmpse5.Controls.play
         ElseIf Left(FormMainMode.wmpse5.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse5.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 6
         If Left(FormMainMode.wmpse6.Status, 2) = "就緒" Then
             wmpse6.Controls.play
         ElseIf Left(FormMainMode.wmpse6.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse6.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 7
         If Left(FormMainMode.wmpse7.Status, 2) = "就緒" Then
             wmpse7.Controls.play
         ElseIf Left(FormMainMode.wmpse7.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse7.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
    Case 8
         If Left(FormMainMode.wmpse8.Status, 2) = "就緒" Then
             wmpse8.Controls.play
         ElseIf Left(FormMainMode.wmpse8.Status, 2) = "播放" Or _
         Left(FormMainMode.wmpse8.Status, 3) = "已停止" Then
             PEMtr1.Enabled = False
         End If
'    Case 9
'         If Left(FormMainMode.wmpse9.Status, 2) = "就緒" Then
'             wmpse9.Controls.play
'         ElseIf Left(FormMainMode.wmpse9.Status, 2) = "播放" Or _
'         Left(FormMainMode.wmpse9.Status, 3) = "已停止" Then
'             PEMtr1.Enabled = False
'         End If
End Select
End Sub

Private Sub personatk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 公用牌實體卡片分隔紀錄數(1)
   card(i).CardEventType = False
Next
'=========
戰鬥系統類.技能說明載入_使用者 Index
'=========
atkinghelpc.Left = atkinghelpxy(1, Index, 1)
atkinghelpc.Top = atkinghelpxy(1, Index, 2)
atkinghelpc.ZOrder
atkinghelpc.Visible = True
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
'   Formend.Left = FormMainMode.Left
'   Formend.Top = FormMainMode.Top
'   FormMainMode.Visible = False
'   Formend.Visible = True
   一般系統類.主選單_PEAttackingEndingForm顯示
   PEAttackingForm.Visible = False
   PEAEtr1num = 0
   PEAEtr1.Enabled = True
   trend.Enabled = False
ElseIf trend暫時變數 = 2 Then
   FormMainMode.wmp.Controls.stop
'   FormMainMode.wmp.Controls.pause
   FormMainMode.wmp.settings.playCount = 1
   FormMainMode.wmp.URL = app_path & "mp3\ulse15.mp3"
   FormMainMode.wmp.Controls.stop
   FormMainMode.wmp.settings.playCount = 1
   trend暫時變數 = trend暫時變數 + 1
Else
   trend暫時變數 = trend暫時變數 + 1
End If
End Sub

Sub trgoi1_Timer()
'If Val(pageusqlead) = 0 And turnatk = 1 And 階段狀態數 = 1 Then
'    攻擊防禦骰子總數(1) = 0
'    攻擊防禦骰子總數(3) = 0
'    goicheck(1) = 0
'End If
'If Val(pageusqlead) = 0 And turnatk = 2 And 階段狀態數 = 1 Then
'    攻擊防禦骰子總數(1) = 0
'    攻擊防禦骰子總數(3) = 0
'    goidefus = 0
'    戰鬥系統類.chkdef
'End If
'
'If atkingpagetot(1, 1) = 0 And turnatk = 1 And movecp = 1 And goicheck(1) = 1 And 階段狀態數 = 1 Then
'   goicheck(1) = 0
'   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - atkus(角色人物對戰人數(1, 2))
'   攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - atkus(角色人物對戰人數(1, 2))
  '=========以下是技能檢查及發動(當前檢查類型：劍)
'   If 異常狀態檢查數(13, 2) = 1 Then
'      異常狀態檢查數(13, 1) = 2
'      異常狀態.聖痕_使用者 '(階段2)
''   End If
''   If 異常狀態檢查數(24, 2) = 1 Then
'      異常狀態檢查數(24, 1) = 2
'      異常狀態.能力低下_使用者 '(階段2)
''   End If
''   If 異常狀態檢查數(7, 2) = 1 Then
'      異常狀態檢查數(7, 1) = 3
'      異常狀態.ATK加_使用者 '(階段3)
''   End If
''   If 異常狀態檢查數(10, 2) = 1 Then
'      異常狀態檢查數(10, 1) = 3
'      異常狀態.ATK減_使用者 '(階段3)
''   End If
'      異常狀態檢查數(39, 1) = 2
'      異常狀態.臨界_使用者 '(階段2)
   '==============
'   If 攻擊防禦骰子總數(1) < 0 Then 攻擊防禦骰子總數(1) = 0
'End If
'If atkingpagetot(1, 5) = 0 And turnatk = 1 And movecp > 1 And goicheck(1) = 1 And 階段狀態數 = 1 Then
'   goicheck(1) = 0
'   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - atkus(角色人物對戰人數(1, 2))
'   攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - atkus(角色人物對戰人數(1, 2))
  '=========以下是技能檢查及發動(當前檢查類型：槍)
''   If 異常狀態檢查數(13, 2) = 1 Then
'      異常狀態檢查數(13, 1) = 2
'      異常狀態.聖痕_使用者 '(階段2)
''   End If
''   If 異常狀態檢查數(24, 2) = 1 Then
'      異常狀態檢查數(24, 1) = 2
'      異常狀態.能力低下_使用者 '(階段2)
''   End If
''   If 異常狀態檢查數(7, 2) = 1 Then
'      異常狀態檢查數(7, 1) = 3
'      異常狀態.ATK加_使用者 '(階段3)
''   End If
''   If 異常狀態檢查數(10, 2) = 1 Then
'      異常狀態檢查數(10, 1) = 3
'      異常狀態.ATK減_使用者 '(階段3)
''   End If
'      異常狀態檢查數(39, 1) = 2
'      異常狀態.臨界_使用者 '(階段2)
   '==============
'   If 攻擊防禦骰子總數(1) < 0 Then 攻擊防禦骰子總數(1) = 0
'End If
'If turnatk = 1 And movecp = 1 Then
' 戰鬥系統類.chkus1
'ElseIf turnatk = 1 And movecp > 1 Then
' 戰鬥系統類.chkus2
'End If
'=========以下是技能檢查及發動
'If atkingck(9, 2) = 1 And turnatk = 1 And 階段狀態數 = 1 Then
'   atkingck(9, 1) = 2
'   技能.帕茉_慈悲的藍眼 '(階段2)
'ElseIf atkingck(9, 2) = 0 And turnatk = 1 And atking_帕茉_慈悲的藍眼_tot(2) = 1 And 階段狀態數 = 1 Then
'   atkingck(9, 1) = 3
'   技能.帕茉_慈悲的藍眼 '(階段3)
'End If
'If atkingckai(37, 2) = 1 And turnatk = 2 And 階段狀態數 = 1 Then
'   atkingckai(37, 1) = 5
'   AI技能.艾茵_十三隻眼 '(階段5)
'End If
'If atkingck(16, 2) = 1 And turnatk = 1 And 階段狀態數 = 1 Then
'   atkingck(16, 1) = 2
'   技能.艾茵_十三隻眼 '(階段2)
'ElseIf atkingck(16, 2) = 0 And turnatk = 1 And atking_艾茵_十三隻眼_tot(2) = 1 And 階段狀態數 = 1 Then
'   atkingck(16, 1) = 3
'   技能.艾茵_十三隻眼 '(階段3)
'End If
'If uspi1(角色人物對戰人數(1, 2)).Caption = "史塔夏" Then
'    If atking_史塔夏_殺戮模式狀態數(2) = 1 And turnatk = 1 And 階段狀態數 = 1 And 攻擊防禦骰子總數(1) > 0 Then
'       atking_史塔夏_殺戮模式狀態數(1) = 1
'       戰鬥系統類.特殊_史塔夏_殺戮狀態_使用者 '(階段1)
'    ElseIf atking_史塔夏_殺戮模式狀態數(2) = 1 And turnatk = 1 And 階段狀態數 = 1 And 攻擊防禦骰子總數(1) = 0 Then
'       atking_史塔夏_殺戮模式狀態數(1) = 2
'       戰鬥系統類.特殊_史塔夏_殺戮狀態_使用者 '(階段2)
'    End If
'End If
'If uspi1(角色人物對戰人數(1, 2)).Caption = "音音夢" Then
'    If atking_音音夢_成長模式狀態數(2) = 1 And turnatk = 1 And 階段狀態數 = 1 And 攻擊防禦骰子總數(1) = 0 Then
'       atking_音音夢_成長模式狀態數(1) = 1
'       戰鬥系統類.特殊_音音夢_成長狀態_使用者 '(階段1)
'    End If
'End If
''======
'If 異常狀態_混沌紀錄數(3) = 1 And turnatk = 1 And 階段狀態數 = 1 And 攻擊防禦骰子總數(1) = 0 Then
'    異常狀態檢查數(31, 1) = 3
'    異常狀態.混沌_使用者 '(階段3)
'Else
'    異常狀態檢查數(31, 1) = 1
'    異常狀態.混沌_使用者 '(階段1)
'End If
''======
'If atking_尤莉卡_超載目前階段紀錄數(3) = 2 And atkingck(49, 2) = 1 Then
'    If atking_尤莉卡_超載目前階段紀錄數(4) = 1 And turnatk = 1 And 階段狀態數 = 1 And 攻擊防禦骰子總數(1) = 0 Then
'        atkingck(49, 1) = 5
'        技能.尤莉卡_超載 '(階段5)
'    Else
'        atkingck(49, 1) = 4
'        技能.尤莉卡_超載 '(階段4)
'    End If
'End If
'=========更新骰子總數量表示
If 攻擊防禦骰子總數(1) < 0 Then
   顯示列1.goi1 = 0
Else
   顯示列1.goi1 = 攻擊防禦骰子總數(1)
End If
FormMainMode.trgoi1.Enabled = False
'=====================

'=====================
End Sub

Sub trgoi2_Timer()
'If Val(pagecomqlead) = 0 And turnatk = 2 And 階段狀態數 = 3 Then
' 攻擊防禦骰子總數(2) = 0
' 攻擊防禦骰子總數(4) = 0
' goicheck(2) = 0
'End If
'
'If Val(pagecomqlead) = 0 And turnatk = 1 And 階段狀態數 = 3 Then
'    攻擊防禦骰子總數(2) = 0
'    攻擊防禦骰子總數(4) = 0
'    chkcomck = 0
'    戰鬥系統類.chkdefcom
'End If
''================
'If atkingpagetot(2, 1) = 0 And turnatk = 2 And movecp = 1 And goicheck(2) = 1 And 階段狀態數 = 3 Then
'   goicheck(2) = 0
'   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - atkcom(角色人物對戰人數(2, 2))
'   攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - atkcom(角色人物對戰人數(2, 2))
  '=========以下是技能檢查及發動(當前檢查類型：劍)
'      異常狀態檢查數(26, 1) = 2
'      異常狀態.聖痕_電腦 '(階段2)
'      '=========
'      異常狀態檢查數(1, 1) = 3
'      異常狀態.ATK加_電腦 '(階段3)
'      '=========
'      異常狀態檢查數(4, 1) = 3
'      異常狀態.ATK減_電腦 '(階段3)
'      '=========
'      異常狀態檢查數(25, 1) = 2
'      異常狀態.能力低下_電腦 '(階段2)
'End If
'If atkingpagetot(2, 5) = 0 And turnatk = 2 And movecp > 1 And goicheck(2) = 1 And 階段狀態數 = 3 Then
'   goicheck(2) = 0
'   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - atkcom(角色人物對戰人數(2, 2))
'   攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - atkcom(角色人物對戰人數(2, 2))
  '=========以下是技能檢查及發動(當前檢查類型：槍)
'      異常狀態檢查數(26, 1) = 2
'      異常狀態.聖痕_電腦 '(階段2)
'      '=======
'      異常狀態檢查數(1, 1) = 3
'      異常狀態.ATK加_電腦 '(階段3)
'      '=======
'      異常狀態檢查數(4, 1) = 3
'      異常狀態.ATK減_電腦 '(階段3)
'      '=========
'      異常狀態檢查數(25, 1) = 2
'      異常狀態.能力低下_電腦 '(階段2)
'End If
'If turnatk = 2 Then
' 戰鬥系統類.chkcom
'End If
'=========以下是技能檢查及發動
'If atkingckai(14, 2) = 0 And turnatk = 2 And 階段狀態數 = 3 Then
'   atkingckai(14, 1) = 1
'   AI技能.羊角獸2012_致命衝撞 '(階段1)
'End If
'If atkingckai(15, 2) = 0 And turnatk = 1 And 階段狀態數 = 3 Then
'   atkingckai(15, 1) = 1
'   AI技能.羊角獸2012_致命格擋 '(階段1)
'End If
'================
'If turnatk = 1 And atkingck(19, 2) = 1 And 階段狀態數 = 3 Then
'    atkingck(19, 1) = 4
'    技能.蕾_EX_輪旋曲_琉璃色的微風  '(階段4)
'End If
'If turnatk = 1 And atkingck(13, 2) = 1 And 階段狀態數 = 3 Then
'    atkingck(13, 1) = 4
'    技能.蕾_輪旋曲_琉璃色的微風  '(階段4)
'End If
''================
'If atkingckai(35, 2) = 1 And turnatk = 2 And 階段狀態數 = 3 Then
'   atkingckai(35, 1) = 2
'   AI技能.帕茉_慈悲的藍眼 '(階段2)
'ElseIf atkingckai(35, 2) = 0 And turnatk = 2 And atking_AI_帕茉_慈悲的藍眼_tot(2) = 1 And 階段狀態數 = 3 Then
'   atkingckai(35, 1) = 3
'   AI技能.帕茉_慈悲的藍眼 '(階段3)
'End If
'If atkingck(16, 2) = 1 And turnatk = 1 And 階段狀態數 = 3 Then
'   atkingck(16, 1) = 5
'   技能.艾茵_十三隻眼 '(階段5)
'End If
'If atkingckai(37, 2) = 1 And turnatk = 2 And 階段狀態數 = 3 Then
'   atkingckai(37, 1) = 2
'   AI技能.艾茵_十三隻眼 '(階段2)
'ElseIf atkingckai(37, 2) = 0 And turnatk = 2 And atking_AI_艾茵_十三隻眼_tot(2) = 1 And 階段狀態數 = 3 Then
'   atkingckai(37, 1) = 3
'   AI技能.艾茵_十三隻眼 '(階段3)
'End If
'If compi1(角色人物對戰人數(2, 2)).Caption = "史塔夏" Then
'    If atking_AI_史塔夏_殺戮模式狀態數(2) = 1 And turnatk = 2 And 階段狀態數 = 3 And 攻擊防禦骰子總數(2) > 0 Then
'       atking_AI_史塔夏_殺戮模式狀態數(1) = 1
'       戰鬥系統類.特殊_史塔夏_殺戮狀態_電腦 '(階段1)
'    ElseIf atking_AI_史塔夏_殺戮模式狀態數(2) = 1 And turnatk = 2 And 階段狀態數 = 3 And 攻擊防禦骰子總數(2) = 0 Then
'       atking_AI_史塔夏_殺戮模式狀態數(1) = 2
'       戰鬥系統類.特殊_史塔夏_殺戮狀態_電腦 '(階段2)
'    End If
'End If
'If compi1(角色人物對戰人數(2, 2)).Caption = "音音夢" Then
'    If atking_AI_音音夢_成長模式狀態數(2) = 1 And turnatk = 2 And 階段狀態數 = 3 And 攻擊防禦骰子總數(2) = 0 Then
'       atking_AI_音音夢_成長模式狀態數(1) = 1
'       戰鬥系統類.特殊_音音夢_成長狀態_電腦 '(階段1)
'    End If
'End If
''=============
'異常狀態檢查數(32, 1) = 1
'異常狀態.混沌_電腦  '(階段1)
''=============
'If atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2 And atkingckai(139, 2) = 1 Then
'    atkingckai(139, 1) = 4
'    AI技能.尤莉卡_超載 '(階段4)
'End If
'=========更新骰子總數量表示
If 攻擊防禦骰子總數(2) < 0 Then
   顯示列1.goi2 = 0
Else
   顯示列1.goi2 = 攻擊防禦骰子總數(2)
End If
trgoi2.Enabled = False

End Sub


Private Sub trnextend_Timer()
'============以下是技能檢查及啟動
'  If turnatk = 2 And atkingck(81, 2) = 1 Then
'       atkingck(81, 1) = 3
'       技能.艾依查庫_不屈之心  '(階段3)
'  End If
'  If turnatk = 1 And atkingckai(27, 2) = 1 Then
'       atkingckai(27, 1) = 3
'       AI技能.艾依查庫_不屈之心  '(階段3)
'  End If
'=============
Select Case Val(擲骰表單溝通暫時變數(3))
   Case 1
      傷害執行_使用者 (Val(擲骰表單溝通暫時變數(2)))
   Case 2
      傷害執行_電腦 (Val(擲骰表單溝通暫時變數(2)))
End Select
'============以下是技能檢查及啟動
'  If turnatk = 2 And atkingck(43, 2) = 1 Then
'       atkingck(43, 1) = 4
'       技能.雪莉_VBE_異質者  '(階段4)
'  End If
'  If turnatk = 2 And atkingck(14, 2) = 1 Then
'       atkingck(14, 1) = 3
'       技能.蕾_安魂曲_死神的鎮魂歌  '(階段3)
'  End If
'  If turnatk = 2 And atkingck(62, 2) = 1 Then
'       atkingck(62, 1) = 3
'       技能.蕾_EX_安魂曲_死神的鎮魂歌  '(階段3)
'  End If
'  If turnatk = 1 And atkingckai(126, 2) = 1 Then
'       atkingckai(126, 1) = 3
'       AI技能.蕾_安魂曲_死神的鎮魂歌  '(階段3)
'  End If
'  If turnatk = 1 And atkingckai(63, 2) = 1 Then
'       atkingckai(63, 1) = 3
'       AI技能.蕾_EX_安魂曲_死神的鎮魂歌  '(階段3)
'  End If
'=============
目前數(24) = 21
等待時間_2.Enabled = True
trnextend.Enabled = False
End Sub

Private Sub trtimeline_Timer()
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
'            時間軸顏色變化紀錄暫時變數(1, 1) = 3
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
'       時間軸顏色變化紀錄暫時變數(4, 2) = 時間軸顏色變化紀錄暫時變數(4, 2) + 1
       Select Case 時間軸顏色變化紀錄暫時變數(4, 1)
           Case 1
'               If 時間軸顏色變化紀錄暫時變數(4, 2) >= 時間軸顏色變化紀錄暫時變數(1, 1) Then
'                    時間軸顏色變化紀錄暫時變數(4, 2) = 0
                    If 255 - Val(時間軸顏色變化紀錄暫時變數(4, 3)) < 9 Then
                       timelinein1.BorderColor = RGB(255, 0, 0)
                       timelinein2.BorderColor = RGB(255, 0, 0)
                       時間軸顏色變化紀錄暫時變數(4, 3) = 255
                    Else
                       timelinein1.BorderColor = RGB(Val(時間軸顏色變化紀錄暫時變數(4, 3)) + 9, 0, 0)
                       timelinein2.BorderColor = RGB(Val(時間軸顏色變化紀錄暫時變數(4, 3)) + 9, 0, 0)
                       時間軸顏色變化紀錄暫時變數(4, 3) = Val(時間軸顏色變化紀錄暫時變數(4, 3)) + 9
                    End If
'                End If
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
'            Case 3
'               時間軸顏色變化紀錄暫時變數(4, 1) = 1
'            Case 4
'               時間軸顏色變化紀錄暫時變數(4, 1) = 2
       End Select
End Select
If timelineout1.X1 >= timelineout1.X2 Then
    戰鬥系統類.時間軸_停止
    turnpageonin = 0
    bnok.Picture = LoadPicture(app_path & "gif\system\ok_3.jpg")
    目前數(24) = 4
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
Dim m As Integer '暫時變數
Do
    Randomize
    m = Int(Rnd() * Val(公用牌各牌類型紀錄數(0, 2))) + 1
    '===========
    If BattleCardNum <= 0 Then
        Exit Do
    End If
    If pagecardnum(m, 6) = 4 Then
       tr牌組_抽牌_使用者.Enabled = False
       戰鬥系統類.getpage 1, m
       Exit Do
    End If
Loop
End Sub

Private Sub tr牌組_抽牌_電腦_Timer()
Dim m As Integer '暫時變數
Do
    Randomize
    m = Int(Rnd() * Val(公用牌各牌類型紀錄數(0, 2))) + 1
    '===========
    If BattleCardNum <= 0 Then
        Exit Do
    End If
    If pagecardnum(m, 6) = 4 Then
       tr牌組_抽牌_電腦.Enabled = False
       戰鬥系統類.getpage 2, m
       Exit Do
    End If
Loop
End Sub

Private Sub tr牌組_翻牌_Timer()
'=======固定牌位置
'card(目前數(16)).Left = 240
'card(目前數(16)).Top = 960
''===========
'戰鬥系統類.執行動作_翻牌 目前數(16)
'tr牌組_翻牌.Enabled = False
    '==============以下是技能檢查及啟動
'    If atkingck(54, 2) = 1 Then
'      atkingck(54, 1) = 5
'      技能.羅莎琳_黑霧幻影 '(階段5)
'   End If
'   If atkingck(55, 2) = 1 Then
'      atkingck(55, 1) = 5
'      技能.羅莎琳_EX_黑霧幻影 '(階段5)
'   End If
   '=======================
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
    '==============以下是技能檢查及啟動
'    If atkingck(61, 2) = 1 Then
'      atkingck(61, 1) = 4
'      技能.古魯瓦爾多_精神力吸收 '(階段4)
'    End If
'    If atkingck(37, 2) = 1 Then
'      atkingck(37, 1) = 4
'      技能.艾茵_一顆心 '(階段4)
'    End If
'   If atkingck(56, 2) = 1 Then
'      atkingck(56, 1) = 4
'      技能.伊芙琳_怠惰的墓表 '(階段4)
'   End If
'   If atkingck(59, 2) = 1 Then
'      atkingck(59, 1) = 5
'      技能.伊芙琳_赤紅石榴 '(階段5)
'   End If
'   If atkingck(72, 2) = 1 Then
'      atkingck(72, 1) = 6
'      技能.艾伯李斯特_雷擊 '(階段6)
'   End If
'   If atkingck(122, 2) = 1 Then
'      atkingck(122, 1) = 4
'      技能.瑪格莉特_月光 '(階段4)
'   End If
'   If atkingck(129, 2) = 1 Then
'      atkingck(129, 1) = 6
'      技能.庫勒尼西_瘋狂眼窩 '(階段6)
'   End If
'   If atkingck(144, 2) = 1 Then
'      atkingck(144, 1) = 4
'      技能.傑多_因果之線 '(階段4)
'   End If
'   If atkingck(156, 2) = 1 Then
'      atkingck(156, 1) = 4
'      技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段4)
'   End If
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
Private Sub uspi1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cardus(Index).Left = uspiin(Index).Left
cardus(Index).Top = 5760
cardus(Index).ZOrder
If 人物卡面背面編號紀錄數(1) = 1 And 人物卡面背面編號紀錄數(2) = Index Then
'    PEAFcardback(1).Visible = True
    cardus(Index).Visible = True
'    PEAFcardback(1).ZOrder
Else
    cardus(Index).Visible = True
'    PEAFcardback(1).Visible = False
End If
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

atkinghelpc.Visible = False
End Sub


Private Sub uspi4_Change(Index As Integer)
If Val(uspi4(Index).Caption) = Val(liveusmax(Index)) Then
'   usbi1(Index).ForeColor = RGB(255, 255, 255)
   uspi4(Index).ForeColor = RGB(255, 255, 255)
'   cardbackus(Index).Visible = False
End If
 If Val(uspi4(Index).Caption) < Val(liveusmax(Index)) Then
'   usbi1(Index).ForeColor = RGB(255, 255, 128)
   uspi4(Index).ForeColor = RGB(255, 255, 128)
'   cardbackus(Index).Visible = False
 End If
 If Val(uspi4(Index).Caption) <= Val(liveus41(Index)) Then
'   usbi1(Index).ForeColor = RGB(255, 0, 0)
   uspi4(Index).ForeColor = RGB(255, 0, 0)
'   cardbackus(Index).Visible = False
 End If
If Val(uspi4(Index).Caption) <= 0 Then
'    usbi1(Index).Caption = 0
    uspi4(Index).Caption = 0
'    cardbackus(Index).Visible = True
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
'Select Case 角色人物對戰人數(1, 2)
'    Case 1
'        pnm = Val(formsettingpersonus.personfleftall.Text)
'    Case 2
'        pnm = Val(formsettingpersonus2.personfleftall.Text)
'    Case 3
'        pnm = Val(formsettingpersonus3.personfleftall.Text)
'End Select
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
'Select Case 角色人物對戰人數(2, 2)
'    Case 1
'        pnm = Val(formsettingpersoncom.personfleftall.Text)
'    Case 2
'        pnm = Val(formsettingpersoncom2.personfleftall.Text)
'    Case 3
'        pnm = Val(formsettingpersoncom3.personfleftall.Text)
'End Select
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
   目前數(24) = 1
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
'If 智慧型AI系統_目前可執行之人物判斷(namecom(角色人物對戰人數(2, 2))) = True Then
    Dim wtyr As Integer '暫時變數
    If moveturn = 2 Then wtyr = 1 Else wtyr = 0
    智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 2, 2, namecom(角色人物對戰人數(2, 2)), movecp, wtyr
    GoTo 智慧型AI出牌_執行階段結束
'End If
'============以下是技能檢查及啟動
'If turnatk = 1 And moveturn = 1 And 異常狀態檢查數(18, 2) = 0 Then
'    AI技能.雪莉_異質者 '(階段1)
'    AI技能.多妮妲_異質者  '(階段1)
'End If
''===================(技能-雪莉/多妮妲-異質者-AI 先攻時檢查)
'If turnatk = 1 And atkingckai(12, 2) = 1 And moveturn = 1 Then
'    GoTo AI技能_雪莉_多妮妲_異質者_執行階段2
'End If
'===================
'----------以下為電腦判斷出牌程式碼（防禦方）
For j = 1 To 公用牌實體卡片分隔紀錄數(1)
     If pagecardnum(j, 1) = a2a And Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
       pagecardnum(j, 11) = 1
     ElseIf pagecardnum(j, 3) = a2a And Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
       cspce = pagecardnum(j, 1)
       cspme = pagecardnum(j, 2)
       pagecardnum(j, 1) = pagecardnum(j, 3)
       pagecardnum(j, 2) = pagecardnum(j, 4)
       pagecardnum(j, 3) = cspce
       pagecardnum(j, 4) = cspme
       If pageonin(j) = 2 Then
          pageonin(j) = 1
       Else
          pageonin(j) = 2
       End If
       pagecardnum(j, 11) = 1
     End If
Next

'============以下是技能檢查及啟動
'If turnatk = 1 And moveturn = 2 And 異常狀態檢查數(18, 2) = 0 Then
'    AI技能.雪莉_異質者 '(階段1)
'    AI技能.多妮妲_異質者  '(階段1)
'End If
'==============
'AI人物.艾伯李斯特 1
'AI人物.梅倫 1
'AI人物.夏洛特 1
'AI人物.艾蕾可 1
''==============
'If moveturn = 2 Then
'   AI人物.全人物通用 2
'End If
'===============
'AI技能_雪莉_多妮妲_異質者_執行階段2:
智慧型AI出牌_執行階段結束:
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
'===========以下是技能檢查及啟動(AI技能-C.C.-原子之心)
'If turnatk = 1 And atkingckai(57, 2) = 1 Then
'    atkingckai(57, 1) = 2
'    AI技能.CC_原子之心  '(階段2)
'End If
''===============以下是技能檢查及啟動(AI-傑多-因果之輪)
'If turnatk = 1 And atkingckai(120, 2) = 1 And atkingckai(120, 1) = 1 Then
'   atkingckai(120, 1) = 2
'   AI技能.傑多_因果之輪 '(階段2)
'   Exit Sub
'ElseIf turnatk = 1 And atkingckai(120, 2) = 1 And atkingckai(120, 1) = 4 Then
'   AI技能.傑多_因果之輪 '(階段4)
'End If
''===============以下是技能檢查及啟動(AI-克頓-竊取資料)
'If turnatk = 1 And atkingckai(131, 2) = 1 And atkingckai(131, 1) = 1 Then
'   atkingckai(131, 1) = 2
'   AI技能.克頓_竊取資料 '(階段2)
'   Exit Sub
'ElseIf turnatk = 1 And atkingckai(131, 2) = 1 And atkingckai(131, 1) = 4 Then
'   AI技能.克頓_竊取資料 '(階段4)
'End If
''===============以下是技能檢查及啟動(梅莉-夢幻魔杖)
'If turnatk = 1 And atkingck(106, 2) = 1 And atkingck(106, 1) = 1 Then
'   atkingck(106, 1) = 2
'   技能.梅莉_夢幻魔杖 '(階段2)
'   Exit Sub
'ElseIf turnatk = 1 And atkingck(106, 2) = 1 And atkingck(106, 1) = 4 Then
'   技能.梅莉_夢幻魔杖 '(階段4)
'End If
''===============以下是技能檢查及啟動(AI-梅莉-徬徨夢羽)
'If turnatk = 1 And atkingckai(100, 2) = 1 And atkingckai(100, 1) = 1 Then
'   atkingckai(100, 1) = 2
'   AI技能.梅莉_徬徨夢羽 '(階段2)
'   Exit Sub
'ElseIf turnatk = 1 And atkingckai(100, 2) = 1 And atkingckai(100, 1) = 4 Then
'   AI技能.梅莉_徬徨夢羽 '(階段4)
'End If
''=====================
'技能動畫顯示階段數 = 1
'戰鬥系統類.技能啟動數量檢查
'    '=================以下是技能檢查及啟動(AI技能-C.C.-原子之心)
'    If turnatk = 1 And atkingckai(57, 2) = 1 Then
'        atkingckai(57, 1) = 3
'        AI技能.CC_原子之心  '(階段3)
''        GoTo AI技能_CC_原子之心_跳入點
'    End If
'   '============以下是技能檢查及啟動
'    If turnatk = 1 And atkingckai(28, 2) = 1 Then
'       atkingckai(28, 1) = 2
'       AI技能.音音夢_溫柔注射  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(58, 2) = 1 Then
'       atkingckai(58, 1) = 2
'       AI技能.蕾_EX_協奏曲_加百烈的守護  '(階段2)
'    End If
'   '==========================
'    If turnatk = 1 And atkingck(1, 2) = 1 Then
'       atkingck(1, 1) = 3
'       技能.雪莉_自殺傾向 Index  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(42, 2) = 1 Then
'       atkingck(42, 1) = 3
'       技能.雪莉_VBE_自殺傾向 Index  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(3, 2) = 1 Then
'       atkingck(3, 1) = 3
'       技能.雪莉_飛刃雨 '(階段3)
'    End If
'    If turnatk = 1 And atkingck(45, 2) = 1 Then
'       atkingck(45, 1) = 3
'       技能.雪莉_VBE_飛刃雨  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(6, 2) = 1 Then
'       atkingck(6, 1) = 2
'       技能.古魯瓦爾多_猛擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(9, 2) = 1 Then
'       atkingck(9, 1) = 4
'       技能.帕茉_慈悲的藍眼 '(階段4)
'    End If
'    If turnatk = 1 And atkingck(18, 2) = 1 Then
'       atkingck(18, 1) = 2
'       技能.帕茉_戰慄的狼牙 '(階段2)
'    End If
'    If turnatk = 1 And atkingck(17, 2) = 1 Then
'       atkingck(17, 1) = 2
'       技能.帕茉_靜謐之背 '(階段2)
'    End If
'    If turnatk = 1 And atkingck(15, 2) = 1 Then
'       atkingck(15, 1) = 2
'       技能.蕾_終曲_無盡輪迴的終結  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(23, 2) = 1 Then
'       atkingck(23, 1) = 2
'       技能.史塔夏_愚者之手  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(25, 2) = 1 Then
'       atkingck(25, 1) = 2
'       技能.史塔夏_命運的鐵門  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(35, 2) = 1 Then
'       atkingck(35, 1) = 2
'       技能.CC_高頻電磁手術刀  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(51, 2) = 1 Then
'       atkingck(51, 1) = 2
'       技能.羅莎琳_染血之刃  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(50, 2) = 1 Then
'       atkingck(50, 1) = 2
'       技能.羅莎琳_EX_染血之刃  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(52, 2) = 1 Then
'       atkingck(52, 1) = 2
'       技能.羅莎琳_黑霧的纏繞  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(58, 2) = 1 Then
'       atkingck(58, 1) = 2
'       技能.伊芙琳_紅蓮車輪  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(65, 2) = 1 Then
'       atkingck(65, 1) = 2
'       技能.梅倫_Lowball  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(66, 2) = 1 Then
'       atkingck(66, 1) = 2
'       技能.梅倫_Gamble  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(69, 2) = 1 Then
'       atkingck(69, 1) = 3
'       技能.音音夢_愉快抽血 Index  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(71, 2) = 1 Then
'       atkingck(71, 1) = 2
'       技能.艾伯李斯特_精密射擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(72, 2) = 1 Then
'       atkingck(72, 1) = 2
'       技能.艾伯李斯特_雷擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(78, 2) = 1 Then
'       atkingck(78, 1) = 2
'       技能.艾依查庫_連射  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(79, 2) = 1 Then
'       atkingck(79, 1) = 4
'       技能.艾依查庫_神速之劍 (0) '(階段4)
'    End If
'    If turnatk = 1 And atkingck(80, 2) = 1 Then
'       atkingck(80, 1) = 2
'       技能.艾依查庫_憤怒一擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(84, 2) = 1 Then
'       atkingck(84, 1) = 2
'       技能.布勞_時間爆彈  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(86, 2) = 1 Then
'       atkingck(86, 1) = 2
'       技能.阿貝爾_霸王閃擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(87, 2) = 1 Then
'       atkingck(87, 1) = 2
'       技能.阿貝爾_閃電旋風刺  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(88, 2) = 1 Then
'       atkingck(88, 1) = 2
'       技能.阿貝爾_幻影劍舞  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(90, 2) = 1 Then
'       atkingck(90, 1) = 2
'       技能.利恩_劫影攻擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(91, 2) = 1 Then
'       atkingck(91, 1) = 2
'       技能.利恩_毒牙  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(93, 2) = 1 Then
'       atkingck(93, 1) = 2
'       技能.利恩_背刺  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(95, 2) = 1 Then
'       atkingck(95, 1) = 2
'       技能.夏洛特_冬之夢  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(116, 2) = 1 Then
'       atkingck(116, 1) = 2
'       技能.泰瑞爾_Rud_913  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(118, 2) = 1 Then
'       atkingck(118, 1) = 2
'       技能.泰瑞爾_Chr_799  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(119, 2) = 1 Then
'       atkingck(119, 1) = 2
'       技能.泰瑞爾_Wil_846  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(122, 2) = 1 Then
'       atkingck(122, 1) = 2
'       技能.瑪格莉特_月光  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(125, 2) = 1 Then
'       atkingck(125, 1) = 2
'       技能.瑪格莉特_地獄獵心獸  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(130, 2) = 1 Then
'       atkingck(130, 1) = 2
'       技能.庫勒尼西_深淵  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(135, 2) = 1 Then
'       atkingck(135, 1) = 2
'       技能.蕾格烈芙_CTL  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(136, 2) = 1 Then
'       atkingck(136, 1) = 2
'       技能.蕾格烈芙_BPA  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(140, 2) = 1 Then
'       atkingck(140, 1) = 2
'       技能.多妮妲_殘虐傾向  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(143, 2) = 1 Then
'       atkingck(143, 1) = 2
'       技能.多妮妲_律死擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(150, 2) = 1 Then
'       atkingck(150, 1) = 2
'       技能.阿奇波爾多_致命槍擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(151, 2) = 1 Then
'       atkingck(151, 1) = 2
'       技能.阿奇波爾多_劫影攻擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(155, 2) = 1 Then
'       atkingck(155, 1) = 2
'       技能.洛洛妮_砲擊壓制  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(156, 2) = 1 Then
'       atkingck(156, 1) = 2
'       技能.洛洛妮_貪婪之刃與嗜血之槍  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(159, 2) = 1 Then
'       atkingck(159, 1) = 2
'       技能.克頓_隱蔽射擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(98, 2) = 1 Then
'       atkingck(98, 1) = 2
'       技能.露緹亞_腐朽之靈  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(101, 2) = 1 Then
'       atkingck(101, 1) = 4
'       技能.露緹亞_渦騎劍閃 (0)  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(104, 2) = 1 Then
'       atkingck(104, 1) = 2
'       技能.艾蕾可_聖王威光  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(109, 2) = 1 Then
'       atkingck(109, 1) = 2
'       技能.梅莉_夢境搖籃  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(112, 2) = 1 Then
'       atkingck(112, 1) = 2
'       技能.貝琳達_裂地冰牙  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(113, 2) = 1 Then
'       atkingck(113, 1) = 2
'       技能.貝琳達_溶魂之雨  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(161, 2) = 1 Then
'       atkingck(161, 1) = 2
'       技能.蕾_EX_終曲_無盡輪迴的終結  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(46, 2) = 1 Then
'       atkingck(46, 1) = 2
'       技能.尤莉卡_奸佞的鐵鎚  '(階段2)
'    End If
'    '=================================================
'    If turnatk = 1 And atkingckai(9, 2) = 1 Then
'       atkingckai(9, 1) = 2
'       AI技能.妖精王妃_煉獄之翼  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(18, 2) = 1 Then
'       atkingckai(18, 1) = 2
'       AI技能.吸血姬蕾米雅_消失  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(27, 2) = 1 Then
'       atkingckai(27, 1) = 2
'       AI技能.艾依查庫_不屈之心  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(31, 2) = 1 Then
'       atkingckai(31, 1) = 2
'       AI技能.梅倫_Jackpot  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(38, 2) = 1 Then
'       atkingckai(38, 1) = 2
'       AI技能.艾茵_兩個身體  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(42, 2) = 1 Then
'       atkingckai(42, 1) = 2
'       AI技能.瑪格莉特_恍惚  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(44, 2) = 1 Then
'       atkingckai(44, 1) = 2
'       AI技能.庫勒尼西_沙漠中的海市蜃樓  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(46, 2) = 1 Then
'       atkingckai(46, 1) = 2
'       AI技能.庫勒尼西_黑暗漩渦  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(47, 2) = 1 Then
'       atkingckai(47, 1) = 2
'       AI技能.蕾格烈芙_LAR  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(54, 2) = 1 Then
'       atkingckai(54, 1) = 2
'       AI技能.蕾_協奏曲_加百烈的守護  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(55, 2) = 1 Then
'       atkingckai(55, 1) = 2
'       AI技能.史塔夏_時間種子  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(56, 2) = 1 Then
'       atkingckai(56, 1) = 2
'       AI技能.艾茵_九個靈魂  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(49, 2) = 1 Then
'       atkingckai(49, 1) = 2
'       AI技能.阿奇波爾多_防護射擊  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(60, 2) = 1 Then
'       atkingckai(60, 1) = 2
'       AI技能.羅莎琳_咀咒的刻印  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(61, 2) = 1 Then
'       atkingckai(61, 1) = 2
'       AI技能.伊芙琳_慟哭之歌  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(62, 2) = 1 Then
'       atkingckai(62, 1) = 2
'       AI技能.古魯瓦爾多_血之恩賜  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(63, 2) = 1 Then
'       atkingckai(63, 1) = 2
'       AI技能.蕾_EX_安魂曲_死神的鎮魂歌  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(64, 2) = 1 Then
'       atkingckai(64, 1) = 2
'       AI技能.梅倫_High_hand  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(67, 2) = 1 Then
'       atkingckai(67, 1) = 2
'       AI技能.艾伯李斯特_茨林  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(70, 2) = 1 Then
'       atkingckai(70, 1) = 2
'       AI技能.布勞_時間追獵  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(74, 2) = 1 Then
'       atkingckai(74, 1) = 2
'       AI技能.利恩_反擊的狼煙  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(76, 2) = 1 Then
'       atkingckai(76, 1) = 2
'       AI技能.泰瑞爾_Von_541  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(79, 2) = 1 Then
'       atkingckai(79, 1) = 2
'       AI技能.庫勒尼西_瘋狂眼窩  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(85, 2) = 1 Then
'       atkingckai(85, 1) = 2
'       AI技能.洛洛妮_風暴感知  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(91, 2) = 1 Then
'       atkingckai(91, 1) = 3
'       AI技能.艾蕾可_王座之炎  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(92, 2) = 1 Then
'       atkingckai(92, 1) = 2
'       AI技能.艾蕾可_白百合  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(96, 2) = 1 Then
'       atkingckai(96, 1) = 2
'       AI技能.露緹亞_朦朧之暗  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(112, 2) = 1 Then
'       atkingckai(112, 1) = 2
'       AI技能.音音夢_秘密苦藥  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(121, 2) = 1 Then
'       atkingckai(121, 1) = 2
'       AI技能.傑多_因果之刻  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(123, 2) = 1 Then
'       atkingckai(123, 1) = 2
'       AI技能.貝琳達_水晶幻鏡  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(126, 2) = 1 Then
'       atkingckai(126, 1) = 2
'       AI技能.蕾_安魂曲_死神的鎮魂歌  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(128, 2) = 1 Then
'       atkingckai(128, 1) = 2
'       AI技能.羅莎琳_黑霧幻影  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(132, 2) = 1 Then
'       atkingckai(132, 1) = 2
'       AI技能.克頓_逃亡計畫  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(137, 2) = 1 Then
'       atkingckai(137, 1) = 2
'       AI技能.尤莉卡_不善的信仰  '(階段2)
'    End If
'    If turnatk = 1 And atkingckai(138, 2) = 1 Then
'       atkingckai(138, 1) = 2
'       AI技能.尤莉卡_曲惡的安寧  '(階段2)
'    End If
'    '==============(相同骰子類)
'    If turnatk = 1 And atkingckai(15, 2) = 1 Then
'        atkingckai(15, 1) = 2
'       AI技能.羊角獸2012_致命格擋  '(階段2)
'    End If
'    '==============(減低對手防禦類)
'    If turnatk = 1 And atkingck(13, 2) = 1 Then
'       atkingck(13, 1) = 3
'       技能.蕾_輪旋曲_琉璃色的微風  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(19, 2) = 1 Then
'       atkingck(19, 1) = 3
'       技能.蕾_EX_輪旋曲_琉璃色的微風  '(階段3)
'    End If
'    '===============================================
'    If turnatk = 1 And atkingck(16, 2) = 1 Then
'       atkingck(16, 1) = 4
'       技能.艾茵_十三隻眼  '(階段4)
'    End If
'     '==================
'    If turnatk = 1 And atkingckai(90, 2) = 1 Then
'       atkingckai(90, 1) = 2
'       AI技能.夏洛特_大聖堂  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(147, 2) = 1 Then
'       atkingck(147, 1) = 2
'       技能.傑多_因果之幻  '(階段2)
'    End If
'==================
'AI技能_CC_原子之心_跳入點:      'AI技能-C.C.-原子之心-程式跳入點
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
'========================================
'  For p = 1 To 攻擊防禦骰子總數(1)
'     Randomize
'     i = Int(Rnd() * 6) + 1
'     If i = 1 Or i = 6 Then 擲骰表單溝通暫時變數(7) = Val(擲骰表單溝通暫時變數(7)) + 1
'  Next
'  For p = 1 To 攻擊防禦骰子總數(2)
'     Randomize
'     j = Int(Rnd() * 6) + 1
'     If j = 1 Or j = 6 Then 擲骰表單溝通暫時變數(8) = Val(擲骰表單溝通暫時變數(8)) + 1
'  Next
  '=============================
'    If turnatk = 1 And atkingckai(12, 2) = 1 Then
'        atkingckai(12, 1) = 3
'        AI技能.雪莉_異質者  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(82, 2) = 1 Then
'        atkingckai(82, 1) = 3
'        AI技能.多妮妲_異質者  '(階段3)
'    End If
    '===================
    '===========================執行階段插入點(ATK-12/DEF-32)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 12, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 32, 2
    '============================
      階段狀態數 = 2
'      atkingtrtot.Interval = 600
'      atkingtrtot.Enabled = True
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
''============以下是技能檢查及啟動
'    If turnatk = 1 And atkingck(13, 2) = 1 Then
'       atkingck(13, 1) = 2
'       技能.蕾_輪旋曲_琉璃色的微風  '(階段2)
'    End If
'    If turnatk = 1 And atkingck(19, 2) = 1 Then
'       atkingck(19, 1) = 2
'       技能.蕾_EX_輪旋曲_琉璃色的微風  '(階段2)
'    End If
'    If atkingck(16, 2) = 1 And turnatk = 1 Then
'        atkingck(16, 1) = 5
'        技能.艾茵_十三隻眼 '(階段5)
'        trgoi2_Timer
'    End If
'=====================
'=====================
'--------以下為防回牌程式碼
'cn22.Visible = False
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
'===========以下是技能檢查及啟動(技能-C.C.-原子之心)
'If turnatk = 2 And atkingck(36, 2) = 1 Then
'    atkingck(36, 1) = 2
'    技能.CC_原子之心  '(階段2)
'End If
''===============以下是技能檢查及啟動(傑多-因果之輪)
'If turnatk = 2 And atkingck(145, 2) = 1 And atkingck(145, 1) = 1 Then
'   atkingck(145, 1) = 2
'   技能.傑多_因果之輪 '(階段2)
'   Exit Sub
'ElseIf turnatk = 2 And atkingck(145, 2) = 1 And atkingck(145, 1) = 4 Then
'   技能.傑多_因果之輪 '(階段4)
'End If
''===============以下是技能檢查及啟動(克頓-竊取資料)
'If turnatk = 2 And atkingck(157, 2) = 1 And atkingck(157, 1) = 1 Then
'   atkingck(157, 1) = 2
'   技能.克頓_竊取資料 '(階段2)
'   Exit Sub
'ElseIf turnatk = 2 And atkingck(157, 2) = 1 And atkingck(157, 1) = 4 Then
'   技能.克頓_竊取資料 '(階段4)
'End If
''===============以下是技能檢查及啟動(AI-梅莉-夢幻魔杖)
'If turnatk = 2 And atkingckai(99, 2) = 1 And atkingckai(99, 1) = 1 Then
'   atkingckai(99, 1) = 2
'   AI技能.梅莉_夢幻魔杖 '(階段2)
'   Exit Sub
'ElseIf turnatk = 2 And atkingckai(99, 2) = 1 And atkingckai(99, 1) = 4 Then
'   AI技能.梅莉_夢幻魔杖 '(階段4)
'End If
''===============以下是技能檢查及啟動(梅莉-徬徨夢羽)
'If turnatk = 2 And atkingck(107, 2) = 1 And atkingck(107, 1) = 1 Then
'   atkingck(107, 1) = 2
'   技能.梅莉_徬徨夢羽 '(階段2)
'   Exit Sub
'ElseIf turnatk = 2 And atkingck(107, 2) = 1 And atkingck(107, 1) = 4 Then
'   技能.梅莉_徬徨夢羽 '(階段4)
'End If
''========================
'技能動畫顯示階段數 = 1
'戰鬥系統類.技能啟動數量檢查
'    '=================以下是技能檢查及啟動(技能-C.C.-原子之心)
'    If turnatk = 2 And atkingck(36, 2) = 1 Then
'        atkingck(36, 1) = 3
'        技能.CC_原子之心  '(階段3)
''        GoTo 技能_CC_原子之心_跳入點
'    End If
'   '============以下是技能檢查及啟動
'   If turnatk = 2 And atkingck(38, 2) = 1 Then
'       atkingck(38, 1) = 2
'       技能.蕾_EX_協奏曲_加百烈的守護  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(68, 2) = 1 Then
'       atkingck(68, 1) = 2
'       技能.音音夢_溫柔注射  '(階段2)
'    End If
'    '----------------------
'    If turnatk = 2 And atkingckai(1, 2) = 1 Then
'       atkingckai(1, 1) = 4
'       AI技能.雪莉_自殺傾向 (0)  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(5, 2) = 1 Then
'       atkingckai(5, 1) = 5
'       AI技能.雪莉_飛刃雨   '(階段5)
'    End If
'    If turnatk = 2 And atkingckai(3, 2) = 1 Then
'       atkingckai(3, 1) = 2
'       AI技能.古魯瓦爾多_猛擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(11, 2) = 1 Then
'       atkingckai(11, 1) = 3
'       AI技能.蕾_終曲_無盡輪迴的終結   '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(7, 2) = 1 Then
'       atkingckai(7, 1) = 2
'       AI技能.南瓜王_重壓  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(8, 2) = 1 Then
'       atkingckai(8, 1) = 2
'       AI技能.妖精王妃_冰結之翼  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(16, 2) = 1 Then
'       atkingckai(16, 1) = 2
'       AI技能.吸血姬蕾米雅_吸血   '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(17, 2) = 1 Then
'       atkingckai(17, 1) = 2
'       AI技能.吸血姬蕾米雅_高貴的晚餐   '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(19, 2) = 1 Then
'       atkingckai(19, 1) = 2
'       AI技能.艾伯李斯特_精密射擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(20, 2) = 1 Then
'       atkingckai(20, 1) = 2
'       AI技能.史塔夏_愚者之手  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(21, 2) = 1 Then
'       atkingckai(21, 1) = 2
'       AI技能.史塔夏_命運的鐵門  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(22, 2) = 1 Then
'       atkingckai(22, 1) = 2
'       AI技能.阿貝爾_霸王閃擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(23, 2) = 1 Then
'       atkingckai(23, 1) = 2
'       AI技能.阿貝爾_幻影劍舞  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(24, 2) = 1 Then
'       atkingckai(24, 1) = 2
'       AI技能.布勞_時間爆彈  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(25, 2) = 1 Then
'       atkingckai(25, 1) = 2
'       AI技能.艾依查庫_連射  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(26, 2) = 1 Then
'       atkingckai(26, 1) = 4
'       AI技能.艾依查庫_神速之劍 (0) '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(30, 2) = 1 Then
'       atkingckai(30, 1) = 2
'       AI技能.梅倫_Gamble  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(32, 2) = 1 Then
'       atkingckai(32, 1) = 2
'       AI技能.羅莎琳_染血之刃  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(140, 2) = 1 Then
'       atkingckai(140, 1) = 2
'       AI技能.羅莎琳_EX_染血之刃  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(34, 2) = 1 Then
'       atkingckai(34, 1) = 2
'       AI技能.帕茉_戰慄的狼牙  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(36, 2) = 1 Then
'       atkingckai(36, 1) = 2
'       AI技能.帕茉_靜謐之背  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(35, 2) = 1 Then
'       atkingckai(35, 1) = 4
'       AI技能.帕茉_慈悲的藍眼   '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(39, 2) = 1 Then
'       atkingckai(39, 1) = 2
'       AI技能.夏洛特_冬之夢  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(40, 2) = 1 Then
'       atkingckai(40, 1) = 2
'       AI技能.泰瑞爾_Rud_913  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(77, 2) = 1 Then
'       atkingckai(77, 1) = 2
'       AI技能.泰瑞爾_Chr_799  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(41, 2) = 1 Then
'       atkingckai(41, 1) = 2
'       AI技能.泰瑞爾_Wil_846  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(43, 2) = 1 Then
'       atkingckai(43, 1) = 2
'       AI技能.瑪格莉特_地獄獵心獸  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(50, 2) = 1 Then
'       atkingckai(50, 1) = 2
'       AI技能.CC_高頻電磁手術刀  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(51, 2) = 1 Then
'       atkingckai(51, 1) = 2
'       AI技能.伊芙琳_紅蓮車輪  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(53, 2) = 1 Then
'       atkingckai(53, 1) = 2
'       AI技能.多妮妲_殘虐傾向  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(52, 2) = 1 Then
'       atkingckai(52, 1) = 2
'       AI技能.多妮妲_律死擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(45, 2) = 1 Then
'       atkingckai(45, 1) = 2
'       AI技能.庫勒尼西_深淵  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(59, 2) = 1 Then
'       atkingckai(59, 1) = 2
'       AI技能.羅莎琳_黑霧的纏繞  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(65, 2) = 1 Then
'       atkingckai(65, 1) = 2
'       AI技能.梅倫_Lowball  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(66, 2) = 1 Then
'       atkingckai(66, 1) = 2
'       AI技能.艾伯李斯特_雷擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(69, 2) = 1 Then
'       atkingckai(69, 1) = 2
'       AI技能.艾依查庫_憤怒一擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(71, 2) = 1 Then
'       atkingckai(71, 1) = 2
'       AI技能.阿貝爾_閃電旋風刺  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(72, 2) = 1 Then
'       atkingckai(72, 1) = 2
'       AI技能.利恩_劫影攻擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(73, 2) = 1 Then
'       atkingckai(73, 1) = 2
'       AI技能.利恩_毒牙  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(75, 2) = 1 Then
'       atkingckai(75, 1) = 2
'       AI技能.利恩_背刺  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(78, 2) = 1 Then
'       atkingckai(78, 1) = 2
'       AI技能.瑪格莉特_月光  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(80, 2) = 1 Then
'       atkingckai(80, 1) = 2
'       AI技能.蕾格烈芙_CTL  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(81, 2) = 1 Then
'       atkingckai(81, 1) = 2
'       AI技能.蕾格烈芙_BPA  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(83, 2) = 1 Then
'       atkingckai(83, 1) = 2
'       AI技能.阿奇波爾多_致命槍擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(84, 2) = 1 Then
'       atkingckai(84, 1) = 2
'       AI技能.阿奇波爾多_劫影攻擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(86, 2) = 1 Then
'       atkingckai(86, 1) = 2
'       AI技能.洛洛妮_砲擊壓制  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(87, 2) = 1 Then
'       atkingckai(87, 1) = 2
'       AI技能.洛洛妮_貪婪之刃與嗜血之槍  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(93, 2) = 1 Then
'       atkingckai(93, 1) = 2
'       AI技能.艾蕾可_聖王威光  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(95, 2) = 1 Then
'       atkingckai(95, 1) = 2
'       AI技能.露緹亞_腐朽之靈  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(98, 2) = 1 Then
'       atkingckai(98, 1) = 4
'       AI技能.露緹亞_渦騎劍閃 (0) '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(102, 2) = 1 Then
'       atkingckai(102, 1) = 2
'       AI技能.梅莉_夢境搖籃  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(111, 2) = 1 Then
'       atkingckai(111, 1) = 3
'       AI技能.音音夢_愉快抽血 (0) '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(124, 2) = 1 Then
'       atkingckai(124, 1) = 2
'       AI技能.貝琳達_裂地冰牙  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(125, 2) = 1 Then
'       atkingckai(125, 1) = 2
'       AI技能.貝琳達_溶魂之雨  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(127, 2) = 1 Then
'       atkingckai(127, 1) = 2
'       AI技能.蕾_EX_終曲_無盡輪迴的終結  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(133, 2) = 1 Then
'       atkingckai(133, 1) = 2
'       AI技能.克頓_隱蔽射擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(136, 2) = 1 Then
'       atkingckai(136, 1) = 2
'       AI技能.尤莉卡_奸佞的鐵鎚  '(階段2)
'    End If
'    '========================================
'    If turnatk = 2 And atkingck(32, 2) = 1 Then
'       atkingck(32, 1) = 2
'       技能.艾茵_兩個身體  '(階段2)
'    End If
'   If turnatk = 2 And atkingck(26, 2) = 1 Then
'       atkingck(26, 1) = 2
'       技能.艾茵_九個靈魂  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(11, 2) = 1 Then
'       atkingck(11, 1) = 2
'       技能.蕾_協奏曲_加百烈的守護  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(14, 2) = 1 Then
'       atkingck(14, 1) = 2
'       技能.蕾_安魂曲_死神的鎮魂歌  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(62, 2) = 1 Then
'       atkingck(62, 1) = 2
'       技能.蕾_EX_安魂曲_死神的鎮魂歌  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(24, 2) = 1 Then
'       atkingck(24, 1) = 2
'       技能.史塔夏_時間種子  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(54, 2) = 1 Then
'       atkingck(54, 1) = 2
'       技能.羅莎琳_黑霧幻影  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(55, 2) = 1 Then
'       atkingck(55, 1) = 2
'       技能.羅莎琳_EX_黑霧幻影  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(60, 2) = 1 Then
'       atkingck(60, 1) = 2
'       技能.古魯瓦爾多_血之恩賜  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(63, 2) = 1 Then
'       atkingck(63, 1) = 2
'       技能.梅倫_High_hand  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(64, 2) = 1 Then
'       atkingck(64, 1) = 2
'       技能.梅倫_Jackpot  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(70, 2) = 1 Then
'       atkingck(70, 1) = 2
'       技能.音音夢_秘密苦藥  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(73, 2) = 1 Then
'       atkingck(73, 1) = 2
'       技能.艾伯李斯特_茨林  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(81, 2) = 1 Then
'       atkingck(81, 1) = 2
'       技能.艾依查庫_不屈之心  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(92, 2) = 1 Then
'       atkingck(92, 1) = 2
'       技能.利恩_反擊的狼煙  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(117, 2) = 1 Then
'       atkingck(117, 1) = 2
'       技能.泰瑞爾_Von_541  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(123, 2) = 1 Then
'       atkingck(123, 1) = 2
'       技能.瑪格莉特_恍惚  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(128, 2) = 1 Then
'       atkingck(128, 1) = 2
'       技能.庫勒尼西_沙漠中的海市蜃樓  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(129, 2) = 1 Then
'       atkingck(129, 1) = 2
'       技能.庫勒尼西_瘋狂眼窩  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(131, 2) = 1 Then
'       atkingck(131, 1) = 2
'       技能.庫勒尼西_黑暗漩渦  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(137, 2) = 1 Then
'       atkingck(137, 1) = 2
'       技能.蕾格烈芙_LAR  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(146, 2) = 1 Then
'       atkingck(146, 1) = 2
'       技能.傑多_因果之刻  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(152, 2) = 1 Then
'       atkingck(152, 1) = 2
'       技能.阿奇波爾多_防護射擊  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(154, 2) = 1 Then
'       atkingck(154, 1) = 2
'       技能.洛洛妮_風暴感知   '(階段2)
'    End If
'    If turnatk = 2 And atkingck(158, 2) = 1 Then
'       atkingck(158, 1) = 2
'       技能.克頓_逃亡計畫   '(階段2)
'    End If
'    If turnatk = 2 And atkingck(99, 2) = 1 Then
'       atkingck(99, 1) = 2
'       技能.露緹亞_朦朧之暗  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(100, 2) = 1 Then
'       atkingck(100, 1) = 2
'       技能.露緹亞_暗影之翼  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(102, 2) = 1 Then
'       atkingck(102, 1) = 3
'       技能.艾蕾可_王座之炎  '(階段3)
'    End If
'    If turnatk = 2 And atkingck(103, 2) = 1 Then
'       atkingck(103, 1) = 2
'       技能.艾蕾可_白百合  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(111, 2) = 1 Then
'       atkingck(111, 1) = 2
'       技能.貝琳達_水晶幻鏡  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(47, 2) = 1 Then
'       atkingck(47, 1) = 2
'       技能.尤莉卡_不善的信仰  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(48, 2) = 1 Then
'       atkingck(48, 1) = 2
'       技能.尤莉卡_曲惡的安寧  '(階段2)
'    End If
'    '====================
'    If turnatk = 2 And atkingckai(14, 2) = 1 Then
'       atkingckai(14, 1) = 2
'       AI技能.羊角獸2012_致命衝撞   '(階段2)
'    End If
'    '====================
'    If turnatk = 2 And atkingckai(4, 2) = 1 Then
'       atkingckai(4, 1) = 2
'       AI技能.蕾_輪旋曲_琉璃色的微風  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(13, 2) = 1 Then
'       atkingckai(13, 1) = 2
'       AI技能.蕾_EX_輪旋曲_琉璃色的微風  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(53, 2) = 1 Then
'       atkingck(53, 1) = 2
'       技能.羅莎琳_咀咒的刻印  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(57, 2) = 1 Then
'       atkingck(57, 1) = 2
'       技能.伊芙琳_慟哭之歌  '(階段2)
'    End If
'    If turnatk = 2 And atkingck(83, 2) = 1 Then
'       atkingck(83, 1) = 2
'       技能.布勞_時間追獵  '(階段2)
'    End If
'    '==============================================
'    If turnatk = 2 And atkingckai(37, 2) = 1 Then
'       atkingckai(37, 1) = 4
'       AI技能.艾茵_十三隻眼  '(階段4)
'    End If
'    '======================
'    If turnatk = 2 And atkingck(94, 2) = 1 Then
'       atkingck(94, 1) = 2
'       技能.夏洛特_大聖堂  '(階段2)
'    End If
'    If turnatk = 2 And atkingckai(48, 2) = 1 Then
'       atkingckai(48, 1) = 2
'       AI技能.傑多_因果之幻  '(階段2)
'    End If
'==================
'技能_CC_原子之心_跳入點: '技能-C.C.-原子之心-程式跳入點
'=================
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
'======================
'  For p = 1 To 攻擊防禦骰子總數(1)
'     Randomize
'     i = Int(Rnd() * 6) + 1
'     If i = 1 Or i = 6 Then 擲骰表單溝通暫時變數(7) = Val(擲骰表單溝通暫時變數(7)) + 1
'  Next
'  For q = 1 To 攻擊防禦骰子總數(2)
'    Randomize
'     j = Int(Rnd() * 6) + 1
'     If j = 1 Or j = 6 Then 擲骰表單溝通暫時變數(8) = Val(擲骰表單溝通暫時變數(8)) + 1
'  Next
  '==================
'  If turnatk = 2 And atkingck(10, 2) = 1 Then
'       atkingck(10, 1) = 2
'       技能.雪莉_異質者  '(階段2)
'  End If
'  If turnatk = 2 And atkingck(43, 2) = 1 Then
'       atkingck(43, 1) = 2
'       技能.雪莉_VBE_異質者  '(階段2)
'  End If
'  If turnatk = 2 And atkingck(141, 2) = 1 Then
'       atkingck(141, 1) = 2
'       技能.多妮妲_異質者  '(階段2)
'  End If
  '=================
'===========================執行階段插入點(ATK-12/DEF-32)
執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 12, 2
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 32, 2
'============================
   階段狀態數 = 4
'   atkingtrtot.Interval = 600
'   atkingtrtot.Enabled = True
   防禦階段_階段初始.Enabled = False
   '============================
   HP檢查變數 = True
   HP檢查階段數 = 2
   目前數(10) = 1
   收牌階段_計算.Enabled = True
End Sub

Sub 使用者出牌_AI出牌控制_Timer()
If turnpageonin = 1 And 牌移動.Enabled = False Then
    If Val(pagecardnum(目前數(32), 11)) = 3 And Val(pagecardnum(目前數(32), 5)) = 1 And Val(pagecardnum(目前數(32), 6)) = 1 Then
        FormMainMode.card_CardClick (目前數(32))
    End If
    目前數(32) = 目前數(32) + 1
    If 目前數(32) > 公用牌實體卡片分隔紀錄數(1) Then
        使用者出牌_AI出牌控制.Enabled = False
        目前數(24) = 47
        等待時間_2.Enabled = True
    End If
End If
End Sub

Sub 使用者出牌_AI出牌控制_事件卡_Timer()
If turnpageonin = 1 And 牌移動.Enabled = False Then
    For i = 公用牌實體卡片分隔紀錄數(2) + 1 To 公用牌實體卡片分隔紀錄數(4)
        If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
            If pagecardnum(i, 1) = a6a Then
                FormMainMode.card_CardClick (i)
                Exit Sub
            End If
            If pagecardnum(i, 1) = a7a And (turnatk = 1 Or turnatk = 2) Then
                FormMainMode.card_CardClick (i)
                Exit Sub
            End If
            If pagecardnum(i, 1) = a8a Then
                FormMainMode.card_CardClick (i)
                Exit Sub
            End If
            If pagecardnum(i, 1) = a9a Then
                FormMainMode.card_CardClick (i)
                Exit Sub
            End If
        End If
    Next
    If i = 公用牌實體卡片分隔紀錄數(4) + 1 Then
        使用者出牌_AI出牌控制_事件卡.Enabled = False
        目前數(24) = 46
        等待時間_2.Enabled = True
    End If
End If
End Sub


Private Sub 使用者出牌_手牌對齊_Timer()
For i = 1 To Val(pageusglead)
   If 出牌順序統計暫時變數(2, i, 1) > 目前數(5) Then
      If 目前數(13) = 0 Then
         If card(出牌順序統計暫時變數(2, i, 2)).Left = 2640 And card(出牌順序統計暫時變數(2, i, 2)).Top = 7980 Then  '指定第2列第1張牌
              目前數(13) = 出牌順序統計暫時變數(2, i, 2)
              pagecardnum(目前數(13), 9) = card(目前數(13)).Left  '指定目前Left(座標)
              pagecardnum(目前數(13), 10) = card(目前數(13)).Top  '指定目前Top(座標)
              '==========戰鬥系統類.計算牌移動距離單位
             距離單位_收牌暫時數(1, 1) = (9840 - pagecardnum(目前數(13), 9)) \ 10 '計算Left
             距離單位_收牌暫時數(1, 2) = -((pagecardnum(目前數(13), 10) - 6700) \ 10)  '計算Top
          End If
     End If
     If 目前數(13) = 出牌順序統計暫時變數(2, i, 2) Then
             card(目前數(13)).Left = card(目前數(13)).Left + 距離單位_收牌暫時數(1, 1)
             card(目前數(13)).Top = card(目前數(13)).Top + 距離單位_收牌暫時數(1, 2)
     Else
             card(出牌順序統計暫時變數(2, i, 2)).Left = card(出牌順序統計暫時變數(2, i, 2)).Left - (900 / 10)
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
'            If atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 10 Then
'               atkingck(59, 1) = 4
'               技能.伊芙琳_赤紅石榴  '(階段4)
'               Exit Sub
'           ElseIf atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
'               atkingck(59, 1) = 6
'               技能.伊芙琳_赤紅石榴  '(階段6)
'               Exit Sub
'           ElseIf atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
'               atkingck(59, 1) = 9
'               技能.伊芙琳_赤紅石榴  '(階段9)
'               Exit Sub
'           End If
       Case 3
           '===========事件卡執行_詛咒術_電腦(階段3)
            事件卡記錄暫時數(2, 3) = 3
            事件卡.詛咒術_電腦 0, 0
       Case 4
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards")) = 2 '(階段2)
            End If
'            If atkingckai(66, 2) = 1 Then
'               atkingckai(66, 1) = 4
'               AI技能.艾伯李斯特_雷擊  '(階段4)
'               Exit Sub
'            End If
       Case 5
             If 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards")) = 2 '(階段2)
            End If
'            If atkingckai(78, 2) = 1 Then
'               atkingckai(78, 1) = 4
'               AI技能.瑪格莉特_月光  '(階段4)
'               Exit Sub
'            End If
'       Case 6
'            If atkingckai(79, 2) = 1 Then
'               atkingckai(79, 1) = 4
'               AI技能.庫勒尼西_瘋狂眼窩  '(階段4)
'               Exit Sub
'            End If
'        Case 7
'            If atkingckai(87, 2) = 1 Then
'               atkingckai(87, 1) = 3
'               AI技能.洛洛妮_貪婪之刃與嗜血之槍  '(階段3)
'               Exit Sub
'            End If
'        Case 8
'            If atkingckai(105, 2) = 1 Then
'               atkingckai(105, 1) = 5
'               AI技能.古魯瓦爾多_精神力吸收  '(階段5)
'               Exit Sub
'            End If
'        Case 9
'            If atkingckai(107, 2) = 1 Then
'               AI技能.伊芙琳_怠惰的墓表  '(階段4/5)
'               Exit Sub
'            End If
'        Case 10
'            If atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 10 Then
'               atkingckai(108, 1) = 4
'               AI技能.伊芙琳_赤紅石榴  '(階段4)
'               Exit Sub
'           ElseIf atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
'               atkingckai(108, 1) = 6
'               AI技能.伊芙琳_赤紅石榴  '(階段6)
'               Exit Sub
'           ElseIf atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
'               atkingckai(108, 1) = 9
'               AI技能.伊芙琳_赤紅石榴  '(階段9)
'               Exit Sub
'           End If
        Case 11
            目前數(24) = 38
            等待時間_2.Enabled = True
    End Select
End If
End Sub



Private Sub 使用者出牌_出牌對齊_靠右_Timer()
For i = 1 To pageusqlead
   If 出牌順序統計暫時變數(1, i, 1) < 目前數(5) Then
      card(出牌順序統計暫時變數(1, i, 2)).Left = card(出牌順序統計暫時變數(1, i, 2)).Left + (480 / 10)
   End If
   If 出牌順序統計暫時變數(1, i, 1) > 目前數(5) Then
      card(出牌順序統計暫時變數(1, i, 2)).Left = card(出牌順序統計暫時變數(1, i, 2)).Left - (500 / 10)
   End If
Next
目前數(3) = 目前數(3) + (480 / 10)
If 目前數(3) >= 480 Then
    使用者出牌_出牌對齊_靠右.Enabled = False
'    對齊完成檢查.Enabled = True
End If
End Sub

Private Sub 使用者出牌_出牌對齊_靠左_Timer()
For i = 1 To (pageusqlead - 1)
   card(出牌順序統計暫時變數(1, i, 2)).Left = card(出牌順序統計暫時變數(1, i, 2)).Left - (480 / 10)
Next
目前數(3) = 目前數(3) + (480 / 10)
If 目前數(3) >= 480 Then
    使用者出牌_出牌對齊_靠左.Enabled = False
'    對齊完成檢查.Enabled = True
End If
End Sub



Private Sub 移動階段_階段初始_Timer()
If 目前數(31) = 0 Then
    目前數(31) = 1
    Dim movecpn As Integer
    movecpn = movecp
    moveus = 0
    movecheckus = 0
    '===============
    movecom = atkingpagetot(2, 3)
    moveus = atkingpagetot(1, 3)
    '=====================以下是異常狀態檢查及啟動
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 21) = True And livecom(角色人物對戰人數(2, 2)) <= 1 Then
    '      異常狀態檢查數(21, 1) = 2
    '      異常狀態.中毒_電腦  '(階段2)
    'End If
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 3) = True Then
    '      異常狀態檢查數(3, 1) = 1
    '      異常狀態.MOV加_電腦  '(階段1)
    'End If
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 6) = True Then
    '      異常狀態檢查數(6, 1) = 1
    '      異常狀態.MOV減_電腦  '(階段1)
    'End If
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(2, 17) = True Then
    '      異常狀態檢查數(17, 1) = 1
    '      異常狀態.麻痺_電腦  '(階段1)
    'End If
    '===========================執行階段插入點(70)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 70, 1
    '============================
    'movecheckcom = movecom
    '顯示列1.電腦方移動值 = movecheckcom
    '========================================
    
    '===========
    'atkingtrn(1) = Val(atkingtrn(1)) + Val(atkingtrn(3))
    'atkingtrn(2) = Val(atkingtrn(2)) + Val(atkingtrn(4))
    'atkingtrn(3) = 0
    'atkingtrn(4) = 0

    '=====================================================
    
    '===============以下是技能檢查及啟動
    
    '===============以下是異常狀態檢查及啟動
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(1, 9) = True Then
    '      異常狀態檢查數(9, 1) = 1
    '      異常狀態.MOV加_使用者  '(階段1)
    'End If
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(1, 12) = True Then
    '      異常狀態檢查數(12, 1) = 1
    '      異常狀態.MOV減_使用者  '(階段1)
    'End If
    'If turnatk = 3 And 執行動作_檢查是否有指定異常狀態(1, 16) = True Then
    '      異常狀態檢查數(16, 1) = 1
    '      異常狀態.麻痺_使用者  '(階段1)
    'End If
    '=================================
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
            回復執行_電腦 1, 1
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
         回復執行_使用者 1, 1
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
    ReDim VBEStageNum(0 To 2) As Integer
    VBEStageNum(0) = 71
    VBEStageNum(1) = moveus '使用者方總移動數
    VBEStageNum(2) = movecom '電腦方總移動數
    '===========================執行階段插入點(71)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 71, 1
    '============================
    
    If movecpn < 1 Then
       movecpn = 1
    ElseIf movecpn > 3 Then
       movecpn = 3
    End If
    
    執行動作_距離變更 movecpn, True
    
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
    
'    If Val(顯示列1.使用者方移動值) > 6 Then
'        顯示列1.使用者方移動值 = 6
'    End If
'    If Val(顯示列1.電腦方移動值) > 6 Then
'        顯示列1.電腦方移動值 = 6
'    End If
    
    擲骰表單溝通暫時變數(4) = moveturn
    '技能動畫顯示階段數 = 2
    '戰鬥系統類.技能啟動數量檢查
    HP檢查變數 = False
    目前數(24) = 23
    FormMainMode.等待時間_2.Enabled = True
Else
   '============以下是技能檢查及啟動
'   If turnatk = 3 And atkingck(4, 2) = 1 Then
'      atkingck(4, 1) = 2
'      技能.雪莉_巨大黑犬 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(44, 2) = 1 Then
'      atkingck(44, 1) = 2
'      技能.雪莉_VBE_巨大黑犬 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(2, 2) = 1 Then
'      atkingckai(2, 1) = 3
'      AI技能.雪莉_巨大黑犬 '(階段3)
'   End If
'   If turnatk = 3 And atkingck(105, 2) = 1 Then
'      atkingck(105, 1) = 2
'      技能.艾蕾可_救濟天使  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(94, 2) = 1 Then
'      atkingckai(94, 1) = 2
'      AI技能.艾蕾可_救濟天使  '(階段2)
'   End If
'   If turnatk = 3 And atkingck(7, 2) = 1 Then
'      atkingck(7, 1) = 2
'      技能.帕茉_憤怒之爪  '(階段2)
'   End If
'   If turnatk = 3 And atkingck(12, 2) = 1 Then
'      atkingck(12, 1) = 2
'      技能.古魯瓦爾多_必殺架勢 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(61, 2) = 1 Then
'      atkingck(61, 1) = 2
'      技能.古魯瓦爾多_精神力吸收 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(21, 2) = 1 Then
'      atkingck(21, 1) = 2
'      技能.史塔夏_殺戮器官 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(37, 2) = 1 Then
'      atkingck(37, 1) = 2
'      技能.艾茵_一顆心 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(59, 2) = 1 Then
'      atkingck(59, 1) = 2
'      技能.伊芙琳_赤紅石榴 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(67, 2) = 1 Then
'      atkingck(67, 1) = 2
'      技能.音音夢_美味牛奶 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(74, 2) = 1 Then
'      atkingck(74, 1) = 2
'      技能.艾伯李斯特_智略 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(82, 2) = 1 Then
'      atkingck(82, 1) = 2
'      技能.布勞_發條機構 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(85, 2) = 1 Then
'      atkingck(85, 1) = 2
'      技能.布勞_夜幕時分 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(89, 2) = 1 Then
'      atkingck(89, 1) = 2
'      技能.阿貝爾_抽刀斷水計 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(96, 2) = 1 Then
'      atkingck(96, 1) = 2
'      技能.夏洛特_夜未央 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(97, 2) = 1 Then
'      atkingck(97, 1) = 2
'      技能.夏洛特_幸福的理由 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(138, 2) = 1 Then
'      atkingck(138, 1) = 2
'      技能.蕾格烈芙_SSS '(階段2)
'   End If
'   If turnatk = 3 And atkingck(142, 2) = 1 Then
'      atkingck(142, 1) = 2
'      技能.多妮妲_超級女主角 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(144, 2) = 1 Then
'      atkingck(144, 1) = 2
'      技能.傑多_因果之線 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(149, 2) = 1 Then
'      atkingck(149, 1) = 2
'      技能.阿奇波爾多_大地崩壞 '(階段2)
'   End If
'   If turnatk = 3 And atkingck(49, 2) = 1 Then
'      atkingck(49, 1) = 2
'      技能.尤莉卡_超載 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(6, 2) = 1 Then
'      atkingckai(6, 1) = 2
'      AI技能.南瓜王_超再生 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(10, 2) = 1 Then
'      atkingckai(10, 1) = 2
'      AI技能.妖精王妃_混沌之翼 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(29, 2) = 1 Then
'      atkingckai(29, 1) = 2
'      AI技能.音音夢_美味牛奶 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(68, 2) = 1 Then
'      atkingckai(68, 1) = 2
'      AI技能.艾伯李斯特_智略 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(88, 2) = 1 Then
'      atkingckai(88, 1) = 2
'      AI技能.史塔夏_殺戮器官 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(89, 2) = 1 Then
'      atkingckai(89, 1) = 2
'      AI技能.阿奇波爾多_大地崩壞 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(104, 2) = 1 Then
'      atkingckai(104, 1) = 2
'      AI技能.古魯瓦爾多_必殺架勢 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(105, 2) = 1 Then
'      atkingckai(105, 1) = 2
'      AI技能.古魯瓦爾多_精神力吸收  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(106, 2) = 1 Then
'      atkingckai(106, 1) = 2
'      AI技能.帕茉_憤怒之爪  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(108, 2) = 1 Then
'      atkingckai(108, 1) = 2
'      AI技能.伊芙琳_赤紅石榴  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(109, 2) = 1 Then
'      atkingckai(109, 1) = 2
'      AI技能.布勞_發條機構  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(110, 2) = 1 Then
'      atkingckai(110, 1) = 2
'      AI技能.布勞_夜幕時分  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(113, 2) = 1 Then
'      atkingckai(113, 1) = 2
'      AI技能.阿貝爾_抽刀斷水計  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(114, 2) = 1 Then
'      atkingckai(114, 1) = 2
'      AI技能.夏洛特_夜未央  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(115, 2) = 1 Then
'      atkingckai(115, 1) = 2
'      AI技能.夏洛特_幸福的理由  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(117, 2) = 1 Then
'      atkingckai(117, 1) = 2
'      AI技能.蕾格烈芙_SSS  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(118, 2) = 1 Then
'      atkingckai(118, 1) = 2
'      AI技能.多妮妲_超級女主角  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(119, 2) = 1 Then
'      atkingckai(119, 1) = 2
'      AI技能.傑多_因果之線  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(135, 2) = 1 Then
'      atkingckai(135, 1) = 2
'      AI技能.艾茵_一顆心  '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(139, 2) = 1 Then
'      atkingckai(139, 1) = 2
'      AI技能.尤莉卡_超載  '(階段2)
'   End If
'   '======================距離相關類(使用者)
'   If turnatk = 3 And atkingck(56, 2) = 1 Then
'      atkingck(56, 1) = 2
'      技能.伊芙琳_怠惰的墓表 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(107, 2) = 1 Then
'      atkingckai(107, 1) = 2
'      AI技能.伊芙琳_怠惰的墓表 '(階段2)
'   End If
'    If turnatk = 3 And atkingck(124, 2) = 1 Then
'      atkingck(124, 1) = 2
'      技能.瑪格莉特_末日幻影 '(階段2)
'   End If
'   If turnatk = 3 And atkingckai(116, 2) = 1 Then
'      atkingckai(116, 1) = 2
'      AI技能.瑪格莉特_末日幻影 '(階段2)
'   End If
'   '==========以下是異常狀態繼承回合／檢查及啟動(特殊)
'      異常狀態檢查數(15, 1) = 1
'      異常狀態.自壞_使用者  '(階段1)
'      '========
'      異常狀態檢查數(16, 1) = 2
'      異常狀態.麻痺_使用者  '(階段2)
'      '========
'      異常狀態檢查數(37, 1) = 1
'      異常狀態.再生_使用者  '(階段1)
'     '========
'      異常狀態檢查數(38, 1) = 1
'      異常狀態.再生_電腦  '(階段1)
'      '==========
'      異常狀態檢查數(20, 1) = 1
'      異常狀態.中毒_使用者  '(階段1)
'      '========
'      異常狀態檢查數(17, 1) = 2
'      異常狀態.麻痺_電腦  '(階段2)
'      '========
'      異常狀態檢查數(19, 1) = 1
'      異常狀態.自壞_電腦  '(階段1)
'      '========
'      異常狀態檢查數(21, 1) = 1
'      異常狀態.中毒_電腦  '(階段1)
'   '==============
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
    戰鬥系統類.出牌順序計算_電腦_出牌
    電腦出牌_亮牌.Enabled = True
End If
移動階段_階段初始.Enabled = False
End Sub

Sub 移動階段_階段前啟動_Timer()
'atkingtrn(3) = atkingtrn(1)
'atkingtrn(4) = atkingtrn(2)
'atkingtrn(1) = 0
'atkingtrn(2) = 0
'=================以下是技能檢查及啟動(移動階段開始階段1)
'If turnatk = 3 And atkingck(153, 2) = 1 Then
'   atkingck(153, 1) = 2
'   技能.洛洛妮_逆轉戰局的槍響  '(階段2)
'End If
'If turnatk = 3 And atkingck(160, 2) = 1 Then
'   atkingck(160, 1) = 2
'   技能.克頓_惡意情報  '(階段2)
'End If
'If turnatk = 3 And atkingck(108, 2) = 1 Then
'   atkingck(108, 1) = 2
'   技能.梅莉_綿羊幻夢  '(階段2)
'End If
'If turnatk = 3 And atkingckai(101, 2) = 1 Then
'   atkingckai(101, 1) = 2
'   AI技能.梅莉_綿羊幻夢  '(階段2)
'End If
'If turnatk = 3 And atkingck(110, 2) = 1 Then
'   atkingck(110, 1) = 2
'   技能.貝琳達_雪光 '(階段2)
'End If
'If turnatk = 3 And atkingckai(122, 2) = 1 Then
'   atkingckai(122, 1) = 2
'   AI技能.貝琳達_雪光 '(階段2)
'End If
'If turnatk = 3 And atkingckai(130, 2) = 1 Then
'   atkingckai(130, 1) = 2
'   AI技能.洛洛妮_逆轉戰局的槍響 '(階段2)
'End If
'If turnatk = 3 And atkingckai(134, 2) = 1 Then
'   atkingckai(134, 1) = 2
'   AI技能.克頓_惡意情報 '(階段2)
'End If
'=================
'技能動畫顯示階段數 = 5
'戰鬥系統類.技能啟動數量檢查
'=================以下是技能檢查及啟動(移動階段開始階段2)
'If turnatk = 3 And atkingck(153, 2) = 1 Then
'   atkingck(153, 1) = 3
'   技能.洛洛妮_逆轉戰局的槍響  '(階段3)
'End If
'If turnatk = 3 And atkingck(160, 2) = 1 Then
'   atkingck(160, 1) = 3
'   技能.克頓_惡意情報  '(階段3)
'End If
'If turnatk = 3 And atkingck(108, 2) = 1 Then
'   atkingck(108, 1) = 3
'   技能.梅莉_綿羊幻夢  '(階段3)
'End If
'If turnatk = 3 And atkingckai(101, 2) = 1 Then
'   atkingckai(101, 1) = 3
'   AI技能.梅莉_綿羊幻夢  '(階段3)
'End If
'If turnatk = 3 And atkingck(110, 2) = 1 Then
'   atkingck(110, 1) = 3
'   技能.貝琳達_雪光 '(階段3)
'End If
'If turnatk = 3 And atkingckai(122, 2) = 1 Then
'   atkingckai(122, 1) = 3
'   AI技能.貝琳達_雪光 '(階段3)
'End If
'If turnatk = 3 And atkingckai(130, 2) = 1 Then
'   atkingckai(130, 1) = 3
'   AI技能.洛洛妮_逆轉戰局的槍響 '(階段3)
'End If
'If turnatk = 3 And atkingckai(134, 2) = 1 Then
'   atkingckai(134, 1) = 3
'   AI技能.克頓_惡意情報 '(階段3)
'End If
'=================
'atkingtrtot.Interval = 600
'atkingtrtot.Enabled = True
Erase Vss_PersonMoveActionChangeNum
Erase Vss_PersonMoveControlNum
'===========================執行階段插入點(2)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 2, 1
'============================
'===========================執行階段插入點(3)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 3, 1
'============================
'===========================執行階段插入點(4)
執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 4, 1
'============================
移動階段_階段前啟動.Enabled = False
目前數(22) = 26
目前數(31) = 0
等待時間.Enabled = True
End Sub

Private Sub 移動圖片完成檢查_Timer()
If 顯示列1.移動方向圖片顯示 = False Then
   收牌階段_計算.Enabled = True
   移動圖片完成檢查.Enabled = False
   bnok.Visible = False
End If
End Sub

Sub 智慧型AI_使用者出牌_Timer()

End Sub

Private Sub 牌移動_Timer()
card(牌移動暫時變數(3)).Left = card(牌移動暫時變數(3)).Left + 距離單位(2, 1, 1)
card(牌移動暫時變數(3)).Top = card(牌移動暫時變數(3)).Top + 距離單位(2, 1, 2)
If Abs(牌移動暫時變數(1) - card(牌移動暫時變數(3)).Left) <= 50 Or Abs(牌移動暫時變數(2) - card(牌移動暫時變數(3)).Top) <= 50 Then
   card(牌移動暫時變數(3)).Left = 牌移動暫時變數(1)
   card(牌移動暫時變數(3)).Top = 牌移動暫時變數(2)
   card(牌移動暫時變數(3)).ZOrder
   For i = 1 To 3
       compiin(i).ZOrder
   Next
   牌移動.Enabled = False
   Select Case 目前數(15)
      Case 1
          發牌檢查.Enabled = True
      Case 2
          戰鬥系統類.出牌順序計算_電腦_手牌
          目前數(8) = 0
          電腦出牌_手牌對齊.Enabled = True
      Case 3
'          If turnatk = 3 And atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
'               atkingck(59, 1) = 8
'               技能.伊芙琳_赤紅石榴  '(階段8)
'               Exit Sub
'          ElseIf turnatk = 3 And atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
'               atkingck(59, 1) = 11
'               技能.伊芙琳_赤紅石榴  '(階段11)
'               Exit Sub
'          End If
      Case 4
             card(目前數(20)).Visible = False
            目前數(4) = 0
            目前數(13) = 0
            戰鬥系統類.出牌順序計算_使用者_手牌
            使用者出牌_手牌對齊.Enabled = True
      Case 5
           card(目前數(16)).Visible = False
           戰鬥系統類.出牌順序計算_電腦_手牌
          目前數(8) = 0
          電腦出牌_手牌對齊.Enabled = True
       Case 6
          '===========事件卡執行_機會_使用者(階段2)
          card(事件卡記錄暫時數(1, 4)).Visible = False
          pagecardnum(事件卡記錄暫時數(1, 4), 6) = 3
          目前數(24) = 6
          等待時間_2.Enabled = True
        Case 7
           '===========事件卡執行_機會_使用者(階段1)
          目前數(24) = 5
          等待時間_2.Enabled = True
        Case 8
           '===========事件卡執行_機會_使用者(階段3)
           事件卡記錄暫時數(1, 3) = 3
           事件卡.機會_使用者 0, 0
        Case 9
            '===========事件卡執行_機會_電腦(階段1)
           目前數(24) = 7
           等待時間_2.Enabled = True
        Case 10
           '===========事件卡執行_機會_電腦(階段3)
          card(事件卡記錄暫時數(2, 4)).Visible = False
          pagecardnum(事件卡記錄暫時數(2, 4), 6) = 3
          目前數(24) = 8
          等待時間_2.Enabled = True
        Case 11
           '===========事件卡執行_機會_電腦(階段4)
           事件卡記錄暫時數(2, 3) = 4
           事件卡.機會_電腦 0, 0
        Case 12
           '===========事件卡執行_詛咒術_使用者(階段1)
           目前數(24) = 11
           等待時間_2.Enabled = True
        Case 13
           '===========事件卡執行_詛咒術_使用者(階段6)
           card(事件卡記錄暫時數(1, 4)).Visible = False
           pagecardnum(事件卡記錄暫時數(1, 4), 6) = 3
           事件卡記錄暫時數(1, 3) = 6
           事件卡.詛咒術_使用者 0, 0
        Case 14
           '===========事件卡執行_詛咒術_電腦(階段1)
           目前數(24) = 13
           等待時間_2.Enabled = True
        Case 15
           '===========事件卡執行_詛咒術_電腦(階段5>6)
           card(事件卡記錄暫時數(2, 4)).Visible = False
           pagecardnum(事件卡記錄暫時數(2, 4), 6) = 3
           事件卡記錄暫時數(2, 3) = 6
           事件卡.詛咒術_電腦 0, 0
        Case 16
           '===========事件卡執行_HP回復_使用者(階段1)
           目前數(24) = 16
           等待時間_2.Enabled = True
           turnpageonin = 0
           FormMainMode.bnok.Enabled = False
        Case 17
           '===========事件卡執行_HP回復_使用者(階段4)
           card(事件卡記錄暫時數(1, 4)).Visible = False
           pagecardnum(事件卡記錄暫時數(1, 4), 6) = 3
           事件卡記錄暫時數(1, 3) = 4
           事件卡.HP回復_使用者 0, 0
        Case 18
           '===========事件卡執行_HP回復_電腦(階段1)
           目前數(24) = 18
           等待時間_2.Enabled = True
        Case 19
           '===========事件卡執行_HP回復_電腦(階段4>5)
           card(事件卡記錄暫時數(2, 4)).Visible = False
           pagecardnum(事件卡記錄暫時數(2, 4), 6) = 3
           事件卡記錄暫時數(2, 3) = 5
           事件卡.HP回復_電腦 0, 0
        Case 20
           戰鬥系統類.出牌順序計算_使用者_手牌
           目前數(4) = 0
           目前數(13) = 0
           使用者出牌_手牌對齊.Enabled = True
        Case 21
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDrawCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDrawCards")) = 1 '(階段1)
            End If
'           If turnatk = 2 And atkingck(54, 2) = 1 Then
'               atkingck(54, 1) = 4
'               技能.羅莎琳_黑霧幻影  '(階段4)
'               Exit Sub
'          End If
'          If turnatk = 2 And atkingck(55, 2) = 1 Then
'               atkingck(55, 1) = 4
'               技能.羅莎琳_EX_黑霧幻影  '(階段4)
'               Exit Sub
'          End If
        Case 22
           If 執行階段系統_搜尋正在執行之執行階段("AtkingGetUsedCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGetUsedCards")) = 2 '(階段2)
            End If
'           If turnatk = 2 And atkingck(64, 2) = 1 Then
'               atkingck(64, 1) = 5
'               技能.梅倫_Jackpot  '(階段5)
'               Exit Sub
'          End If
'          If turnatk = 1 And atkingckai(31, 2) = 1 Then
'               atkingckai(31, 1) = 5
'               AI技能.梅倫_Jackpot  '(階段5)
'               Exit Sub
'          End If
        Case 23
            If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 3 '(階段3)
            End If
'           If turnatk = 3 And atkingck(74, 2) = 1 Then
'               atkingck(74, 1) = 3
'               技能.艾伯李斯特_智略  '(階段3)
'               Exit Sub
'          End If
'          If turnatk = 3 And atkingckai(68, 2) = 1 Then
'               atkingckai(68, 1) = 3
'               AI技能.艾伯李斯特_智略  '(階段3)
'               Exit Sub
'          End If
'        Case 24
'          If turnatk = 3 And atkingck(82, 2) = 1 Then
'               atkingck(82, 1) = 4
'               技能.布勞_發條機構  '(階段4)
'               Exit Sub
'          End If
'        Case 25
'          If turnatk = 2 And atkingck(92, 2) = 1 Then
'               atkingck(92, 1) = 3
'               技能.利恩_反擊的狼煙  '(階段3)
'               Exit Sub
'          End If
'          If turnatk = 1 And atkingckai(74, 2) = 1 Then
'               atkingckai(74, 1) = 3
'               AI技能.利恩_反擊的狼煙  '(階段3)
'               Exit Sub
'          End If
'        Case 26
'          If turnatk = 2 And atkingck(146, 2) = 1 Then
'               atkingck(146, 1) = 5
'               技能.傑多_因果之刻 '(階段5)
'               Exit Sub
'          End If
'        Case 27
'          If turnatk = 3 And atkingck(153, 2) = 1 Then
'               atkingck(153, 1) = 4
'               技能.洛洛妮_逆轉戰局的槍響 '(階段4)
'               Exit Sub
'          End If
'        Case 28
'          If turnatk = 3 And atkingck(160, 2) = 1 Then
'               atkingck(160, 1) = 4
'               技能.克頓_惡意情報 '(階段4)
'               Exit Sub
'          End If
'        Case 29
'          If turnatk = 3 And atkingck(108, 2) = 1 Then
'               atkingck(108, 1) = 4
'               技能.梅莉_綿羊幻夢 '(階段4)
'               Exit Sub
'          End If
'        Case 30
'          If turnatk = 3 And atkingckai(101, 2) = 1 Then
'               atkingckai(101, 1) = 4
'               AI技能.梅莉_綿羊幻夢 '(階段4)
'               Exit Sub
'          End If
'        Case 31
'           If turnatk = 3 And atkingck(110, 2) = 1 Then
'               atkingck(110, 1) = 4
'               技能.貝琳達_雪光 '(階段4)
'               Exit Sub
'          End If
'        Case 32
'           If turnatk = 2 And atkingck(111, 2) = 1 Then
'               atkingck(111, 1) = 5
'               技能.貝琳達_水晶幻鏡  '(階段5)
'               Exit Sub
'          End If
'        Case 33
'          If turnatk = 3 And atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
'               atkingckai(108, 1) = 8
'               AI技能.伊芙琳_赤紅石榴  '(階段8)
'               Exit Sub
'          ElseIf turnatk = 3 And atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
'               atkingckai(108, 1) = 11
'               AI技能.伊芙琳_赤紅石榴  '(階段11)
'               Exit Sub
'          End If
'        Case 34
'          If turnatk = 3 And atkingckai(109, 2) = 1 Then
'               atkingckai(109, 1) = 4
'               AI技能.布勞_發條機構  '(階段4)
'               Exit Sub
'          End If
'        Case 35
'          If turnatk = 1 And atkingckai(121, 2) = 1 Then
'               atkingckai(121, 1) = 5
'               AI技能.傑多_因果之刻 '(階段5)
'               Exit Sub
'          End If
'        Case 36
'           If turnatk = 3 And atkingckai(122, 2) = 1 Then
'               atkingckai(122, 1) = 4
'               AI技能.貝琳達_雪光 '(階段4)
'               Exit Sub
'          End If
'        Case 37
'           If turnatk = 1 And atkingckai(123, 2) = 1 Then
'               atkingckai(123, 1) = 5
'               AI技能.貝琳達_水晶幻鏡  '(階段5)
'               Exit Sub
'          End If
'        Case 38
'           If turnatk = 1 And atkingckai(128, 2) = 1 Then
'               atkingckai(128, 1) = 4
'               AI技能.羅莎琳_黑霧幻影  '(階段4)
'               Exit Sub
'          End If
'          If turnatk = 1 And atkingckai(129, 2) = 1 Then
'               atkingckai(129, 1) = 4
'               AI技能.羅莎琳_EX_黑霧幻影  '(階段4)
'               Exit Sub
'          End If
'       Case 39
'          If turnatk = 3 And atkingckai(130, 2) = 1 Then
'               atkingckai(130, 1) = 4
'               AI技能.洛洛妮_逆轉戰局的槍響 '(階段4)
'               Exit Sub
'          End If
      Case 40
          目前數(24) = 37
          等待時間_2.Enabled = True
      Case 41
           '===========事件卡執行_聖水_使用者(階段1)
           目前數(24) = 39
           等待時間_2.Enabled = True
           turnpageonin = 0
           FormMainMode.bnok.Enabled = False
      Case 42
           '===========事件卡執行_聖水_使用者(階段4>5)
           card(事件卡記錄暫時數(1, 4)).Visible = False
           pagecardnum(事件卡記錄暫時數(1, 4), 6) = 3
           事件卡記錄暫時數(1, 3) = 4
           事件卡.聖水_使用者 0, 0
      Case 43
           '===========事件卡執行_聖水_電腦(階段1)
           目前數(24) = 41
           等待時間_2.Enabled = True
      Case 44
           '===========事件卡執行_聖水_電腦(階段4>5)
           card(事件卡記錄暫時數(2, 4)).Visible = False
           pagecardnum(事件卡記錄暫時數(2, 4), 6) = 3
           事件卡記錄暫時數(2, 3) = 5
           事件卡.聖水_電腦 0, 0
   End Select
'   If turnatk = 4 Then
'      發牌檢查.Enabled = True
'   End If
End If
End Sub


Private Sub 牌移動_收牌_Timer()
If 目前數(11) = pageqlead(目前數(10)) Then
'   FormMainMode.wmpse1.Controls.stop
'    FormMainMode.wmpse1.Controls.play
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
'         MsgBox "收牌測試"
         card(距離單位_收牌暫時數(i, 3)).Visible = False
         pagecardnum(距離單位_收牌暫時數(i, 3), 6) = 3
         目前數(11) = 目前數(11) + 1
'         FormMainMode.wmpse1.Controls.stop
'         FormMainMode.wmpse1.Controls.play
     End If
     card(距離單位_收牌暫時數(i, 3)).Left = card(距離單位_收牌暫時數(i, 3)).Left + 距離單位_收牌暫時數(i, 1)
     card(距離單位_收牌暫時數(i, 3)).Top = card(距離單位_收牌暫時數(i, 3)).Top + 距離單位_收牌暫時數(i, 2)
     If 目前數(12) > 0 Then
         目前數(12) = 目前數(12) - 1
     End If
Next

End Sub

Private Sub 發牌_使用者階段_Timer()
Dim m As Integer '暫時變數
'-----------使用者階段
'Do While 目前數(1) > Val(pageusglead) And 目前數(1) <= 牌總階段數(1)
Do While Val(pageusglead) < 牌總階段數(1)
          Randomize
          m = Int(Rnd() * Val(公用牌各牌類型紀錄數(0, 2))) + 1
          '===========
          If pagecardnum(m, 6) = 4 Then
             戰鬥系統類.getpage 1, m
             目前數(2) = 2
             發牌_使用者階段.Enabled = False
             Exit Sub
          End If
Loop
'If 目前數(1) > 牌總階段數(1) Or 目前數(1) <= Val(pageusglead) Then
If Val(pageusglead) >= 牌總階段數(1) Then
   發牌_使用者階段.Enabled = False
'   發牌_電腦階段.Enabled = True
   目前數(2) = 2
   發牌檢查.Enabled = True
End If
End Sub


Private Sub 發牌_電腦階段_Timer()
'-----------電腦階段
Dim m As Integer '暫時變數
'Do While 目前數(1) >= Val(pagecomglead) And 目前數(1) <= 牌總階段數(2)
Do While Val(pagecomglead) < 牌總階段數(2)
          Randomize
          m = Int(Rnd() * Val(公用牌各牌類型紀錄數(0, 2))) + 1
          '===========
          If pagecardnum(m, 6) = 4 Then
             戰鬥系統類.getpage 2, m
             目前數(2) = 3
             發牌_電腦階段.Enabled = False
             Exit Sub
          End If
Loop
If Val(pagecomglead) >= 牌總階段數(2) Then
   目前數(2) = 3
   發牌_電腦階段.Enabled = False
   發牌檢查.Enabled = True
End If
End Sub

Private Sub 發牌檢查_Timer()
'If 目前數(1) > 牌總階段數(3) Then
If (Val(pageusglead) >= 牌總階段數(1) And Val(pagecomglead) >= 牌總階段數(2)) Or BattleCardNum <= 0 Then
'   cnmove.Visible = True
   發牌檢查.Enabled = False
   目前數(15) = 0
   目前數(22) = 3
   等待時間.Enabled = True
Else
   '發牌_使用者階段.Enabled = True
   Select Case 目前數(2)
       Case 1
           發牌_使用者階段.Enabled = True
           發牌檢查.Enabled = False
       Case 2
           發牌_電腦階段.Enabled = True
           發牌檢查.Enabled = False
        Case 3
'           目前數(1) = 目前數(1) + 1
           目前數(2) = 1
           '發牌檢查.Enabled = True
    End Select
End If

End Sub


Private Sub 等待時間_2_Timer()
Select Case 目前數(14)
   Case 0
      目前數(14) = 目前數(14) + 1
   Case 1
      目前數(14) = 0
      等待時間_2.Enabled = False
      Select Case 目前數(24)
          Case 1
              '========開始初始階段1
                顯示列1.Visible = True
                顯示列1.移動階段圖顯示 = False
                顯示列1.移動方向圖片顯示 = False
                FormMainMode.wmpse6.Controls.play
                一般系統類.檢查音樂播放 6
                If 系統顯示界面紀錄數 = 1 Then
                    draw1.Visible = False
                    draw2.Visible = True
                Else
                    FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\draw2.gif"
                End If
                目前數(22) = 2
                等待時間.Enabled = True
          Case 2
              cn22_Click
              bnok.Visible = False
           Case 3
              cn32_Click
              bnok.Visible = False
           Case 4
              Select Case turnatk
                    Case 1
                        目前數(22) = 7
                        等待時間.Enabled = True
                    Case 2
                        目前數(22) = 8
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
                FormMainMode.移動階段_階段初始.Enabled = True
'               If atkingck(122, 2) = 1 Then
'                    atkingck(122, 1) = 6
'                    技能.瑪格莉特_月光 '(階段6)
'                End If
            Case 24
'                戰鬥系統類.擲骰表單顯示
                If FormMainMode.PEAFDiceInterface.DiceStop = True Or 骰數零檢查值(1) = True Or 骰數零檢查值(2) = True Then
                    戰鬥系統類.擲骰後續判斷
                    If 執行階段系統_搜尋正在執行之執行階段("BattleStartDice") <> 0 Then
'                        Dim uscomt As Integer
'                        Dim atknumtnum As Integer
'                        '================
'                        Select Case vbecommadnum(3, 執行階段系統_搜尋正在執行之執行階段("BattleStartDice"))
'                            Case Is <= 12 '==主動技-使用者
'                                uscomt = 1
'                                For pnnumt = 1 To 4
'                                    For atknumt = 1 To 4
'                                        If (uscomt - 1) * 12 + (4 * pnnumt - 4) + atknumt = vbecommadnum(3, 執行階段系統_搜尋正在執行之執行階段("BattleStartDice")) Then
'                                            atknumtnum = atknumt
'                                            pnnumt = 4 '跳離For
'                                        End If
'                                    Next
'                                Next
'                            Case Is <= 24 '==主動技-電腦
'                                uscomt = 2
'                                For pnnumt = 1 To 4
'                                    For atknumt = 1 To 4
'                                        If (uscomt - 1) * 12 + (4 * pnnumt - 4) + atknumt = vbecommadnum(3, 執行階段系統_搜尋正在執行之執行階段("BattleStartDice")) Then
'                                            atknumtnum = atknumt
'                                            pnnumt = 4 '跳離For
'                                        End If
'                                    Next
'                                Next
'                        End Select
'                        '================
                        vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("BattleStartDice")) = 2 '(階段2)
'                        執行指令集.執行指令_執行擲骰子 uscomt, 1, atknumtnum
                    End If
                Else
                    目前數(24) = 24
                    等待時間_2.Enabled = True
                End If
'               If atkingck(122, 2) = 1 Then
'                    atkingck(122, 1) = 7
'                    技能.瑪格莉特_月光 '(階段7)
'                End If
            Case 25
                If FormMainMode.PEAFDiceInterface.DiceStop = True Or 骰數零檢查值(1) = True Or 骰數零檢查值(2) = True Then
                    戰鬥系統類.擲骰後續判斷
                Else
                    目前數(24) = 25
                    等待時間_2.Enabled = True
                End If
'               If atkingckai(78, 2) = 1 Then
'                    atkingckai(78, 1) = 5
'                    AI技能.瑪格莉特_月光 '(階段5)
'                End If
'            Case 26
'               If atkingckai(78, 2) = 1 Then
'                    atkingckai(78, 1) = 6
'                    AI技能.瑪格莉特_月光 '(階段6)
'                End If
'            Case 27
'                If atkingck(153, 2) = 1 Then
'                    atkingck(153, 1) = 5
'                    技能.洛洛妮_逆轉戰局的槍響 '(階段5)
'                End If
'            Case 28
'                If atkingck(156, 2) = 1 Then
'                    atkingck(156, 1) = 5
'                    技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段5)
'                End If
'            Case 29
'                If atkingckai(87, 2) = 1 Then
'                    atkingckai(87, 1) = 4
'                    AI技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段4)
'                End If
            Case 30
                If 電腦出牌_亮牌.Enabled = False Then
                    顯示列1.移動方向圖片顯示 = True
                    移動圖片完成檢查.Enabled = True
                Else
                    目前數(24) = 30
                    等待時間_2.Enabled = True
                End If
'            Case 31
'                If atkingck(108, 2) = 1 Then
'                    atkingck(108, 1) = 5
'                    技能.梅莉_綿羊幻夢 '(階段5)
'                End If
'            Case 32
'                If atkingckai(101, 2) = 1 Then
'                    atkingckai(101, 1) = 5
'                    AI技能.梅莉_綿羊幻夢 '(階段5)
'                End If
'            Case 33
'                If atkingck(110, 2) = 1 Then
'                    atkingck(110, 1) = 5
'                    技能.貝琳達_雪光 '(階段5)
'                End If
'            Case 34
'                 If atkingckai(107, 2) = 1 Then
'                    atkingckai(107, 1) = 6
'                    AI技能.伊芙琳_怠惰的墓表 '(階段6)
'                 End If
'            Case 35
'                If atkingckai(122, 2) = 1 Then
'                    atkingckai(122, 1) = 5
'                    AI技能.貝琳達_雪光 '(階段5)
'                End If
'            Case 36
'                If atkingckai(130, 2) = 1 Then
'                    atkingckai(130, 1) = 5
'                    AI技能.洛洛妮_逆轉戰局的槍響 '(階段5)
'                End If
'            Case 37
'                If turnatk = 3 And atkingckai(134, 2) = 1 Then
'                     atkingckai(134, 1) = 4
'                     AI技能.克頓_惡意情報 '(階段4)
'                     Exit Sub
'                End If
'            Case 38
'                If atkingckai(134, 2) = 1 Then
'                   atkingckai(134, 1) = 5
'                   AI技能.克頓_惡意情報  '(階段5)
'                   Exit Sub
'                End If
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
                bnok.Enabled = False
                目前數(32) = 1
                FormMainMode.使用者出牌_AI出牌控制_事件卡.Enabled = True
            Case 46
                '====================試驗智慧型AI出牌系統
'                If 智慧型AI系統_目前可執行之人物判斷(nameus(角色人物對戰人數(1, 2))) = True Then
                    Dim wtyr As Integer '暫時變數
                    If (moveturn = 1 And turnatk = 2) Or (moveturn = 2 And turnatk = 1) Then wtyr = 1 Else wtyr = 0
                    智慧型AI系統類.智慧型AI系統計算_引導程序_選擇 1, turnatk, nameus(角色人物對戰人數(1, 2)), movecp, wtyr
                    智慧型AI系統類.智慧型AI系統_使用者出牌階段判斷反轉
                    目前數(32) = 1
                    FormMainMode.使用者出牌_AI出牌控制.Enabled = True
'                End If
            Case 47
                '=============使用者方選擇行動
                If turnatk = 3 Then
                    顯示列1.移動階段選擇值 = 目前數(33)
                End If
                '====================
                bnok.Enabled = True
                FormMainMode.bnok_Click
                For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
                    FormMainMode.card(ckl).CardEnabledType = True
                Next
            Case 48
                執行動作_電腦方各階段出牌完畢後行動 turnatk
      End Select
End Select
End Sub

Private Sub 等待時間_Timer()
Select Case 目前數(14)
   Case 0
      目前數(14) = 目前數(14) + 1
   Case 1
      目前數(14) = 0
      等待時間.Enabled = False
      Select Case 目前數(22)
          Case 1
'              If atkingck(56, 2) = 1 Then
'                  atkingck(56, 1) = 6
'                  技能.伊芙琳_怠惰的墓表 '(階段6)
'              End If
          Case 2   '========開始初始階段2
             目前數(22) = 5
             等待時間.Enabled = True
          Case 3
            目前數(22) = 22
            等待時間.Enabled = True
          Case 4
                戰鬥系統類.廣播訊息 "現在的距離" & movecp & "。"
                交換角色紀錄暫時變數(4) = 1
                戰鬥系統類.執行動作_移動階段選擇執行
'                Select Case moveturn
'                  Case 1
'                     cn2_Click
'                  Case 2
'                     cn3_Click
'                End Select
           Case 5
              cn1_Click
           Case 6
                Erase Vss_EventPlayerAllActionOffNum
                '===========================執行階段插入點(1)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 1, 1
                '============================
                cnmove_Click
           Case 7
              目前數(24) = 2
              等待時間_2.Enabled = True
           Case 8
              目前數(24) = 3
              等待時間_2.Enabled = True
           Case 9
               cn2_Click
               顯示列1.Visible = True
               戰鬥系統類.時間軸_顯示
           Case 10
               cn3_Click
               顯示列1.Visible = True
               戰鬥系統類.時間軸_顯示
           Case 11
              戰鬥系統類.時間軸_隱藏
              顯示列1.Visible = False
              目前數(22) = 12
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
'                擲骰表單溝通暫時變數(5) = 擲骰表單溝通暫時變數(7)
'                擲骰表單溝通暫時變數(6) = 擲骰表單溝通暫時變數(8)
                擲骰表單溝通暫時變數(9) = 攻擊防禦骰子總數(1)
                擲骰表單溝通暫時變數(10) = 攻擊防禦骰子總數(2)
                是否系統公骰 = True
                戰鬥系統類.擲骰表單顯示
                目前數(24) = 25
                FormMainMode.等待時間_2.Enabled = True
'                FormMainMode.骰子執行完啟動.Enabled = True
           Case 13
               目前數(22) = 9
               等待時間.Enabled = True
           Case 14
               目前數(22) = 10
               等待時間.Enabled = True
           Case 15
               cn4_Click
           Case 16
'               If atkingck(61, 2) = 1 Then
'                  atkingck(61, 1) = 6
'                  技能.古魯瓦爾多_精神力吸收 '(階段6)
'              End If
           Case 17
              '===========================執行階段插入點(9)
               執行階段系統類.執行階段系統總主要程序_執行階段開始 moveturn, 9, 1
              '============================
              Select Case moveturn
                  Case 1
                     cn2_Click
                  Case 2
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
                目前數(22) = 6
                等待時間.Enabled = True
            Case 23
'                If atkingck(122, 2) = 1 Then
'                    atkingck(122, 1) = 6
'                    技能.瑪格莉特_月光 '(階段6)
'                End If
'            Case 24
'                If atkingck(146, 2) = 1 Then
'                    atkingck(146, 1) = 5
'                    技能.傑多_因果之刻 '(階段5)
'                End If
'            Case 25
'                If atkingckai(78, 2) = 1 Then
'                    atkingckai(78, 1) = 5
'                    AI技能.瑪格莉特_月光 '(階段5)
'                End If
            Case 26
                移動階段_階段初始.Enabled = True
'            Case 27
'                If atkingck(156, 2) = 1 Then
'                    atkingck(156, 1) = 5
'                    技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段5)
'                End If
'            Case 28
'                If atkingckai(87, 2) = 1 Then
'                    atkingckai(87, 1) = 4
'                    AI技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段4)
'                End If
'            Case 29
'                If atkingck(111, 2) = 1 Then
'                    atkingck(111, 1) = 5
'                    技能.貝琳達_水晶幻鏡 '(階段5)
'                End If
            Case 30
                電腦出牌.Enabled = True
'            Case 31
'                 If atkingckai(105, 2) = 1 Then
'                    atkingckai(105, 1) = 6
'                    AI技能.古魯瓦爾多_精神力吸收 '(階段6)
'                 End If
'            Case 32
'                 If atkingckai(107, 2) = 1 Then
'                    atkingckai(107, 1) = 5
'                    AI技能.伊芙琳_怠惰的墓表 '(階段5)
'                 End If
'            Case 33
'                 If atkingck(59, 2) = 1 Then
'                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
'                        If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) < 2 Then
'                            目前數(22) = 33
'                            等待時間.Enabled = True
'                        Else
'                             atkingck(59, 1) = 12
'                             技能.伊芙琳_赤紅石榴 '(階段12)
'                        End If
'                 End If
'            Case 34
'                If atkingckai(121, 2) = 1 Then
'                    atkingckai(121, 1) = 5
'                    AI技能.傑多_因果之刻 '(階段5)
'                End If
'            Case 35
'                If atkingckai(123, 2) = 1 Then
'                    atkingckai(123, 1) = 5
'                    AI技能.貝琳達_水晶幻鏡 '(階段5)
'                End If
            Case 36
                FormMainMode.trend.Enabled = True
            Case 37
'                If atkingckai(87, 2) = 1 Then
'                    atkingckai(87, 1) = 5
'                    AI技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段5)
'                End If
      End Select
End Select
End Sub



Private Sub 電腦出牌_Timer()
 '=========================專屬事件卡出牌階段
For i = 公用牌實體卡片分隔紀錄數(2) + 1 To 公用牌實體卡片分隔紀錄數(4)
    If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
        If pagecardnum(i, 1) = a6a Then
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        ElseIf pagecardnum(i, 3) = a6a Then
            cspce = pagecardnum(i, 1)
            cspme = pagecardnum(i, 2)
            pagecardnum(i, 1) = pagecardnum(i, 3)
            pagecardnum(i, 2) = pagecardnum(i, 4)
            pagecardnum(i, 3) = cspce
            pagecardnum(i, 4) = cspme
            If pageonin(i) = 2 Then
               pageonin(i) = 1
            Else
               pageonin(i) = 2
            End If
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        End If
        If pagecardnum(i, 1) = a7a And (turnatk = 1 Or turnatk = 2) Then
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        ElseIf pagecardnum(i, 3) = a7a And (turnatk = 1 Or turnatk = 2) Then
            cspce = pagecardnum(i, 1)
            cspme = pagecardnum(i, 2)
            pagecardnum(i, 1) = pagecardnum(i, 3)
            pagecardnum(i, 2) = pagecardnum(i, 4)
            pagecardnum(i, 3) = cspce
            pagecardnum(i, 4) = cspme
            If pageonin(i) = 2 Then
               pageonin(i) = 1
            Else
               pageonin(i) = 2
            End If
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        End If
        If pagecardnum(i, 1) = a8a Then
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        ElseIf pagecardnum(i, 3) = a8a Then
            cspce = pagecardnum(i, 1)
            cspme = pagecardnum(i, 2)
            pagecardnum(i, 1) = pagecardnum(i, 3)
            pagecardnum(i, 2) = pagecardnum(i, 4)
            pagecardnum(i, 3) = cspce
            pagecardnum(i, 4) = cspme
            If pageonin(i) = 2 Then
               pageonin(i) = 1
            Else
               pageonin(i) = 2
            End If
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        End If
        If pagecardnum(i, 1) = a9a Then
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        ElseIf pagecardnum(i, 3) = a9a Then
            cspce = pagecardnum(i, 1)
            cspme = pagecardnum(i, 2)
            pagecardnum(i, 1) = pagecardnum(i, 3)
            pagecardnum(i, 2) = pagecardnum(i, 4)
            pagecardnum(i, 3) = cspce
            pagecardnum(i, 4) = cspme
            If pageonin(i) = 2 Then
               pageonin(i) = 1
            Else
               pageonin(i) = 2
            End If
            pagecardnum(i, 11) = 1
            電腦牌_模擬按牌 i
            電腦出牌.Enabled = False
            Exit Sub
        End If
    End If
Next
If 電腦方事件卡是否出完選擇數 = False Then
    電腦方事件卡是否出完選擇數 = True
    電腦出牌.Enabled = False
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
            If 目前數(6) > 公用牌實體卡片分隔紀錄數(4) Then
                電腦出牌.Enabled = False
                電腦方事件卡是否出完選擇數 = False
                Select Case turnatk
                   Case 1
                      目前數(6) = 0
                      目前數(10) = 1
                      戰鬥系統類.時間軸_停止
                      戰鬥系統類.出牌順序計算_電腦_出牌
                      電腦出牌_亮牌.Enabled = True
                      trgoi2_Timer
                   Case 2
                      目前數(6) = 0
                      目前數(10) = 1
                      戰鬥系統類.時間軸_停止
                      戰鬥系統類.出牌順序計算_電腦_出牌
                      電腦出牌_亮牌.Enabled = True
                      trgoi2_Timer
                      trgoi1_Timer
                   Case 3
'                      turnpageonin = 1
'                      階段狀態數 = 1
'        '              cnmove2.Visible = True
'                      bnok.Picture = LoadPicture(app_path & "gif\system\ok_1.jpg")
'                      bnok.Visible = True
'                      If Formsetting.chkusenewaipersonauto.Value = 1 Then
'                            For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
'                                FormMainMode.card(ckl).CardEnabledType = False
'                            Next
'                            目前數(24) = 45
'                            等待時間_2.Enabled = True
'                      End If
                        執行動作_電腦方各階段出牌完畢後行動 3
                End Select
                Exit Do
             End If
            If Val(pagecardnum(目前數(6), 5)) = 2 And Val(pagecardnum(目前數(6), 6)) = 1 And Val(pagecardnum(目前數(6), 11)) = 1 Then
               電腦牌_模擬按牌 目前數(6)
               電腦出牌.Enabled = False
               Exit Do
            End If
        Loop
End If
End Sub


Private Sub 電腦出牌_手牌對齊_Timer()
For i = 1 To Val(pagecomglead)
   If 出牌順序統計暫時變數(4, i, 1) > 目前數(9) Then
       card(出牌順序統計暫時變數(4, i, 2)).Left = card(出牌順序統計暫時變數(4, i, 2)).Left + (240 / 10)
   End If
Next
目前數(8) = 目前數(8) + (240 / 10)
If 目前數(8) >= 240 Then
    電腦出牌_手牌對齊.Enabled = False
    Select Case 目前數(17)
        Case 1
            電腦出牌.Enabled = True
        Case 2
            '======結束動作
        Case 3
            If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 3 '(階段3)
            End If
'           If atkingck(56, 2) = 1 And atkingck(56, 1) <> 6 Then
'               atkingck(56, 1) = 5
'               技能.伊芙琳_怠惰的墓表 '(階段5)
'           ElseIf atkingck(56, 2) = 1 And atkingck(56, 1) = 6 Then
'               技能.伊芙琳_怠惰的墓表 '(階段6)
'           End If
        Case 4
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDestroyCards")) = 3 '(階段3)
            End If
'           If atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 10 Then
'               atkingck(59, 1) = 4
'               技能.伊芙琳_赤紅石榴  '(階段4)
'               Exit Sub
'           ElseIf atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
'               atkingck(59, 1) = 7
'               技能.伊芙琳_赤紅石榴  '(階段7)
'               Exit Sub
'           ElseIf atkingck(59, 2) = 1 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
'               atkingck(59, 1) = 10
'               技能.伊芙琳_赤紅石榴  '(階段10)
'               Exit Sub
'           End If
        Case 5
            If 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingGiveCards")) = 3 '(階段3)
            End If
'           If atkingck(61, 2) = 1 Then
'               atkingck(61, 1) = 5
'               技能.古魯瓦爾多_精神力吸收 '(階段5)
'           End If
        Case 6
           '===========事件卡執行_詛咒術_使用者(階段3)
            事件卡記錄暫時數(1, 3) = 3
            事件卡.詛咒術_使用者 0, 0
        Case 7
'            If turnatk = 1 And atkingck(72, 2) = 1 Then
'               atkingck(72, 1) = 4
'               技能.艾伯李斯特_雷擊  '(階段4)
'               Exit Sub
'            End If
'        Case 8
'            If turnatk = 1 And atkingck(122, 2) = 1 Then
'               atkingck(122, 1) = 5
'               技能.瑪格莉特_月光  '(階段5)
'               Exit Sub
'            End If
'        Case 9
'            If turnatk = 2 And atkingck(129, 2) = 1 Then
'               atkingck(129, 1) = 4
'               技能.庫勒尼西_瘋狂眼窩  '(階段4)
'               Exit Sub
'            End If
'        Case 10
'            If atkingck(156, 2) = 1 Then
'                atkingck(156, 1) = 3
'                技能.洛洛妮_貪婪之刃與嗜血之槍 '(階段3)
'            End If
'        Case 11
'            If atkingck(160, 2) = 1 Then
'                atkingck(160, 1) = 5
'                技能.克頓_惡意情報 '(階段5)
'            End If
'        Case 12
'           If atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 10 Then
'               atkingckai(108, 1) = 4
'               AI技能.伊芙琳_赤紅石榴  '(階段4)
'               Exit Sub
'           ElseIf atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
'               atkingckai(108, 1) = 7
'               AI技能.伊芙琳_赤紅石榴  '(階段7)
'               Exit Sub
'           ElseIf atkingckai(108, 2) = 1 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
'               atkingckai(108, 1) = 10
'               AI技能.伊芙琳_赤紅石榴  '(階段10)
'               Exit Sub
'           End If
    End Select
    
End If
End Sub


Private Sub 電腦出牌_出牌對齊_靠右_Timer()
For i = 1 To Val(pagecomqlead)
   If 出牌順序統計暫時變數(3, i, 1) < 目前數(9) Then
      card(出牌順序統計暫時變數(3, i, 2)).Left = card(出牌順序統計暫時變數(3, i, 2)).Left + (480 / 10)
   End If
   If 出牌順序統計暫時變數(3, i, 1) > 目前數(9) Then
      card(出牌順序統計暫時變數(3, i, 2)).Left = card(出牌順序統計暫時變數(3, i, 2)).Left - (500 / 10)
   End If
Next
目前數(7) = 目前數(7) + (480 / 10)
If 目前數(7) >= 480 Then
    電腦出牌_出牌對齊_靠右.Enabled = False
End If
End Sub

Private Sub 電腦出牌_出牌對齊_靠左_Timer()
For i = 1 To (pageqlead(2) - 1)
   card(出牌順序統計暫時變數(3, i, 2)).Left = card(出牌順序統計暫時變數(3, i, 2)).Left - (480 / 10)
Next
目前數(7) = 目前數(7) + (480 / 10)
If 目前數(7) >= 480 Then
    電腦出牌_出牌對齊_靠左.Enabled = False
    戰鬥系統類.出牌順序計算_電腦_手牌
    電腦出牌_手牌對齊.Enabled = True
End If
End Sub


Private Sub 電腦出牌_亮牌_Timer()
目前數(6) = 目前數(6) + 1
If 目前數(6) > pageqlead(2) Then
    電腦出牌_亮牌.Enabled = False
    Select Case turnatk
       Case 1
'          攻擊階段_階段2.Enabled = True
            執行動作_電腦方各階段出牌完畢後行動 1
       Case 2
'          cn32.Visible = True
'          bnok.Picture = LoadPicture(app_path & "gif\system\ok_1.jpg")
'          bnok.Visible = True
'          '==============
'            小人物頭像移動方向數(1) = 1
'            小人物頭像移動方向數(2) = 2
'            小人物頭像移動_使用者.Enabled = True
'            小人物頭像移動_電腦.Enabled = True
'          '==============
'          階段狀態數 = 1
'          FormMainMode.wmpse6.Controls.play
'          一般系統類.檢查音樂播放 6
'          戰鬥系統類.時間軸_重設
'          trtimeline.Enabled = True
'          '===========================
'            If Vss_EventPlayerAllActionOffNum(1) = 1 Then
'                For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
'                    FormMainMode.card(ckl).CardEnabledType = False
'                Next
'                目前數(24) = 47
'                等待時間_2.Enabled = True
'            ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
'                For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
'                    FormMainMode.card(ckl).CardEnabledType = False
'                Next
'                目前數(24) = 45
'                等待時間_2.Enabled = True
'            End If
            執行動作_電腦方各階段出牌完畢後行動 2
       Case 3
'          atkingtrtot.Interval = 600
'          atkingtrtot.Enabled = True
'          等待時間.Enabled = True
            目前數(24) = 30
            等待時間_2.Enabled = True
    End Select
    Exit Sub
End If
    card(出牌順序統計暫時變數(3, 目前數(6), 2)).Width = 810
    card(出牌順序統計暫時變數(3, 目前數(6), 2)).Height = 1260
    card(出牌順序統計暫時變數(3, 目前數(6), 2)).CardImage = app_path & "card\" & pagecardnum(出牌順序統計暫時變數(3, 目前數(6), 2), 8) & ".png"
    card(出牌順序統計暫時變數(3, 目前數(6), 2)).CardRotationType = pageonin(出牌順序統計暫時變數(3, 目前數(6), 2))
    FormMainMode.wmpse4.Controls.stop
    FormMainMode.wmpse4.Controls.play
    一般系統類.檢查音樂播放 4
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
'===========結束技能跳回執行點
'   If turnatk = 2 And atkingck(54, 2) = 1 And atkingck(54, 1) = 6 Then
'       技能.羅莎琳_黑霧幻影  '(階段6)
'       GoTo 技能_羅莎琳_黑霧幻影_跳入點
'   End If
'   If turnatk = 2 And atkingck(55, 2) = 1 And atkingck(55, 1) = 6 Then
'       技能.羅莎琳_EX_黑霧幻影  '(階段6)
'       GoTo 技能_羅莎琳_黑霧幻影_跳入點
'   End If
'   If turnatk = 1 And atkingckai(128, 2) = 1 And atkingckai(128, 1) = 6 Then
'       AI技能.羅莎琳_黑霧幻影  '(階段6)
'       GoTo AI技能_羅莎琳_黑霧幻影_跳入點
'   End If
'   If turnatk = 1 And atkingckai(129, 2) = 1 And atkingckai(129, 1) = 6 Then
'       AI技能.羅莎琳_EX_黑霧幻影  '(階段6)
'       GoTo AI技能_羅莎琳_黑霧幻影_跳入點
'   End If
'   If turnatk = 1 And atkingck(72, 2) = 1 And atkingck(72, 1) = 5 Then
'       技能.艾伯李斯特_雷擊  '(階段5)
'       GoTo 技能_艾伯李斯特_雷擊_跳入點
'   End If
'   If turnatk = 2 And atkingck(92, 2) = 1 And atkingck(92, 1) = 4 Then
'       技能.利恩_反擊的狼煙  '(階段4)
'       GoTo 技能_利恩_反擊的狼煙_跳入點
'   End If
'   If turnatk = 1 And atkingckai(74, 2) = 1 And atkingckai(74, 1) = 4 Then
'       AI技能.利恩_反擊的狼煙  '(階段4)
'       GoTo 技能_利恩_反擊的狼煙_跳入點
'   End If
'   If turnatk = 2 And atkingck(129, 2) = 1 And atkingck(129, 1) = 5 Then
'       技能.庫勒尼西_瘋狂眼窩  '(階段5)
'       GoTo 技能_庫勒尼西_瘋狂眼窩_跳入點
'   End If
'   If turnatk = 1 And atkingckai(79, 2) = 1 And atkingckai(79, 1) = 5 Then
'       AI技能.庫勒尼西_瘋狂眼窩  '(階段5)
'       GoTo 技能_庫勒尼西_瘋狂眼窩_跳入點
'   End If
'   If turnatk = 2 And atkingckai(66, 2) = 1 And atkingckai(66, 1) = 5 Then
'       AI技能.艾伯李斯特_雷擊  '(階段5)
'       GoTo 技能_艾伯李斯特_雷擊_跳入點
'   End If
''========================完成HP檢查
''  If 目前數(26) = 1 Then
''      GoTo HP檢查完畢_跳入點
''  End If
''=========================以下是技能檢查及啟動(擲多次骰子)
'    If turnatk = 2 And atkingck(94, 2) = 1 And atkingck(94, 1) = 3 Then
'       技能.夏洛特_大聖堂  '(階段3)
'       Exit Sub
'    ElseIf turnatk = 2 And atkingck(94, 2) = 1 And atkingck(94, 1) = 4 Then
'       技能.夏洛特_大聖堂  '(階段4)
'    End If
'    If turnatk = 1 And atkingckai(90, 2) = 1 And atkingckai(90, 1) = 3 Then
'       AI技能.夏洛特_大聖堂  '(階段3)
'       Exit Sub
'    ElseIf turnatk = 1 And atkingckai(90, 2) = 1 And atkingckai(90, 1) = 4 Then
'       AI技能.夏洛特_大聖堂  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(147, 2) = 1 And atkingck(147, 1) = 3 Then
'       技能.傑多_因果之幻  '(階段3)
'       Exit Sub
'    ElseIf turnatk = 1 And atkingck(147, 2) = 1 And atkingck(147, 1) = 4 Then
'       技能.傑多_因果之幻  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(48, 2) = 1 And atkingckai(48, 1) = 3 Then
'       AI技能.傑多_因果之幻  '(階段3)
'       Exit Sub
'    ElseIf turnatk = 2 And atkingckai(48, 2) = 1 And atkingckai(48, 1) = 4 Then
'       AI技能.傑多_因果之幻  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(159, 2) = 1 And atkingck(159, 1) = 3 Then
'       技能.克頓_隱蔽射擊  '(階段3)
'       Exit Sub
'    ElseIf turnatk = 1 And atkingck(159, 2) = 1 And atkingck(159, 1) = 4 Then
'       技能.克頓_隱蔽射擊  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(133, 2) = 1 And atkingckai(133, 1) = 3 Then
'       AI技能.克頓_隱蔽射擊  '(階段3)
'       Exit Sub
'    ElseIf turnatk = 2 And atkingckai(133, 2) = 1 And atkingckai(133, 1) = 4 Then
'       AI技能.克頓_隱蔽射擊  '(階段4)
'    End If
''============以下是技能檢查及啟動
'    '=============================(梅倫-Lowball/Gamble)
'    If turnatk = 1 And atkingck(65, 2) = 1 Then
'       atkingck(65, 1) = 3
'       技能.梅倫_Lowball  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(65, 2) = 1 Then
'       atkingckai(65, 1) = 3
'       AI技能.梅倫_Lowball  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(66, 2) = 1 Then
'       atkingck(66, 1) = 3
'       技能.梅倫_Gamble  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(30, 2) = 1 Then
'       atkingckai(30, 1) = 3
'       AI技能.梅倫_Gamble  '(階段3)
'    End If
'    '=============================(普通)
'    If turnatk = 1 And atkingck(25, 2) = 1 Then
'       atkingck(25, 1) = 4
'       技能.史塔夏_命運的鐵門  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(23, 2) = 1 Then
'        技能.史塔夏_愚者之手 '(階段3)
'    End If
'    If turnatk = 1 And atkingck(35, 2) = 1 Then
'       atkingck(35, 1) = 3
'       技能.CC_高頻電磁手術刀  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(50, 2) = 1 Then
'       atkingckai(50, 1) = 3
'       AI技能.CC_高頻電磁手術刀  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(58, 2) = 1 Then
'       atkingck(58, 1) = 4
'       技能.伊芙琳_紅蓮車輪  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(51, 2) = 1 Then
'       atkingckai(51, 1) = 4
'       AI技能.伊芙琳_紅蓮車輪  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(20, 2) = 1 Then
'        AI技能.史塔夏_愚者之手 '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(21, 2) = 1 Then
'       atkingckai(21, 1) = 4
'       AI技能.史塔夏_命運的鐵門  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(98, 2) = 1 Then
'       atkingck(98, 1) = 3
'       技能.露緹亞_腐朽之靈  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(95, 2) = 1 Then
'       atkingckai(95, 1) = 3
'       AI技能.露緹亞_腐朽之靈  '(階段3)
'    End If
'    '=======================(追加攻擊骰數類)
'    If turnatk = 1 And atkingck(17, 2) = 1 Then
'       技能.帕茉_靜謐之背  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(36, 2) = 1 Then
'       atkingckai(36, 1) = 3
'       AI技能.帕茉_靜謐之背  '(階段3)
'    End If
'    '=======================(減輕攻擊骰數類)
'    If turnatk = 1 And atkingckai(18, 2) = 1 Then
'       AI技能.吸血姬蕾米雅_消失  '(階段3)
'    End If
'    If turnatk = 2 And atkingck(38, 2) = 1 Then
'        atkingck(38, 1) = 4
'        技能.蕾_EX_協奏曲_加百烈的守護  '(階段4)
'    End If
'    If turnatk = 1 And atkingckai(58, 2) = 1 Then
'        atkingckai(58, 1) = 4
'        AI技能.蕾_EX_協奏曲_加百烈的守護  '(階段4)
'    End If
'    If turnatk = 2 And atkingck(102, 2) = 1 Then
'       atkingck(102, 1) = 4
'       技能.艾蕾可_王座之炎  '(階段4)
'    End If
'    If turnatk = 1 And atkingckai(91, 2) = 1 Then
'       atkingckai(91, 1) = 4
'       AI技能.艾蕾可_王座之炎  '(階段4)
'    End If
'    '===============(骰數其他類)
'    If turnatk = 2 And atkingck(137, 2) = 1 Then
'       atkingck(137, 1) = 4
'       技能.蕾格烈芙_LAR  '(階段4)
'    End If
'    If turnatk = 1 And atkingckai(47, 2) = 1 Then
'       atkingckai(47, 1) = 4
'       AI技能.蕾格烈芙_LAR  '(階段4)
'    End If
'    If turnatk = 2 And atkingck(117, 2) = 1 Then
'       atkingck(117, 1) = 3
'       技能.泰瑞爾_Von_541  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(76, 2) = 1 Then
'       atkingckai(76, 1) = 3
'       AI技能.泰瑞爾_Von_541  '(階段3)
'    End If
'    If turnatk = 2 And atkingck(103, 2) = 1 Then
'       atkingck(103, 1) = 3
'       技能.艾蕾可_白百合  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(92, 2) = 1 Then
'       atkingckai(92, 1) = 3
'       AI技能.艾蕾可_白百合  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(104, 2) = 1 Then
'       atkingck(104, 1) = 4
'       技能.艾蕾可_聖王威光  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(93, 2) = 1 Then
'       atkingckai(93, 1) = 4
'       AI技能.艾蕾可_聖王威光  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(106, 2) = 1 Then
'       atkingck(106, 1) = 5
'       技能.梅莉_夢幻魔杖  '(階段5)
'    End If
'    If turnatk = 2 And atkingckai(99, 2) = 1 Then
'       atkingckai(99, 1) = 5
'       AI技能.梅莉_夢幻魔杖  '(階段5)
'    End If
'    '=================(防禦成功類)
'    If turnatk = 2 And atkingck(123, 2) = 1 Then
'       atkingck(123, 1) = 3
'       技能.瑪格莉特_恍惚 '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(42, 2) = 1 Then
'       atkingckai(42, 1) = 3
'       AI技能.瑪格莉特_恍惚 '(階段3)
'    End If
'    If turnatk = 2 And atkingck(47, 2) = 1 Then
'       atkingck(47, 1) = 3
'       技能.尤莉卡_不善的信仰  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(137, 2) = 1 Then
'       atkingckai(137, 1) = 3
'       AI技能.尤莉卡_不善的信仰  '(階段3)
'    End If
'    If turnatk = 2 And atkingck(54, 2) = 1 Then
'       atkingck(54, 1) = 4
'       技能.羅莎琳_黑霧幻影  '(階段4)
'       Exit Sub
'    End If
'    If turnatk = 2 And atkingck(55, 2) = 1 Then
'       atkingck(55, 1) = 4
'       技能.羅莎琳_EX_黑霧幻影  '(階段4)
'       Exit Sub
'    End If
'    If turnatk = 1 And atkingckai(128, 2) = 1 Then
'       atkingckai(128, 1) = 4
'       AI技能.羅莎琳_黑霧幻影 '(階段4)
'       Exit Sub
'    End If
'    If turnatk = 1 And atkingckai(129, 2) = 1 Then
'       atkingckai(129, 1) = 4
'       AI技能.羅莎琳_EX_黑霧幻影 '(階段4)
'       Exit Sub
'    End If
'    '=====================
'技能_羅莎琳_黑霧幻影_跳入點: '技能-羅莎琳-(普、Ex)-黑霧幻影 結束執行繼續點
'AI技能_羅莎琳_黑霧幻影_跳入點: '技能-AI-羅莎琳-(普、Ex)-黑霧幻影 結束執行繼續點
'    '=======================(攻擊成功類)
'    If turnatk = 2 And atkingckai(7, 2) = 1 Then
'       AI技能.南瓜王_重壓  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(16, 2) = 1 Then
'       AI技能.吸血姬蕾米雅_吸血  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(51, 2) = 1 Then
'       atkingck(51, 1) = 4
'       技能.羅莎琳_染血之刃  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(50, 2) = 1 Then
'       atkingck(50, 1) = 4
'       技能.羅莎琳_EX_染血之刃  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(52, 2) = 1 Then
'       atkingck(52, 1) = 3
'       技能.羅莎琳_黑霧的纏繞  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(90, 2) = 1 Then
'       atkingck(90, 1) = 3
'       技能.利恩_劫影攻擊  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(91, 2) = 1 Then
'       atkingck(91, 1) = 3
'       技能.利恩_毒牙  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(72, 2) = 1 Then
'       atkingckai(72, 1) = 3
'       AI技能.利恩_劫影攻擊  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(73, 2) = 1 Then
'       atkingckai(73, 1) = 3
'       AI技能.利恩_毒牙  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(140, 2) = 1 Then
'       atkingck(140, 1) = 3
'       技能.多妮妲_殘虐傾向  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(143, 2) = 1 Then
'       atkingck(143, 1) = 3
'       技能.多妮妲_律死擊  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(151, 2) = 1 Then
'       atkingck(151, 1) = 3
'       技能.阿奇波爾多_劫影攻擊  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(155, 2) = 1 Then
'       atkingck(155, 1) = 3
'       技能.洛洛妮_砲擊壓制  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(112, 2) = 1 Then
'       atkingck(112, 1) = 3
'       技能.貝琳達_裂地冰牙  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(124, 2) = 1 Then
'       atkingckai(124, 1) = 3
'       AI技能.貝琳達_裂地冰牙  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(86, 2) = 1 Then
'       atkingckai(86, 1) = 3
'       AI技能.洛洛妮_砲擊壓制  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(53, 2) = 1 Then
'       atkingckai(53, 1) = 3
'       AI技能.多妮妲_殘虐傾向  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(52, 2) = 1 Then
'       atkingckai(52, 1) = 3
'       AI技能.多妮妲_律死擊  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(59, 2) = 1 Then
'       atkingckai(59, 1) = 3
'       AI技能.羅莎琳_黑霧的纏繞  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(32, 2) = 1 Then
'       atkingckai(32, 1) = 4
'       AI技能.羅莎琳_染血之刃  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(140, 2) = 1 Then
'       atkingckai(140, 1) = 4
'       AI技能.羅莎琳_EX_染血之刃  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(84, 2) = 1 Then
'       atkingckai(84, 1) = 3
'       AI技能.阿奇波爾多_劫影攻擊  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(46, 2) = 1 Then
'       atkingck(46, 1) = 4
'       技能.尤莉卡_奸佞的鐵鎚  '(階段4)
'    End If
'    If turnatk = 2 And atkingckai(136, 2) = 1 Then
'       atkingckai(136, 1) = 4
'       AI技能.尤莉卡_奸佞的鐵鎚  '(階段4)
'    End If
'    '====================以下是異常狀態檢查及啟動(狂戰士、恐怖)
'    異常狀態檢查數(27, 1) = 1
'    異常狀態.狂戰士_使用者  '(階段1)
'    '=============
'    異常狀態檢查數(28, 1) = 1
'    異常狀態.狂戰士_電腦  '(階段1)
'    '=============
'    異常狀態檢查數(29, 1) = 1
'    異常狀態.恐怖_使用者  '(階段1)
'    '=============
'    異常狀態檢查數(30, 1) = 1
'    異常狀態.恐怖_電腦   '(階段1)
'    '=======================(防禦骰數相關類)
'    If turnatk = 2 And atkingck(60, 2) = 1 Then
'       atkingck(60, 1) = 4
'       技能.古魯瓦爾多_血之恩賜  '(階段4)
'    End If
'    If turnatk = 1 And atkingckai(62, 2) = 1 Then
'       atkingckai(62, 1) = 4
'       AI技能.古魯瓦爾多_血之恩賜  '(階段4)
'    End If
'    If turnatk = 2 And atkingck(73, 2) = 1 Then
'       atkingck(73, 1) = 4
'       技能.艾伯李斯特_茨林  '(階段4)
'    End If
'    If turnatk = 1 And atkingckai(67, 2) = 1 Then
'       atkingckai(67, 1) = 4
'       AI技能.艾伯李斯特_茨林  '(階段4)
'    End If
''=============
''HP檢查完畢_跳入點:
''==========================
'技能_艾伯李斯特_雷擊_跳入點: '(一般/AI)技能-艾伯李斯特-雷擊 結束執行繼續點
'技能_利恩_反擊的狼煙_跳入點: '(一般/AI)技能-利恩-反擊的狼煙 結束執行繼續點
'技能_庫勒尼西_瘋狂眼窩_跳入點: '(一般/AI)技能-庫勒尼西-瘋狂眼窩 結束執行繼續點
''=======================
''If 目前數(26) = 0 Then
''    HP檢查階段數 = 5
''    戰鬥系統類.雙方HP檢查
''    Exit Sub
''End If
''=============以下是技能檢查及啟動(回牌及抽牌類)
'    If turnatk = 2 And atkingck(92, 2) = 1 Then
'       atkingck(92, 1) = 3
'       技能.利恩_反擊的狼煙  '(階段3)
'       Exit Sub
'    End If
'    If turnatk = 1 And atkingckai(74, 2) = 1 Then
'       atkingckai(74, 1) = 3
'       AI技能.利恩_反擊的狼煙  '(階段3)
'       Exit Sub
'    End If
'    '===============以下是異常狀態檢查及啟動(骰數傷害歸0)
'     If turnatk = 2 Then
'        異常狀態檢查數(14, 1) = 1
'        異常狀態.不死_使用者 '(階段1)
'    End If
'    '=================
'    If turnatk = 1 Then
'        異常狀態檢查數(18, 1) = 1
'        異常狀態.不死_電腦 '(階段1)
'    End If
'    '========================(丟棄牌類)
'    If turnatk = 1 And atkingck(72, 2) = 1 Then
'       atkingck(72, 1) = 3
'       技能.艾伯李斯特_雷擊  '(階段3)
'       Exit Sub
'    End If
'    If turnatk = 2 And atkingckai(66, 2) = 1 Then
'       atkingckai(66, 1) = 3
'       AI技能.艾伯李斯特_雷擊  '(階段3)
'       Exit Sub
'    End If
'    If turnatk = 2 And atkingck(129, 2) = 1 Then
'       atkingck(129, 1) = 3
'       技能.庫勒尼西_瘋狂眼窩  '(階段3)
'       Exit Sub
'    End If
'    If turnatk = 1 And atkingckai(79, 2) = 1 Then
'       atkingckai(79, 1) = 3
'       AI技能.庫勒尼西_瘋狂眼窩  '(階段3)
'       Exit Sub
'    End If
''=============================(傷害骰數轉移類)
'    If turnatk = 2 And atkingckai(11, 2) = 1 Then
'       atkingckai(11, 1) = 4
'       AI技能.蕾_終曲_無盡輪迴的終結  '(階段4)
'    End If
'    If turnatk = 1 And atkingck(15, 2) = 1 Then
'       atkingck(15, 1) = 3
'       技能.蕾_終曲_無盡輪迴的終結  '(階段3)
'    End If
'    If turnatk = 1 And atkingck(161, 2) = 1 Then
'       atkingck(161, 1) = 3
'       技能.蕾_EX_終曲_無盡輪迴的終結  '(階段3)
'    End If
'    If turnatk = 2 And atkingckai(127, 2) = 1 Then
'       atkingckai(127, 1) = 3
'       AI技能.蕾_EX_終曲_無盡輪迴的終結  '(階段3)
'    End If
'    If turnatk = 2 And atkingck(32, 2) = 1 Then
'       atkingck(32, 1) = 3
'       技能.艾茵_兩個身體  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(38, 2) = 1 Then
'       atkingckai(38, 1) = 3
'       AI技能.艾茵_兩個身體  '(階段3)
'    End If
'    If turnatk = 2 And atkingck(158, 2) = 1 Then
'       atkingck(158, 1) = 3
'       技能.克頓_逃亡計畫  '(階段3)
'    End If
'    If turnatk = 1 And atkingckai(132, 2) = 1 Then
'       atkingckai(132, 1) = 3
'       AI技能.克頓_逃亡計畫  '(階段3)
'    End If
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
'測試表單.Visible = True
測試表單.smallleftus.Caption = personusminijpg.小人物影子Left
測試表單.smalltopus.Caption = personusminijpg.小人物影子top差
測試表單.smallleftcom.Caption = personcomminijpg.小人物影子Left
測試表單.smalltopcom.Caption = personcomminijpg.小人物影子top差
測試表單.smallpnleftus.Caption = personusminijpg.Left
測試表單.smallpntopus.Caption = personusminijpg.Top
測試表單.smallpnleftcom.Caption = personcomminijpg.Left
測試表單.smallpntopcom.Caption = personcomminijpg.Top
測試表單.personfus.Caption = 顯示列1.使用者方小人物圖片left
測試表單.personfcom.Caption = 顯示列1.電腦方小人物圖片left
If Formsetting.checktest.Value = 1 Then
    測試表單.Height = 6825
ElseIf Formsetting.checktestpersondown.Value = 1 Then
    測試表單.Height = 3075
End If
'=================
戰鬥系統類.時間軸_停止
'=================
測試表單.Show 1
End Sub
Private Sub bnabout_Click()
Form2.Show 1
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
End Sub

Private Sub bnconfig_Click()
Formsetting.Left = FormMainMode.Left + 915
Formsetting.Top = FormMainMode.Top + 300
一般系統類.自由戰鬥模式設定表單各式設定讀入程序
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
Formsetting.Show 1
End Sub


Private Sub bnstart_Click()
PEGameFreeModeSettingForm.Enabled = False
一般系統類.開始遊戲進行程序
End Sub



Private Sub Form_Load()
'============
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'==============
一般系統類.判斷字型_FormMainMode
一般系統類.主選單_PEStartForm顯示
End Sub
Private Sub personreadifus_Click()
cdgpersonus.ShowOpen
Formgamesetting.Visible = True
人物系統類.卡片人物資訊讀入_初階段 cdgpersonus.filename
End Sub

Private Sub personsettingcom_Click(Index As Integer)
Select Case Index
   Case 1
        formsettingpersoncom.Left = FormMainMode.Left + 2205
        formsettingpersoncom.Top = FormMainMode.Top + 2000
        formsettingpersoncom.Height = 6915
        formsettingpersoncom.Width = 7065
        formsettingpersoncom.Visible = True
   Case 2
        formsettingpersoncom2.Left = FormMainMode.Left + 2205
        formsettingpersoncom2.Top = FormMainMode.Top + 2000
        formsettingpersoncom2.Height = 6915
        formsettingpersoncom2.Width = 7065
        formsettingpersoncom2.Visible = True
  Case 3
        formsettingpersoncom3.Left = FormMainMode.Left + 2205
        formsettingpersoncom3.Top = FormMainMode.Top + 2000
        formsettingpersoncom3.Height = 6915
        formsettingpersoncom3.Width = 7065
        formsettingpersoncom3.Visible = True
End Select
End Sub

Private Sub personsettingus_Click(Index As Integer)
Select Case Index
   Case 1
        formsettingpersonus.Left = Me.Left + 2205
        formsettingpersonus.Top = Me.Top + 2000
        formsettingpersonus.Height = 6915
        formsettingpersonus.Width = 7065
        formsettingpersonus.Visible = True
   Case 2
        formsettingpersonus2.Left = Me.Left + 2205
        formsettingpersonus2.Top = Me.Top + 2000
        formsettingpersonus2.Height = 6915
        formsettingpersonus2.Width = 7065
        formsettingpersonus2.Visible = True
  Case 3
        formsettingpersonus3.Left = Me.Left + 2205
        formsettingpersonus3.Top = Me.Top + 2000
        formsettingpersonus3.Height = 6915
        formsettingpersonus3.Width = 7065
        formsettingpersonus3.Visible = True
End Select
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
'MsgBox formmainmode.personnameus(index).ListIndex
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
'personlevelcom(Index).ListIndex = -1
personlevelcom(Index).Clear
End Sub

Private Sub personresetus_Click(Index As Integer)
personnameus(Index).ListIndex = -1
'personlevelus(Index).ListIndex = -1
personlevelus(Index).Clear
End Sub
Private Sub start1_Timer()
If st > 200 Then
   stup.Enabled = True
   stdown.Enabled = True
   start1.Enabled = False
'   cardustr.Enabled = True
'   cardcomtr.Enabled = True
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
'ElseIf st = 150 Then
'   FormMainMode.wmp.Controls.play
'   st = Val(st) + 1
Else
  st = Val(st) + 1
End If
End Sub

Private Sub start2_Timer()
If sq = 401 Then
'   stup.Enabled = True
'   stdown.Enabled = True
'   start2.Enabled = False
   tr大人物形像_使用者.Enabled = True
   tr大人物形像_電腦.Enabled = True
'   cardustr.Enabled = True
'   cardcomtr.Enabled = True
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
'If 2580 - bigw < 0 Or formsettingpersonus.atkingjpgleftallzero.Value = 1 Then
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
'       If Abs(大人物形像_使用者.Left - bigall) < kp And 大人物形像_使用者.Left >= 0 Then
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
'If kn > 5460 Then
'   kn = 5460
'End If
Dim bigwn, bigall As Integer
bigwn = (大人物形像_電腦.大人物圖片width / 2)
If 8760 - bigwn > Val(FormMainMode.ScaleWidth) - Val(大人物形像_電腦.大人物圖片width) Or Val(VBEPerson(2, 1, 2, 2, 5)) = 1 Then
    bigall = Val(FormMainMode.ScaleWidth) - Val(大人物形像_電腦.大人物圖片width)
Else
    bigall = 8760 - bigwn
End If
'kr = (大人物形像_電腦.大人物圖片width / 28)
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
