VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form Formsetting 
   Appearance      =   0  '平面
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE-進階設定"
   ClientHeight    =   8730
   ClientLeft      =   6330
   ClientTop       =   2535
   ClientWidth     =   17670
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formsetting.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   17670
   Begin VB.Frame 表單_系統 
      Caption         =   "系統"
      Height          =   5895
      Left            =   9480
      TabIndex        =   135
      Top             =   2160
      Width           =   9015
      Begin VB.ComboBox cbsimilarlevel 
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   150
         Text            =   "Combo1"
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkusesimilarlevel 
         Caption         =   "人物角色隨機時以相似之等級進行對戰："
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   149
         Top             =   2880
         Width           =   6495
      End
      Begin VB.TextBox 自訂AI手牌張數 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   300
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   146
         Text            =   "7"
         Top             =   5040
         Width           =   375
      End
      Begin VB.CheckBox chksetcomaipagenum 
         Caption         =   "自訂AI計算之手牌張數：        張"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   145
         Top             =   5040
         Width           =   6495
      End
      Begin VB.CheckBox chkusenewinterface 
         Caption         =   "使用新式階段顯示介面"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   144
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkusenewaipersonauto 
         Caption         =   "使用者方使用「智能判斷型人工智慧AI」進行對戰"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   141
         Top             =   2400
         Width           =   6495
      End
      Begin VB.CheckBox checktestpersondown 
         Caption         =   "測試模式(開啟自訂影子座標)"
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
         TabIndex        =   138
         Top             =   1920
         Width           =   6135
      End
      Begin VB.CheckBox chkusenewpage 
         Caption         =   "使用新式隨場景變化之行動卡牌組"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   137
         Top             =   960
         Width           =   6495
      End
      Begin VB.CheckBox chkusenewai 
         Caption         =   "電腦方使用「智能判斷型人工智慧AI」進行對戰"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   136
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label8 
         Caption         =   "以此選擇正負2等級為隨機之上下限，若3次隨機角色都不符合範圍者則以第3次結果+隨機等級為主。"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   151
         Top             =   3240
         Width           =   8175
      End
      Begin VB.Label Label7 
         Caption         =   "請注意：若將計算張數設定過高可能會導致程式沒有回應。"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   147
         Top             =   5400
         Width           =   5655
      End
   End
   Begin VB.Frame 事件卡_電腦 
      Caption         =   "事件卡編輯(電腦方)"
      Height          =   5895
      Left            =   9480
      TabIndex        =   51
      Top             =   2040
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CheckBox persontgrecom 
         Caption         =   "遵守Unlight事件卡規則"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   108
         Top             =   360
         Value           =   1  '核取
         Width           =   2535
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         ScaleHeight     =   615
         ScaleWidth      =   3615
         TabIndex        =   103
         Top             =   240
         Width           =   3615
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "無"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "自訂"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "選擇最大值"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1080
            TabIndex        =   105
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "隨機"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   1080
            TabIndex        =   104
            Top             =   220
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   8535
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   88
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   87
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   86
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   6000
            TabIndex        =   85
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   6000
            TabIndex        =   83
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   6000
            TabIndex        =   84
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label persontgcom 
            Caption         =   "1:"
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   96
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "2:"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   95
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "3:"
            Height          =   375
            Index           =   3
            Left            =   2040
            TabIndex        =   94
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "4:"
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   93
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "5:"
            Height          =   375
            Index           =   5
            Left            =   5640
            TabIndex        =   92
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "6:"
            Height          =   375
            Index           =   6
            Left            =   5640
            TabIndex        =   91
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label personlevelcom 
            BackStyle       =   0  '透明
            Caption         =   "LV"
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
            TabIndex        =   90
            Top             =   720
            Width           =   495
         End
         Begin VB.Label personnamecom 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   89
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         Width           =   8535
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   6000
            TabIndex        =   73
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   6000
            TabIndex        =   72
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   10
            Left            =   6000
            TabIndex        =   71
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   2400
            TabIndex        =   70
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   69
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   68
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label personnamecom 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   600
            TabIndex        =   81
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label personlevelcom 
            BackStyle       =   0  '透明
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   80
            Top             =   720
            Width           =   495
         End
         Begin VB.Label persontgcom 
            Caption         =   "6:"
            Height          =   375
            Index           =   12
            Left            =   5640
            TabIndex        =   79
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "5:"
            Height          =   375
            Index           =   11
            Left            =   5640
            TabIndex        =   78
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "4:"
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   77
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "3:"
            Height          =   375
            Index           =   9
            Left            =   2040
            TabIndex        =   76
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "2:"
            Height          =   375
            Index           =   8
            Left            =   2040
            TabIndex        =   75
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "1:"
            Height          =   375
            Index           =   7
            Left            =   2040
            TabIndex        =   74
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   52
         Top             =   4080
         Width           =   8535
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   18
            Left            =   6000
            TabIndex        =   58
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   17
            Left            =   6000
            TabIndex        =   57
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   16
            Left            =   6000
            TabIndex        =   56
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   55
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   54
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   53
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label personnamecom 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   66
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label personlevelcom 
            BackStyle       =   0  '透明
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   495
         End
         Begin VB.Label persontgcom 
            Caption         =   "6:"
            Height          =   375
            Index           =   18
            Left            =   5640
            TabIndex        =   64
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "5:"
            Height          =   375
            Index           =   17
            Left            =   5640
            TabIndex        =   63
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "4:"
            Height          =   375
            Index           =   16
            Left            =   5640
            TabIndex        =   62
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "3:"
            Height          =   375
            Index           =   15
            Left            =   2040
            TabIndex        =   61
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "2:"
            Height          =   375
            Index           =   14
            Left            =   2040
            TabIndex        =   60
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "1:"
            Height          =   375
            Index           =   13
            Left            =   2040
            TabIndex        =   59
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label personwagcom 
         Caption         =   "(角色有隨機時無法自訂)"
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
         Left            =   3000
         TabIndex        =   109
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame 一般設定_表單 
      Caption         =   "一般設定"
      Height          =   6015
      Left            =   120
      TabIndex        =   111
      Top             =   2040
      Width           =   9015
      Begin VB.Frame 其他設定 
         Caption         =   "其他設定"
         Height          =   1215
         Left            =   120
         TabIndex        =   129
         Top             =   4560
         Width           =   8775
         Begin VB.TextBox 大亂鬥模式選項_牌數 
            Appearance      =   0  '平面
            BorderStyle     =   0  '沒有框線
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   148
            Text            =   "17"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox 挑戰模式選項_牌數 
            Appearance      =   0  '平面
            BorderStyle     =   0  '沒有框線
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   132
            Text            =   "4"
            Top             =   360
            Width           =   375
         End
         Begin VB.CheckBox 挑戰模式選項 
            Caption         =   "挑戰模式（對戰對手多發        張牌）(Max:30)"
            Height          =   300
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   5655
         End
         Begin VB.CheckBox checktest 
            Caption         =   "測試模式(Admin)"
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
            Left            =   5880
            TabIndex        =   131
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox 大亂鬥選項 
            Caption         =   "大亂鬥模式（雙方角色發        張牌，HP=99） (Min:1)"
            Height          =   375
            Left            =   240
            TabIndex        =   130
            Top             =   720
            Width           =   6255
         End
      End
      Begin VB.Frame 其他 
         Caption         =   "其他"
         Height          =   1215
         Left            =   120
         TabIndex        =   125
         Top             =   3240
         Width           =   8775
         Begin VB.TextBox ckendturnnum 
            Height          =   420
            Left            =   1080
            TabIndex        =   126
            Text            =   "18"
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox chkpersonvsmode 
            Caption         =   "仿對戰模式"
            Height          =   300
            Left            =   240
            TabIndex        =   128
            Top             =   280
            Width           =   1575
         End
         Begin VB.CheckBox ckendturn 
            Caption         =   "對戰           回合"
            Height          =   375
            Left            =   240
            TabIndex        =   127
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Frame 背景圖片 
         Caption         =   "背景圖片及音樂"
         Height          =   2895
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   8775
         Begin VB.CommandButton Command2 
            Caption         =   "從檔案..."
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
            Left            =   4080
            TabIndex        =   139
            Top             =   1920
            Width           =   1095
         End
         Begin VB.ComboBox BGM選擇 
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
            Left            =   6120
            TabIndex        =   117
            Text            =   "Combo1"
            Top             =   2280
            Width           =   2535
         End
         Begin VB.ComboBox 對戰地圖選擇 
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2880
            TabIndex        =   116
            Text            =   "Combo2"
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "從檔案..."
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
            Left            =   7560
            TabIndex        =   115
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox ckbgmmute 
            Caption         =   "靜音"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7920
            TabIndex        =   114
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox cksemute 
            Caption         =   "靜音"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7920
            TabIndex        =   113
            Top             =   960
            Width           =   735
         End
         Begin MSComDlg.CommonDialog cdgBGMchooce 
            Left            =   7080
            Top             =   1800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "UnlightVBE-BGM選擇-開啟檔案"
         End
         Begin MSComDlg.CommonDialog cdgMAPchooce 
            Left            =   3600
            Top             =   1800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "UnlightVBE-對戰地圖選擇-開啟檔案"
         End
         Begin ComctlLib.Slider sdrbgm 
            Height          =   495
            Left            =   4680
            TabIndex        =   142
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            _Version        =   327682
            LargeChange     =   10
            Min             =   1
            Max             =   100
            SelStart        =   50
            Value           =   50
         End
         Begin ComctlLib.Slider sdrse 
            Height          =   495
            Left            =   4680
            TabIndex        =   143
            Top             =   840
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            _Version        =   327682
            LargeChange     =   10
            Min             =   1
            Max             =   100
            SelStart        =   50
            Value           =   45
         End
         Begin ComctlLib.ImageList ImageListback 
            Left            =   3000
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin VB.Label lopnmapjpgtext 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2880
            TabIndex        =   140
            Top             =   1800
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label Label4 
            Caption         =   "BGM："
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
            Left            =   5400
            TabIndex        =   124
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "BGM音量："
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
            Left            =   3720
            TabIndex        =   123
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "SE音量："
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
            Left            =   3960
            TabIndex        =   122
            Top             =   960
            Width           =   855
         End
         Begin VB.Label bgmve 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "50"
            Height          =   375
            Left            =   7320
            TabIndex        =   121
            Top             =   480
            Width           =   495
         End
         Begin VB.Label seve 
            Alignment       =   1  '靠右對齊
            BackStyle       =   0  '透明
            Caption         =   "45"
            Height          =   375
            Left            =   7320
            TabIndex        =   120
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lopnmusictext 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5400
            TabIndex        =   119
            Top             =   1440
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Image 場地顯示 
            Height          =   2415
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label randomtext 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   48
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   960
            TabIndex        =   118
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Shape randombk 
            BackColor       =   &H00000000&
            BackStyle       =   1  '不透明
            BorderStyle     =   0  '透明
            Height          =   2415
            Left            =   240
            Top             =   360
            Width           =   2535
         End
      End
   End
   Begin TabDlg.SSTab t1 
      Height          =   6615
      Left            =   0
      TabIndex        =   134
      Top             =   1560
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   697
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微軟正黑體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "Formsetting.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "事件卡（使用者）"
      TabPicture(1)   =   "Formsetting.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "事件卡（電腦）"
      TabPicture(2)   =   "Formsetting.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "系統"
      TabPicture(3)   =   "Formsetting.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin VB.Frame 事件卡_使用者 
      Caption         =   "事件卡編輯(使用者方)"
      Height          =   5895
      Left            =   9600
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox persontgrenus 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6000
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   98
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton persontgruonus 
            Caption         =   "隨機"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1080
            TabIndex        =   102
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton persontgruonus 
            Caption         =   "選擇最大值"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1080
            TabIndex        =   101
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton persontgruonus 
            Caption         =   "自訂"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton persontgruonus 
            Caption         =   "無"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   99
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CheckBox persontgreus 
         Caption         =   "遵守Unlight事件卡規則"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   97
         Top             =   360
         Value           =   1  '核取
         Width           =   2535
      End
      Begin VB.Frame f3 
         Height          =   1695
         Left            =   120
         TabIndex        =   36
         Top             =   4080
         Width           =   8535
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   42
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   41
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   40
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   16
            Left            =   6000
            TabIndex        =   39
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   17
            Left            =   6000
            TabIndex        =   38
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   18
            Left            =   6000
            TabIndex        =   37
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label persontgus 
            Caption         =   "1:"
            Height          =   375
            Index           =   13
            Left            =   2040
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "2:"
            Height          =   375
            Index           =   14
            Left            =   2040
            TabIndex        =   49
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "3:"
            Height          =   375
            Index           =   15
            Left            =   2040
            TabIndex        =   48
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "4:"
            Height          =   375
            Index           =   16
            Left            =   5640
            TabIndex        =   47
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "5:"
            Height          =   375
            Index           =   17
            Left            =   5640
            TabIndex        =   46
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "6:"
            Height          =   375
            Index           =   18
            Left            =   5640
            TabIndex        =   45
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label personlevelus 
            BackStyle       =   0  '透明
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   495
         End
         Begin VB.Label personnameus 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   43
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame f2 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   8535
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   26
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   2400
            TabIndex        =   25
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   10
            Left            =   6000
            TabIndex        =   24
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   6000
            TabIndex        =   23
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   6000
            TabIndex        =   22
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label persontgus 
            Caption         =   "1:"
            Height          =   375
            Index           =   7
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "2:"
            Height          =   375
            Index           =   8
            Left            =   2040
            TabIndex        =   34
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "3:"
            Height          =   375
            Index           =   9
            Left            =   2040
            TabIndex        =   33
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "4:"
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "5:"
            Height          =   375
            Index           =   11
            Left            =   5640
            TabIndex        =   31
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "6:"
            Height          =   375
            Index           =   12
            Left            =   5640
            TabIndex        =   30
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label personlevelus 
            BackStyle       =   0  '透明
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   495
         End
         Begin VB.Label personnameus 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   600
            TabIndex        =   28
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame f1 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   8535
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   6000
            TabIndex        =   13
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   6000
            TabIndex        =   12
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   6000
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   10
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   8
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label personnameus 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "微軟正黑體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
         Begin VB.Label personlevelus 
            BackStyle       =   0  '透明
            Caption         =   "LV"
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
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.Label persontgus 
            Caption         =   "6:"
            Height          =   375
            Index           =   6
            Left            =   5640
            TabIndex        =   18
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "5:"
            Height          =   375
            Index           =   5
            Left            =   5640
            TabIndex        =   17
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "4:"
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "3:"
            Height          =   375
            Index           =   3
            Left            =   2040
            TabIndex        =   15
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "2:"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   14
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "1:"
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label personwagus 
         Caption         =   "(角色有隨機時無法自訂)"
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
         Left            =   3840
         TabIndex        =   110
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9255
      TabIndex        =   3
      Top             =   8040
      Width           =   9255
      Begin VB.Label Label3 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "OK"
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
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3480
         Picture         =   "Formsetting.frx":0D3A
         Top             =   120
         Width           =   2145
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '不透明
         Height          =   525
         Left            =   -120
         Top             =   120
         Width           =   9375
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      Picture         =   "Formsetting.frx":18BD
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "大小姐，這裡能夠讓您做更詳細的設定歐"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "還需要什麼嗎?"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '不透明
      Height          =   1845
      Left            =   0
      Top             =   -240
      Width           =   9255
   End
End
Attribute VB_Name = "Formsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub checktest_Click()
If checktest.Value = 1 Then
    persontgreus.Value = 0
    persontgrecom.Value = 0
    persontgruoncom(3).Value = True
    persontgruonus(3).Value = True
    Formsetting.對戰地圖選擇.ListIndex = Formsetting.對戰地圖選擇.ListCount - 1
Else
    persontgreus.Value = 1
    persontgrecom.Value = 1
    persontgruoncom(1).Value = True
    persontgruonus(1).Value = True
    Formsetting.對戰地圖選擇.ListIndex = 0
End If
End Sub

Private Sub chkusesimilarlevel_Click()
If chkusesimilarlevel.Value = 0 Then
   cbsimilarlevel.Enabled = False
Else
    cbsimilarlevel.Enabled = True
End If
End Sub

Private Sub ckbgmmute_Click()
If ckbgmmute.Value = 1 Then
   FormMainMode.cMusicPlayer(0).Mute = True
Else
   FormMainMode.cMusicPlayer(0).Mute = False
End If
End Sub

Private Sub ckendturn_Click()
If ckendturn.Value = 1 Then
    ckendturnnum.Enabled = True
Else
    ckendturnnum.Enabled = False
End If
End Sub

Private Sub cksemute_Click()
If cksemute.Value = 1 Then
    For i = 1 To FormMainMode.cMusicPlayer.UBound
        FormMainMode.cMusicPlayer(i).Mute = True
    Next
Else
    For i = 1 To FormMainMode.cMusicPlayer.UBound
        FormMainMode.cMusicPlayer(i).Mute = False
    Next
End If
End Sub

Private Sub Command1_Click()
On Error GoTo VBEError
cdgBGMchooce.ShowOpen
If lopnmusictext.Caption = "" And cdgBGMchooce.filename <> "" Then
    BGM選擇.AddItem "《其他》"
End If
If cdgBGMchooce.filename <> "" Then
    lopnmusictext.Caption = cdgBGMchooce.filename
    BGM選擇.ListIndex = BGM選擇.ListCount - 1
Else
    BGM選擇.ListIndex = 0
End If
Exit Sub
'===============
VBEError:
BGM選擇.ListIndex = 0
End Sub

Private Sub Command2_Click()
On Error GoTo VBEError
cdgMAPchooce.ShowOpen
If lopnmapjpgtext.Caption = "" And cdgMAPchooce.filename <> "" Then
    對戰地圖選擇.AddItem "《其他》"
    ImageListback.ListImages.Add 15, , LoadPicture(cdgMAPchooce.filename)
ElseIf lopnmapjpgtext.Caption <> "" And cdgMAPchooce.filename <> "" Then
    ImageListback.ListImages.Remove 15
    ImageListback.ListImages.Add 15, , LoadPicture(cdgMAPchooce.filename)
End If
If cdgMAPchooce.filename <> "" Then
    lopnmapjpgtext.Caption = cdgMAPchooce.filename
    對戰地圖選擇.ListIndex = 對戰地圖選擇.ListCount - 1
Else
    對戰地圖選擇.ListIndex = 0
End If
Formsetting.對戰地圖選擇_Click
Exit Sub
'===============
VBEError:
對戰地圖選擇.ListIndex = 0
End Sub

Private Sub Form_Activate()
Formsetting.Width = 9315
Formsetting.Height = 9150
For i = 1 To 3
   If FormMainMode.personnameus(i).Text <> "《隨機》" Then
       personlevelus(i).Caption = FormMainMode.personlevelus(i).Text
       personnameus(i).Caption = FormMainMode.personnameus(i).Text
   Else
       personlevelus(i).Caption = ""
       personnameus(i).Caption = FormMainMode.personnameus(i).Text
   End If
   If FormMainMode.personnamecom(i).Text <> "《隨機》" Then
       personlevelcom(i).Caption = FormMainMode.personlevelcom(i).Text
       personnamecom(i).Caption = FormMainMode.personnamecom(i).Text
   Else
       personlevelcom(i).Caption = ""
       personnamecom(i).Caption = FormMainMode.personnamecom(i).Text
   End If
Next
End Sub

Private Sub Form_Load()
對戰地圖選擇.AddItem "《隨機》"
對戰地圖選擇.AddItem "人魂墓地"
對戰地圖選擇.AddItem "白魔的圓環石陣"
對戰地圖選擇.AddItem "冰封湖畔(新)"
對戰地圖選擇.AddItem "冰封湖畔(舊)"
對戰地圖選擇.AddItem "垃圾之街"
對戰地圖選擇.AddItem "風暴荒野"
對戰地圖選擇.AddItem "烏波斯的黑湖"
對戰地圖選擇.AddItem "萊丁貝魯格城堡"
對戰地圖選擇.AddItem "瘋狂山脈"
對戰地圖選擇.AddItem "盡頭之村"
對戰地圖選擇.AddItem "誘惑森林"
對戰地圖選擇.AddItem "藩骸兒的遺跡"
對戰地圖選擇.AddItem "魔女山谷"
對戰地圖選擇.AddItem "魔都羅占布爾克"
'對戰地圖選擇.ListIndex = 0
'=============
For i = 1 To 14
    ImageListback.ListImages.Add i, , LoadPicture(app_path & "gif\system\map\" & i & ".jpg")
Next
'=============
BGM選擇.AddItem "《隨機-地圖組合》"
BGM選擇.AddItem "《隨機》"
BGM選擇.AddItem "人魂墓地"
'BGM選擇.AddItem "白魔的圓環石陣"
BGM選擇.AddItem "冰封湖畔(新)"
BGM選擇.AddItem "垃圾之街"
BGM選擇.AddItem "風暴荒野"
BGM選擇.AddItem "烏波斯的黑湖"
BGM選擇.AddItem "萊丁貝魯格城堡"
BGM選擇.AddItem "瘋狂山脈"
BGM選擇.AddItem "盡頭之村"
BGM選擇.AddItem "誘惑森林"
BGM選擇.AddItem "藩骸兒的遺跡"
BGM選擇.AddItem "魔都羅占布爾克"
BGM選擇.AddItem "舊版"
'BGM選擇.ListIndex = 0
'=============
cdgBGMchooce.Filter = "MP3音樂檔(*.mp3)|*.mp3|Wave音訊檔(*.wav)|*.wav|MP4音樂檔(*.m4a)|*.m4a|所有檔案(*.*)|*.*"
cdgMAPchooce.Filter = "JPG圖片檔(*.jpg)|*.jpg|BMP點陣圖(*.bmp)|*.bmp|所有檔案(*.*)|*.*"
For i = 1 To 18
    personus(i).AddItem "(無)"
    personus(i).AddItem "劍1"
    personus(i).AddItem "劍2"
    personus(i).AddItem "劍3"
    personus(i).AddItem "劍4"
    personus(i).AddItem "劍5"
    personus(i).AddItem "劍6"
    personus(i).AddItem "劍7"
    personus(i).AddItem "劍8"
    personus(i).AddItem "槍1"
    personus(i).AddItem "槍2"
    personus(i).AddItem "槍3"
    personus(i).AddItem "槍4"
    personus(i).AddItem "槍5"
    personus(i).AddItem "槍6"
    personus(i).AddItem "槍7"
    personus(i).AddItem "槍8"
    personus(i).AddItem "特1"
    personus(i).AddItem "特2"
    personus(i).AddItem "特3"
    personus(i).AddItem "特4"
    personus(i).AddItem "特5"
    personus(i).AddItem "防1"
    personus(i).AddItem "防2"
    personus(i).AddItem "防3"
    personus(i).AddItem "防4"
    personus(i).AddItem "防5"
    personus(i).AddItem "防7"
    personus(i).AddItem "移1"
    personus(i).AddItem "移2"
    personus(i).AddItem "移3"
    personus(i).AddItem "移4"
    personus(i).AddItem "移5"
    personus(i).AddItem "機會1"
    personus(i).AddItem "機會2"
    personus(i).AddItem "機會3"
    personus(i).AddItem "機會4"
    personus(i).AddItem "機會5"
    personus(i).AddItem "詛咒術1"
    personus(i).AddItem "詛咒術2"
    personus(i).AddItem "詛咒術3"
    personus(i).AddItem "詛咒術5"
    personus(i).AddItem "HP回復1"
    personus(i).AddItem "HP回復2"
    personus(i).AddItem "HP回復3"
    personus(i).AddItem "劍3/槍1"
    personus(i).AddItem "劍4/槍2"
    personus(i).AddItem "劍5/槍3"
    personus(i).AddItem "槍3/劍1"
    personus(i).AddItem "槍4/劍2"
    personus(i).AddItem "槍5/劍3"
    personus(i).AddItem "防3/移1"
    personus(i).AddItem "防4/移1"
    personus(i).AddItem "防5/移1"
    personus(i).AddItem "特1/防1"
    personus(i).AddItem "特2/防2"
    personus(i).AddItem "特3/防3"
    personus(i).AddItem "劍3/移1"
    personus(i).AddItem "劍4/移1"
    personus(i).AddItem "劍5/移1"
    personus(i).AddItem "槍3/移1"
    personus(i).AddItem "槍4/移1"
    personus(i).AddItem "槍5/移1"
    personus(i).AddItem "劍3/防1"
    personus(i).AddItem "槍3/防1"
    personus(i).AddItem "移1/特1"
    personus(i).AddItem "移2/特2"
    personus(i).AddItem "移3/特3"
    personcom(i).AddItem "(無)"
    personcom(i).AddItem "劍1"
    personcom(i).AddItem "劍2"
    personcom(i).AddItem "劍3"
    personcom(i).AddItem "劍4"
    personcom(i).AddItem "劍5"
    personcom(i).AddItem "劍6"
    personcom(i).AddItem "劍7"
    personcom(i).AddItem "劍8"
    personcom(i).AddItem "槍1"
    personcom(i).AddItem "槍2"
    personcom(i).AddItem "槍3"
    personcom(i).AddItem "槍4"
    personcom(i).AddItem "槍5"
    personcom(i).AddItem "槍6"
    personcom(i).AddItem "槍7"
    personcom(i).AddItem "槍8"
    personcom(i).AddItem "特1"
    personcom(i).AddItem "特2"
    personcom(i).AddItem "特3"
    personcom(i).AddItem "特4"
    personcom(i).AddItem "特5"
    personcom(i).AddItem "防1"
    personcom(i).AddItem "防2"
    personcom(i).AddItem "防3"
    personcom(i).AddItem "防4"
    personcom(i).AddItem "防5"
    personcom(i).AddItem "防7"
    personcom(i).AddItem "移1"
    personcom(i).AddItem "移2"
    personcom(i).AddItem "移3"
    personcom(i).AddItem "移4"
    personcom(i).AddItem "移5"
    personcom(i).AddItem "機會1"
    personcom(i).AddItem "機會2"
    personcom(i).AddItem "機會3"
    personcom(i).AddItem "機會4"
    personcom(i).AddItem "機會5"
    personcom(i).AddItem "詛咒術1"
    personcom(i).AddItem "詛咒術2"
    personcom(i).AddItem "詛咒術3"
    personcom(i).AddItem "詛咒術5"
    personcom(i).AddItem "HP回復1"
    personcom(i).AddItem "HP回復2"
    personcom(i).AddItem "HP回復3"
    personcom(i).AddItem "劍3/槍1"
    personcom(i).AddItem "劍4/槍2"
    personcom(i).AddItem "劍5/槍3"
    personcom(i).AddItem "槍3/劍1"
    personcom(i).AddItem "槍4/劍2"
    personcom(i).AddItem "槍5/劍3"
    personcom(i).AddItem "防3/移1"
    personcom(i).AddItem "防4/移1"
    personcom(i).AddItem "防5/移1"
    personcom(i).AddItem "特1/防1"
    personcom(i).AddItem "特2/防2"
    personcom(i).AddItem "特3/防3"
    personcom(i).AddItem "劍3/移1"
    personcom(i).AddItem "劍4/移1"
    personcom(i).AddItem "劍5/移1"
    personcom(i).AddItem "槍3/移1"
    personcom(i).AddItem "槍4/移1"
    personcom(i).AddItem "槍5/移1"
    personcom(i).AddItem "劍3/防1"
    personcom(i).AddItem "槍3/防1"
    personcom(i).AddItem "移1/特1"
    personcom(i).AddItem "移2/特2"
    personcom(i).AddItem "移3/特3"
    persontgus(i).Visible = False
    persontgcom(i).Visible = False
'    personus(i).ListIndex = 0
'    personcom(i).ListIndex = 0
Next
'persontgruonus(1).Value = True
'persontgruoncom(1).Value = True
'lopnmusictext.Visible = False
If FormMainMode.personsettingus(1).Caption <> "人物資訊" Then
'    其他設定.Visible = False
'    chkpersonvsmode.Value = 1
'    persontgruonus(4).Value = True
'    ckendturn.Value = 1
    checktest.Value = 1
End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Image2_Click
End Sub

Private Sub Image2_Click()
Formsetting.Visible = False
If Val(挑戰模式選項_牌數.Text) > 30 Then 挑戰模式選項_牌數.Text = 30
If Val(大亂鬥模式選項_牌數.Text) < 1 Then 大亂鬥模式選項_牌數.Text = 1
If Val(ckendturnnum.Text) <= 0 Then
    ckendturnnum.Text = 18
End If
End Sub

Private Sub Label3_Click()
Formsetting.Visible = False
If Val(挑戰模式選項_牌數.Text) > 30 Then 挑戰模式選項_牌數.Text = 30
If Val(大亂鬥模式選項_牌數.Text) < 1 Then 大亂鬥模式選項_牌數.Text = 1
If Val(ckendturnnum.Text) <= 0 Then
    ckendturnnum.Text = 18
End If
End Sub

Private Sub Label9_Click()

End Sub

Private Sub personcom1_Change(Index As Integer)

End Sub

Private Sub personcom1_Click(Index As Integer)

End Sub


Private Sub personcom_Click(Index As Integer)
If personcom(Index).ListIndex <> 0 Then
    If 一般系統類.事件卡資料庫(personcom(Index).Text, 1) <> persontgcom(Index).Caption And _
        一般系統類.事件卡資料庫(personcom(Index).Text, 1) <> 0 And persontgrecom.Value = 1 Then
        MsgBox "此事件卡違反角色色格使用原則!", 64, "UnlightVBE系統提示"
        personcom(Index).ListIndex = 0
    End If
End If
End Sub


Private Sub personnamecom_Change(Index As Integer)
If persontgrecom.Value = 1 Then
    If FormMainMode.opnpersonvs(2).Value = True Then
        If personnamecom(Index).Caption = "《隨機》" Then
            If persontgruoncom(2).Value = True Then
                persontgruoncom(2).Value = False
                persontgruoncom(1).Value = True
                persontgruoncom_Click (1)
            End If
            persontgruoncom(2).Enabled = False
            personwagcom.Visible = True
        ElseIf personnamecom(1).Caption <> "《隨機》" And personnamecom(2).Caption <> "《隨機》" And _
            personnamecom(3).Caption <> "《隨機》" Then
            persontgruoncom(2).Enabled = True
            personwagcom.Visible = False
        End If
    Else
        If personnamecom(1).Caption = "《隨機》" Then
            If persontgruoncom(2).Value = True Then
                persontgruoncom(2).Value = False
                persontgruoncom(1).Value = True
                persontgruoncom_Click (1)
            End If
            persontgruoncom(2).Enabled = False
            personwagcom.Visible = True
        ElseIf personnamecom(1).Caption <> "《隨機》" Then
            persontgruoncom(2).Enabled = True
            personwagcom.Visible = False
        End If
    End If
End If
End Sub

Private Sub personnameus_Change(Index As Integer)
If persontgreus.Value = 1 Then
    If FormMainMode.opnpersonvs(2).Value = True Then
        If personnameus(Index).Caption = "《隨機》" Then
            If persontgruonus(2).Value = True Then
                persontgruonus(2).Value = False
                persontgruonus(1).Value = True
                persontgruonus_Click (1)
            End If
            persontgruonus(2).Enabled = False
            personwagus.Visible = True
        ElseIf personnameus(1).Caption <> "《隨機》" And personnameus(2).Caption <> "《隨機》" And _
            personnameus(3).Caption <> "《隨機》" Then
            persontgruonus(2).Enabled = True
            personwagus.Visible = False
        End If
    Else
        If personnameus(1).Caption = "《隨機》" Then
            If persontgruonus(2).Value = True Then
                persontgruonus(2).Value = False
                persontgruonus(1).Value = True
                persontgruonus_Click (1)
            End If
            persontgruonus(2).Enabled = False
            personwagus.Visible = True
        ElseIf personnameus(1).Caption <> "《隨機》" Then
            persontgruonus(2).Enabled = True
            personwagus.Visible = False
        End If
    End If
End If
End Sub

Private Sub persontgcom_Change(Index As Integer)
Select Case persontgcom(Index).Caption
    Case 0
        personcom(Index).BackColor = RGB(192, 192, 192)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 1
        personcom(Index).BackColor = RGB(255, 0, 0)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 2
        personcom(Index).BackColor = RGB(0, 255, 0)
        personcom(Index).ForeColor = RGB(0, 0, 0)
    Case 3
        personcom(Index).BackColor = RGB(0, 0, 255)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 4
        personcom(Index).BackColor = RGB(255, 0, 255)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 5
        personcom(Index).BackColor = RGB(255, 255, 255)
        personcom(Index).ForeColor = RGB(0, 0, 0)
    Case 6
        personcom(Index).BackColor = RGB(0, 0, 0)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 7
        personcom(Index).BackColor = RGB(255, 255, 0)
        personcom(Index).ForeColor = RGB(0, 0, 0)
End Select
End Sub

Private Sub persontgrecom_Click()
If persontgrecom.Value = 1 Then
    For i = 1 To 18
       personcom_Click (i)
    Next
    For i = 1 To 3
       personnamecom_Change (i)
    Next
Else
   personwagcom.Visible = False
   persontgruoncom(2).Enabled = True
End If

    
End Sub

Private Sub persontgreus_Click()
If persontgreus.Value = 1 Then
    For i = 1 To 18
       personus_Click (i)
    Next
    For i = 1 To 3
       personnameus_Change (i)
    Next
Else
   personwagus.Visible = False
   persontgruonus(2).Enabled = True
End If
End Sub

Private Sub persontgruoncom_Click(Index As Integer)
Select Case Index
    Case 1
       For i = 1 To 18
           personcom(i).Enabled = False
       Next
    Case 2
       For i = 1 To 18
           personcom(i).Enabled = True
       Next
    Case 3
       For i = 1 To 18
           personcom(i).Enabled = False
       Next
    Case 4
       For i = 1 To 18
           personcom(i).Enabled = False
       Next
End Select
End Sub

Private Sub persontgruonus_Click(Index As Integer)
Select Case Index
    Case 1
       For i = 1 To 18
           personus(i).Enabled = False
       Next
    Case 2
       For i = 1 To 18
           personus(i).Enabled = True
       Next
    Case 3
       For i = 1 To 18
           personus(i).Enabled = False
       Next
    Case 4
       For i = 1 To 18
           personus(i).Enabled = False
       Next
End Select
End Sub

Private Sub persontgus_Change(Index As Integer)
Select Case persontgus(Index).Caption
    Case 0
        personus(Index).BackColor = RGB(192, 192, 192)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 1
        personus(Index).BackColor = RGB(255, 0, 0)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 2
        personus(Index).BackColor = RGB(0, 255, 0)
        personus(Index).ForeColor = RGB(0, 0, 0)
    Case 3
        personus(Index).BackColor = RGB(0, 0, 255)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 4
        personus(Index).BackColor = RGB(255, 0, 255)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 5
        personus(Index).BackColor = RGB(255, 255, 255)
        personus(Index).ForeColor = RGB(0, 0, 0)
    Case 6
        personus(Index).BackColor = RGB(0, 0, 0)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 7
        personus(Index).BackColor = RGB(255, 255, 0)
        personus(Index).ForeColor = RGB(0, 0, 0)
End Select
End Sub

Private Sub personus_Click(Index As Integer)
If personus(Index).ListIndex <> 0 Then
    If 一般系統類.事件卡資料庫(personus(Index).Text, 1) <> persontgus(Index).Caption And _
        一般系統類.事件卡資料庫(personus(Index).Text, 1) <> 0 And persontgreus.Value = 1 Then
        MsgBox "此事件卡違反角色色格使用原則!", 64, "UnlightVBE系統提示"
        personus(Index).ListIndex = 0
    End If
End If
End Sub

Private Sub sdrbgm_Change()
bgmve.Caption = sdrbgm.Value
End Sub

Private Sub sdrbgm_Scroll()
bgmve.Caption = sdrbgm.Value
FormMainMode.cMusicPlayer(0).Volume = sdrbgm.Value
End Sub

Private Sub sdrse_Change()
seve.Caption = sdrse.Value
End Sub

Private Sub sdrse_Scroll()
seve.Caption = sdrse.Value
For i = 1 To FormMainMode.cMusicPlayer.UBound
    FormMainMode.cMusicPlayer(i).Volume = sdrse.Value
Next
End Sub
Private Sub t1_Click(PreviousTab As Integer)
Select Case t1.Tab
     Case 0
           一般設定_表單.Left = 120
           一般設定_表單.Top = 2040
           一般設定_表單.Visible = True
           事件卡_使用者.Visible = False
           事件卡_電腦.Visible = False
           一般設定_表單.ZOrder
     Case 1
           事件卡_使用者.Left = 120
           事件卡_使用者.Top = 2040
           事件卡_使用者.Visible = True
           事件卡_電腦.Visible = False
           一般設定_表單.Visible = False
           事件卡_使用者.ZOrder
     Case 2
           事件卡_電腦.Left = 120
           事件卡_電腦.Top = 2040
           事件卡_電腦.Visible = True
           事件卡_使用者.Visible = False
           一般設定_表單.Visible = False
           事件卡_電腦.ZOrder
     Case 3
           表單_系統.Left = 120
           表單_系統.Top = 2040
           事件卡_電腦.Visible = False
           事件卡_使用者.Visible = False
           一般設定_表單.Visible = False
           表單_系統.Visible = True
           表單_系統.ZOrder
End Select
End Sub

Private Sub 大亂鬥模式選項_牌數_Change()
Dim i As Integer, j As Integer, k As Integer
j = 1
Do While j <= Len(大亂鬥模式選項_牌數.Text)
   k = 0
      For k = 0 To 9
         If Asc(Mid(大亂鬥模式選項_牌數.Text, j, 1)) = Asc(k) Then
             j = j + 1
             Exit For
         End If
      Next
      If k = 10 Then
         MsgBox "大小姐，請輸入數字歐...", 64
         大亂鬥模式選項_牌數.Text = ""
         Exit Sub
      End If
Loop
End Sub

Private Sub 大亂鬥選項_Click()
If 大亂鬥選項.Value = 1 Then
    大亂鬥模式選項_牌數.Enabled = True
Else
    大亂鬥模式選項_牌數.Enabled = False
End If
End Sub

Private Sub 挑戰模式選項_Click()
If 挑戰模式選項.Value = 1 Then
    挑戰模式選項_牌數.Enabled = True
Else
   挑戰模式選項_牌數.Enabled = False
End If
End Sub

Private Sub 挑戰模式選項_牌數_Change()
Dim i As Integer, j As Integer, k As Integer
j = 1
Do While j <= Len(挑戰模式選項_牌數.Text)
   k = 0
      For k = 0 To 9
         If Asc(Mid(挑戰模式選項_牌數.Text, j, 1)) = Asc(k) Then
             j = j + 1
             Exit For
         End If
      Next
      If k = 10 Then
         MsgBox "大小姐，請輸入數字歐...", 64
         挑戰模式選項_牌數.Text = ""
         Exit Sub
      End If
Loop

End Sub

Sub 對戰地圖選擇_Click()
Select Case 對戰地圖選擇.Text
   Case "冰封湖畔(舊)"
      BGM選擇.ListIndex = 13
   Case "魔女山谷"
      BGM選擇.ListIndex = 7
   Case "白魔的圓環石陣"
      BGM選擇.ListIndex = 7
   Case Else
      For i = 0 To BGM選擇.ListCount - 1
         BGM選擇.ListIndex = i
         If 對戰地圖選擇.Text = BGM選擇.Text Then
            Exit For
         End If
      Next
End Select
If 對戰地圖選擇.ListIndex > 0 Then
   場地顯示.Picture = ImageListback.ListImages(對戰地圖選擇.ListIndex).Picture
   場地顯示.Visible = True
   randomtext.Visible = False
   randombk.Visible = False
Else
   場地顯示.Visible = False
   randomtext.Visible = True
   randombk.Visible = True
End If

End Sub
