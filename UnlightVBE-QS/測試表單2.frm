VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form 測試表單2 
   Appearance      =   0  '平面
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
   Caption         =   "測試表單2"
   ClientHeight    =   9840
   ClientLeft      =   4005
   ClientTop       =   1275
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "測試表單2.frx":0000
   ScaleHeight     =   9840
   ScaleWidth      =   11340
   Begin VB.Timer t3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8520
      Top             =   480
   End
   Begin VB.PictureBox PEAFcardback 
      Appearance      =   0  '平面
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   8760
      Picture         =   "測試表單2.frx":4ACC1
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   13
      Top             =   3480
      Width           =   2535
      Begin VB.Label Label4 
         Height          =   495
         Left            =   0
         TabIndex        =   55
         Top             =   240
         Width           =   2535
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   34
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   19
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   1
         Left            =   100
         TabIndex        =   15
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   2
         Left            =   100
         TabIndex        =   16
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   3
         Left            =   100
         TabIndex        =   17
         Top             =   1530
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   4
         Left            =   100
         TabIndex        =   18
         Top             =   1960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   20
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   21
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   22
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   23
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   24
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   25
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   26
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   27
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   0
         Left            =   580
         TabIndex        =   31
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   1
         Left            =   740
         TabIndex        =   32
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   2
         Left            =   880
         TabIndex        =   33
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   35
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   36
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   37
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   38
         Top             =   600
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   39
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   40
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   41
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   42
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   43
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   44
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   45
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   46
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   47
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   48
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   49
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   50
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   3
         Left            =   1635
         TabIndex        =   51
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   52
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   53
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin VB.Label PEAFpersoncardback_main 
         BackStyle       =   0  '透明
         Caption         =   "DEF+7。防禦成功時，對手受到與所超過之防禦同值的傷害"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   8.25
            Charset         =   136
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   54
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1245
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   14
         Top             =   315
         Width           =   2295
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   5
         Left            =   120
         Top             =   340
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "測試表單2.frx":506F6
         Props           =   13
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   8400
      Picture         =   "測試表單2.frx":507CB
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   8
      Top             =   2280
      Width           =   2535
      Begin ImageX.aicAlphaImage aicAlphaImage3 
         Height          =   795
         Left            =   480
         Top             =   1560
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1402
         Image           =   "測試表單2.frx":55FC4
         Props           =   13
      End
      Begin UnlightVBE.uc異常狀態 uc異常狀態1 
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin VB.Label Label3 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1020
         TabIndex        =   10
         Top             =   745
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   700
         TabIndex        =   9
         Top             =   745
         Width           =   285
      End
      Begin ImageX.aicAlphaImage aicAlphaImage1 
         Height          =   300
         Left            =   360
         Top             =   720
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Image           =   "測試表單2.frx":586F9
         Scaler          =   2
         Props           =   25
         MaskColor       =   16777215
      End
      Begin ImageX.aicAlphaImage aie2 
         Height          =   300
         Left            =   360
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Image           =   "測試表單2.frx":58812
         Scaler          =   1
         Props           =   25
         MaskColor       =   16777215
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer trtimeline 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5520
      Top             =   7440
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   4200
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7080
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   300
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin UnlightVBE.uc技能視窗 uc技能視窗1 
      Height          =   9855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   17383
   End
   Begin ImageX.aicAlphaImage aicAlphaImage2 
      Height          =   3180
      Left            =   4920
      Top             =   2760
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   5609
      Image           =   "測試表單2.frx":588F3
      Scaler          =   3
      Props           =   13
   End
   Begin VB.Line timelineout2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   6060
      X2              =   11310
      Y1              =   7365
      Y2              =   7365
   End
   Begin VB.Line timelineout1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   0
      X2              =   5250
      Y1              =   7365
      Y2              =   7365
   End
   Begin VB.Image Image4 
      Height          =   105
      Left            =   5295
      Picture         =   "測試表單2.frx":5F12C
      Top             =   7320
      Width           =   750
   End
   Begin VB.Shape timelinein2 
      BackColor       =   &H00808080&
      BorderStyle     =   6  '內實線
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  '實心
      Height          =   90
      Left            =   6045
      Top             =   7320
      Width           =   5295
   End
   Begin VB.Shape timelinein1 
      BorderStyle     =   6  '內實線
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  '實心
      Height          =   90
      Left            =   0
      Top             =   7320
      Width           =   5295
   End
   Begin UnlightVBE.顯示列 顯示列1 
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Width           =   11355
      _ExtentX        =   19923
      _ExtentY        =   2355
   End
   Begin ImageX.aicAlphaImage aie1 
      Height          =   1335
      Left            =   240
      Top             =   1320
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   2355
      Image           =   "測試表單2.frx":5F198
      Scaler          =   3
      Props           =   17
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3600
      TabIndex        =   1
      Top             =   7800
      Width           =   13110
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "測試表單2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal blendFunction As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Dim dc1 As Long
Dim dc2 As Long
Dim dc3 As Long
Dim pe As PictureBox
Dim a As Long

Private Const HWND_BROADCAST = &HFFFF&              '通知其他應用程式加入新的字型
Private Const WM_FONTCHANGE = &H1D

Private Declare Function AddFontResource& Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String)

Private Declare Function RemoveFontResource& Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Const GWL_EXSTYLE = (-20)
'Private Const LWA_COLORKEY = &H1
'Private Const LWA_ALPHA = &H2
'Private Const ULW_COLORKEY = &H1
'Private Const ULW_ALPHA = &H2
'Private Const ULW_OPAQUE = &H4
'Private Const WS_EX_LAYERED = &H80000
'Private Type rBlendProps
'    tBlendOp As Byte
'    tBlendOptions As Byte
'    tBlendAmount As Byte
'    tAlphaType As Byte
'End Type
'Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
'        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
'        ByVal nHeight As Long, ByVal hSrcDC As Long, _
'        ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
'        ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'        (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
 
'Dim ctlNew As PictureBox, ctlNewWnd As Long


Private Sub Command1_Click()
'Picture1 = LoadPicture(App.Path & "\gif\lineusatk1拷貝.gif")
'Picture2 = LoadPicture(App.Path & "\gif\linemove3拷貝.gif")
'Dim i As Byte
'    For i = 1 To 2
'        Me.PaintPicture Picture2, 0, 150, , , , , , , vbSrcAnd
'        Me.PaintPicture Picture1, 0, 150, , , , , , , vbSrcInvert
'    Next i
'    For i = 1 To 2
'        BitBlt Me.hdc, 0, 208, 761, 81, dc2, 0, 0, vbSrcAnd
'        BitBlt Me.hdc, 0, 208, 761, 81, dc1, 0, 0, vbSrcInvert
'    Next i
''    TransparentBlt Me.hdc, 0, 0, 761, 81, dc1, 0, 0, 761, 81, vbWhite
'    Me.Refresh
''''''''''''AlphaBlend Me.hdc, 0, 100, 765, 85, Picture1.hdc, 0, 0, 765, 85, 150 * &H10000
'Picture1.Visible = False
'MakeTransparent Me.hWnd, 20
'TransPic2 Picture1, Picture3, 150
'==================
'sf1.ScaleMode = 1
''sf1.WMode = transparent
'sf1.Movie = App.Path & "\cutinchara32_skl02.swf"
'''''小人物形象1.小人物影子top = 3500
'''''小人物形象1.小人物圖片 = App.Path & "\gif\帕茉\Palmomini1.gif"
'''''小人物形象1.小人物影子圖片 = App.Path & "\gif\帕茉\image 4.gif"
'''''小人物形象1.Visible = True


'aie1.LoadImage_FromFile "C:\Users\Andy Ciu\Desktop\lineusatk1-2.png"
顯示列1.Left = 0
'顯示列1.顯示列圖片 = "C:\Users\Andy Ciu\Desktop\lineusatk1-2.png"
顯示列1.顯示列圖片 = App.Path & "\gif\system\linemove.png"
顯示列1.使用者方小人物圖片 = App.Path & "\gif\古魯瓦爾多\Grunwaldf1.gif"
顯示列1.電腦方小人物圖片 = App.Path & "\gif\雪莉\sherif2.gif"
顯示列1.goi1顯示 = False
顯示列1.goi2顯示 = False
顯示列1.移動階段圖顯示 = True
顯示列1.移動階段選擇值 = 0
'aie2.AutoSize = True
'aie2.AutoRedraw = True
'aie2.RotateEnabled = True
'aie2.LoadImage_FromFile App.Path & "\gif\艾茵\Aynmini.png"
'aie2.rotation = 180
'小人物形象1.小人物圖片 = App.Path & "\gif\雪莉\sherimini.png"
'小人物形象1.小人物影子圖片 = App.Path & "\gif\雪莉\sheriminidown1.gif"
'小人物形象1.小人物影子top差 = -150
't1.Enabled = True
End Sub
Sub TransPic2(cSrc As PictureBox, cDest As PictureBox, ByVal nLevel As Byte)
'Dim LrProps As rBlendProps
'Dim LnBlendPtr As Long
'Dim Mode As Integer, AutoDraw As Boolean
'    '保存?置
'    Mode = cSrc.ScaleMode
'    AutoDraw = cDest.AutoRedraw
'    cSrc.ScaleMode = 3
'    cDest.AutoRedraw = True
'
'    '透明?理
'    cDest.Cls
'    LrProps.tBlendAmount = nLevel
'    CopyMemory LnBlendPtr, LrProps, 4
'    With cSrc
'        AlphaBlend cDest.hdc, 0, 0, .ScaleWidth, .ScaleHeight, _
'                .hdc, 0, 0, .ScaleWidth, .ScaleHeight, LnBlendPtr
'    End With
'    cDest.Refresh
'
'    '恢复?置
'    cSrc.ScaleMode = Mode
'    cDest.AutoRedraw = AutoDraw
End Sub

Private Sub Command2_Click()
'顯示列1.使用者方移動值 = 4
'顯示列1.使用者方移動內外 = 2
'顯示列1.電腦方移動值 = 4
'顯示列1.電腦方移動內外 = 2
'顯示列1.移動方向圖片顯示 = True
'Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'
'Dim PathName As String
'PathName = "D:\ulse06.mp3"
'mciSendString "close MyWav", vbNullString, 0, 0
'mciSendString "open " & PathName & " alias MyWav", vbNullString, 0, 0
'mciSendString "play MyWav", vbNullString, 0, 0
End Sub

Private Sub Command3_Click()
'Label1.Caption = "12345678" & Chr(10) & "9123456798"
'Label1.Width = 81
'Dim ret As Long
'Dim str As String
'str = App.Path & "\ttf\dfttww5.TTC"
'    ret = AddFontResource(str)                 '成功的話會傳回加入字型的數目,失敗傳回0.
'    If ret = 0 Then MsgBox "加入字型失敗!!", vbExclamation: Exit Sub
''SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
'Label1.Font = "華康娃娃體".
'MsgBox Val(Text1.Text)
'MsgBox Asc(9)
'Dim aa(2) As Boolean
'aa(1) = True
'MsgBox aa(1)
'Erase aa
'MsgBox aa(1)
'trtimeline.Enabled = True
'測試表單2.trtimeline_Timer
'uc異常狀態1.person_num = 0
'uc異常狀態1.person_turn = 10

'uc技能視窗1.技能圖片 = App.Path & "\gif\system\map\魔女山谷.jpg"
'uc技能視窗1.ZOrder
'Picture2.Picture = LoadPicture()
t3.Enabled = True
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Activate()
'dc1 = CreateCompatibleDC(0)
'SelectObject dc1, Picture1
'
'dc2 = CreateCompatibleDC(0)
'SelectObject dc2, Picture2
'mmmm = 2
End Sub

Private Sub Form_Load()
'Picture1.Visible = False
'pe.Picture = LoadPicture(App.Path & "\gif\lineusatk1.jpg")
'dc1 = CreateCompatibleDC(0)
'SelectObject dc1, Picture1
'小人物形象1.Visible = False
'顯示列1.顯示列圖片 = "C:\Users\Andy Ciu\Desktop\lineusatk1-2.png"
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteDC dc1
'DeleteDC dc2
'DeleteDC dc3
End Sub
Public Function isTransparent(ByVal hWnd As Long) As Boolean
'On Error Resume Next
'Dim Msg As Long
'Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
'If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
'isTransparent = True
'Else
'isTransparent = False
'End If
'If Err Then
'isTransparent = False
'End If
End Function

Public Function MakeTransparent(ByVal hWnd As Long, ByVal Perc As Integer) As Long
'Dim Msg As Long
'On Error Resume Next
'
'Perc = 100
'If Perc < 0 Or Perc > 255 Then
'MakeTransparent = 1
'Else
'Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
'Msg = Msg Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, Msg
'SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
'MakeTransparent = 0
'End If
'If Err Then
'MakeTransparent = 2
'End If
End Function

Public Function MakeOpaque(ByVal hWnd As Long) As Long
'Dim Msg As Long
'On Error Resume Next
'Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
'Msg = Msg And Not WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, Msg
'SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
'MakeOpaque = 0
'If Err Then
'MakeOpaque = 2
'End If
End Function

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "123142"
End Sub

Private Sub PEAFcardbackBR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "123"
End Sub

Private Sub t1_Timer()
'aie1.Left = aie1.Left + 5
'If aie1.Left >= 測試表單2.ScaleWidth Then t1.Enabled = False
顯示列1.使用者方小人物圖片left = 顯示列1.使用者方小人物圖片left - 10
顯示列1.電腦方小人物圖片left = 顯示列1.電腦方小人物圖片left + 10
End Sub

Private Sub t3_Timer()
If aicAlphaImage2.Opacity <> 0 Then
    aicAlphaImage2.Opacity = aicAlphaImage2.Opacity - 1
Else
    t3.Enabled = False
End If
End Sub

Sub trtimeline_Timer()
timelineout1.X1 = timelineout1.X1 + 2
timelineout2.X2 = timelineout2.X2 - 2
If timelineout1.X1 = timelineout1.X2 Then 戰鬥系統類.時間軸_停止
End Sub

