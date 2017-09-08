VERSION 5.00
Begin VB.Form 測試表單 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE-影子設定"
   ClientHeight    =   6360
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
   Icon            =   "測試表單.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "測試表單.frx":0CCA
   ScaleHeight     =   424
   ScaleMode       =   3  '像素
   ScaleWidth      =   432
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      Caption         =   "顯示列"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   47
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
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton personfcomright 
         Caption         =   "R"
         Height          =   375
         Left            =   4080
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton personfcomleft 
         Caption         =   "L"
         Height          =   375
         Left            =   3600
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton personfusright 
         Caption         =   "R"
         Height          =   375
         Left            =   2520
         TabIndex        =   53
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton personfusleft 
         Caption         =   "L"
         Height          =   375
         Left            =   2040
         TabIndex        =   52
         Top             =   240
         Width           =   495
      End
      Begin VB.Label personfcom 
         Height          =   375
         Left            =   5520
         TabIndex        =   51
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '透明
         Caption         =   "右Left:"
         Height          =   375
         Left            =   4680
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label personfus 
         Height          =   375
         Left            =   840
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "左Left:"
         Height          =   375
         Left            =   120
         TabIndex        =   48
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
      TabIndex        =   37
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton smallpntdncom 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   41
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton smallpntupcom 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   40
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallpnrcom 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   39
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallpnlcom 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   38
         Top             =   960
         Width           =   495
      End
      Begin VB.Label smallpntopcom 
         Height          =   375
         Left            =   1440
         TabIndex        =   46
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label smallpnleftcom 
         Height          =   375
         Left            =   1560
         TabIndex        =   45
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   44
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   43
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "COM(人物)"
         Height          =   255
         Left            =   480
         TabIndex        =   42
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
      TabIndex        =   27
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton smallpnlus 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   31
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallpnrus 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   30
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallpntupus 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallpntdnus 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   28
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   35
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label smallpnleftus 
         Height          =   375
         Left            =   1560
         TabIndex        =   34
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label smallpntopus 
         Height          =   375
         Left            =   1440
         TabIndex        =   33
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "1P(人物)"
         Height          =   375
         Left            =   480
         TabIndex        =   32
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
      TabIndex        =   12
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton smallclcom 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallcrcom 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallctupcom 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallctdncom 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "COM(影子)"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label smallleftcom 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label smalltopcom 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
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
      TabIndex        =   3
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton smallctdnus 
         Caption         =   "Tdn"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton smallctupus 
         Caption         =   "Tup"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton smallcrus 
         Caption         =   "R"
         Height          =   420
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton smallclus 
         Caption         =   "L"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "1P(影子)"
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label smalltopus 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label smallleftus 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Top:"
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Left:"
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   7560
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   0
      Picture         =   "測試表單.frx":31CC1
      ScaleHeight     =   1275
      ScaleWidth      =   11475
      TabIndex        =   1
      Top             =   6120
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   1260
      Left            =   0
      Picture         =   "測試表單.frx":31EE6
      ScaleHeight     =   1260
      ScaleWidth      =   11535
      TabIndex        =   0
      Top             =   4320
      Width           =   11535
      Begin VB.PictureBox Picture5 
         Height          =   735
         Left            =   4080
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   26
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture8 
         Height          =   735
         Left            =   6600
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   25
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture7 
         Height          =   735
         Left            =   5760
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture6 
         Height          =   735
         Left            =   4920
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   2400
         Top             =   120
         Width           =   375
      End
   End
End
Attribute VB_Name = "測試表單"
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
 
'Dim ctlNew As PictureBox, ctlNewWnd As Long


Private Sub Command1_Click()
Picture1 = LoadPicture(App.Path & "\gif\lineusatk1拷貝.gif")
Picture2 = LoadPicture(App.Path & "\gif\linemove3拷貝.gif")
Dim i As Byte
    For i = 1 To 2
        Me.PaintPicture Picture2, 0, 150, , , , , , , vbSrcAnd
        Me.PaintPicture Picture1, 0, 150, , , , , , , vbSrcInvert
    Next i
'    For i = 1 To 2
'        BitBlt Me.hdc, 0, 208, 761, 81, dc2, 0, 0, vbSrcAnd
'        BitBlt Me.hdc, 0, 208, 761, 81, dc1, 0, 0, vbSrcInvert
'    Next i
''    TransparentBlt Me.hdc, 0, 0, 761, 81, dc1, 0, 0, 761, 81, vbWhite
'    Me.Refresh
''''''''AlphaBlend Me.hdc, 0, 100, 765, 85, Picture1.hdc, 0, 0, 765, 85, 150 * &H10000
''''''''Picture1.Visible = False
'MakeTransparent Me.hWnd, 20
'TransPic2 Picture1, Picture3, 150
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
戰鬥系統類.時間軸_停止
End Sub

Private Sub Command3_Click()
FormMainMode.trtimeline.Enabled = True
End Sub

Private Sub Form_Activate()
'dc1 = CreateCompatibleDC(0)
'SelectObject dc1, Picture1
'
'dc2 = CreateCompatibleDC(0)
'SelectObject dc2, Picture2
End Sub

Private Sub Form_Load()
'Picture1.Visible = False
'pe.Picture = LoadPicture(App.Path & "\gif\lineusatk1.jpg")
'dc1 = CreateCompatibleDC(0)
'SelectObject dc1, Picture1
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

Private Sub smallcl_Click()

End Sub

Private Sub smallcr_Click()

End Sub

Private Sub smallctdn_Click()

End Sub

Private Sub smallctup_Click()

End Sub

Private Sub smalltop_Click()

End Sub

Private Sub Label15_Click()

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
