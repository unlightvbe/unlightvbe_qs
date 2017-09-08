VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   ClientHeight    =   7905
   ClientLeft      =   5265
   ClientTop       =   1770
   ClientWidth     =   9720
   Icon            =   "f.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9720
   Begin VB.PictureBox jpgcom_test 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   3960
      Picture         =   "f.frx":0CCA
      ScaleHeight     =   7935
      ScaleWidth      =   8535
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.PictureBox jpgus_test 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   600
      Picture         =   "f.frx":6209
      ScaleHeight     =   8775
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   4215
   End
   Begin UnlightVBE.uc攻擊防禦視窗 adfe 
      Height          =   8175
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   4515
      _ExtentX        =   8705
      _ExtentY        =   14420
   End
   Begin VB.Timer trwait 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   120
   End
   Begin VB.Timer trout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9240
      Top             =   0
   End
   Begin VB.Timer trjpgshow 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   240
   End
   Begin VB.Timer trhide 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   6000
   End
   Begin VB.Timer trshow 
      Enabled         =   0   'False
      Interval        =   130
      Left            =   120
      Top             =   6000
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   6960
      X2              =   6960
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   2760
      X2              =   2760
      Y1              =   0
      Y2              =   600
   End
   Begin UnlightVBE.大人物形像 jpgcom 
      Height          =   7935
      Left            =   10000
      TabIndex        =   4
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13996
   End
   Begin UnlightVBE.大人物形像 jpgus 
      Height          =   7935
      Left            =   -10000
      TabIndex        =   3
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   13996
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim app_path As String
Dim showus, showcom, showendus, showendcom, trshowendus, trshowendcom, hideall, timeout, tot As Integer
Dim 距離單位(1 To 1, 1 To 2, 1 To 2) As Integer  '距離單位暫時儲存資料(1.HP血條,1.使用者/2.電腦,1.Left單位/2.Top單位)
Dim bigallzero(1 To 2) As Integer

Private Sub Form_Activate()
'距離單位(1, 1, 1) = 5295 \ Val(form7.usbi1.Caption)
'距離單位(1, 2, 1) = (11580 - 6060) \ Val(form7.cardcompi1.Caption)

If Val(擲骰表單溝通暫時變數(4)) = 1 Then
  If Val(擲骰表單溝通暫時變數(1)) = 1 Then
    adfe.Left = 5160
  Else
    adfe.Left = 0
  End If
Else
  If Val(擲骰表單溝通暫時變數(1)) = 2 Then
    adfe.Left = 5160
  Else
    adfe.Left = 0
  End If
End If
'-----以下為人物頭像調整
'=====================
jpgus.Height = jpgus.大人物圖片height
jpgus.Width = jpgus.大人物圖片width
jpgus.Top = Form6.ScaleHeight - jpgus.大人物圖片height
If Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 2, 5)) = 1 Then
    bigallzero(1) = 1
Else
    bigallzero(1) = 0
End If
'Select Case 角色人物對戰人數(1, 2)
'   Case 1
''        jpgus.Height = formsettingpersonus.bight.Text
''        jpgus.Top = formsettingpersonus.bigtop.Text
''        jpgus.Width = formsettingpersonus.bigwh.Text
'        jpgus.Height = jpgus.大人物圖片height
'        jpgus.Width = jpgus.大人物圖片width
''        jpgus.Top = formsettingpersonus.bigtop.Text
'        jpgus.Top = Form6.ScaleHeight - jpgus.大人物圖片height
''        If formsettingpersonus.atkingjpgleftallzero.Value = 1 Then
'        If formsettingpersonus.atkingjpgleftallzero.Value = 1 Then
'            bigallzero(1) = 1
'        Else
'            bigallzero(1) = 0
'        End If
'   Case 2
''        jpgus.Height = formsettingpersonus2.bight.Text
''        jpgus.Top = formsettingpersonus2.bigtop.Text
''        jpgus.Width = formsettingpersonus2.bigwh.Text
'        jpgus.Height = jpgus.大人物圖片height
'        jpgus.Width = jpgus.大人物圖片width
'        jpgus.Top = Form6.ScaleHeight - jpgus.大人物圖片height
'        If formsettingpersonus2.atkingjpgleftallzero.Value = 1 Then
'            bigallzero(1) = 1
'        Else
'            bigallzero(1) = 0
'        End If
'   Case 3
''        jpgus.Height = formsettingpersonus3.bight.Text
''        jpgus.Top = formsettingpersonus3.bigtop.Text
''        jpgus.Width = formsettingpersonus3.bigwh.Text
'        jpgus.Height = jpgus.大人物圖片height
'        jpgus.Width = jpgus.大人物圖片width
'        jpgus.Top = Form6.ScaleHeight - jpgus.大人物圖片height
'        If formsettingpersonus3.atkingjpgleftallzero.Value = 1 Then
'            bigallzero(1) = 1
'        Else
'            bigallzero(1) = 0
'        End If
'End Select
'=================
jpgcom.Height = jpgcom.大人物圖片height
jpgcom.Width = jpgcom.大人物圖片width
jpgcom.Top = Form6.ScaleHeight - jpgcom.大人物圖片height
If Val(VBEPerson(2, 角色人物對戰人數(2, 2), 2, 2, 5)) = 1 Then
    bigallzero(2) = 1
Else
    bigallzero(2) = 0
End If
'Select Case 角色人物對戰人數(2, 2)
'   Case 1
''        jpgcom.Height = formsettingpersoncom.bight.Text
''        jpgcom.Top = formsettingpersoncom.bigtop.Text
''        jpgcom.Width = formsettingpersoncom.bigwh.Text
'        jpgcom.Height = jpgcom.大人物圖片height
'        jpgcom.Width = jpgcom.大人物圖片width
'        jpgcom.Top = Form6.ScaleHeight - jpgcom.大人物圖片height
'        If VBEPerson(2, 1, 2, 2, 5) = 1 Then
'            bigallzero(2) = 1
'        Else
'            bigallzero(2) = 0
'        End If
'   Case 2
''        jpgcom.Height = formsettingpersoncom2.bight.Text
''        jpgcom.Top = formsettingpersoncom2.bigtop.Text
''        jpgcom.Width = formsettingpersoncom2.bigwh.Text
'        jpgcom.Height = jpgcom.大人物圖片height
'        jpgcom.Width = jpgcom.大人物圖片width
'        jpgcom.Top = Form6.ScaleHeight - jpgcom.大人物圖片height
'        If formsettingpersoncom2.atkingjpgleftallzero.Value = 1 Then
'            bigallzero(2) = 1
'        Else
'            bigallzero(2) = 0
'        End If
'   Case 3
''        jpgcom.Height = formsettingpersoncom3.bight.Text
''        jpgcom.Top = formsettingpersoncom3.bigtop.Text
''        jpgcom.Width = formsettingpersoncom3.bigwh.Text
'        jpgcom.Height = jpgcom.大人物圖片height
'        jpgcom.Width = jpgcom.大人物圖片width
'        jpgcom.Top = Form6.ScaleHeight - jpgcom.大人物圖片height
'        If formsettingpersoncom3.atkingjpgleftallzero.Value = 1 Then
'            bigallzero(2) = 1
'        Else
'            bigallzero(2) = 0
'        End If
'End Select
'----------------
adfe.Visible = False
jpgus.Left = -jpgus.Width
jpgcom.Left = 9360
jpgus.Visible = False
jpgcom.Visible = False
'===============
trjpgshow.Enabled = True
trshow.Enabled = True
showus = 1
showcom = 1
trshowendus = 0
trshowendcom = 0
Randomize
showendus = Val(擲骰表單溝通暫時變數(5))
showendcom = Val(擲骰表單溝通暫時變數(6))
If showendus > 20 Then
   showendus = 20
End If
If showendcom > 20 Then
   showendcom = 20
End If
End Sub

Private Sub Form_Load()
    app_path = App.Path
    If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
    jpgus_test.Visible = False
    jpgcom_test.Visible = False
End Sub

Private Sub jpgusback_Click()

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

Private Sub trjpgshow_Timer()
Dim bigall(1 To 2) As Integer
Dim bigw(1 To 2) As Integer
bigw(1) = Val(jpgus.大人物圖片width) / 2
bigw(2) = Val(jpgcom.大人物圖片width) / 2
If 2760 - bigw(1) < 0 Or bigallzero(1) = 1 Then
    bigall(1) = 0
Else
    bigall(1) = 2760 - bigw(1)
End If
If 6960 - bigw(2) > Val(Form6.Width) - Val(jpgcom.大人物圖片width) Or bigallzero(2) = 1 Then
    bigall(2) = Val(Form6.Width) - Val(jpgcom.大人物圖片width)
Else
    bigall(2) = 6960 - bigw(2)
End If
If Val(擲骰表單溝通暫時變數(4)) = 1 Then
  If Val(擲骰表單溝通暫時變數(1)) = 1 Then
    jpgus.Visible = True
    jpgus.Left = Val(jpgus.Left) + 150
'    If Val(jpgus.Left) >= Val(formsettingpersonus.bigleftall) Then
    If bigall(1) - Val(jpgus.Left) <= 150 Then
      jpgus.Left = bigall(1)
      trjpgshow.Enabled = False
    End If
  Else
    jpgcom.Visible = True
    jpgcom.Left = Val(jpgcom.Left) - 150
    If Val(jpgcom.Left) <= bigall(2) Then
      trjpgshow.Enabled = False
    End If
  End If
Else
  If Val(擲骰表單溝通暫時變數(1)) = 2 Then
    jpgus.Visible = True
    jpgus.Left = Val(jpgus.Left) + 150
'    If Val(jpgus.Left) >= Val(formsettingpersonus.bigleftall) Then
    If bigall(1) - Val(jpgus.Left) <= 150 Then
      jpgus.Left = bigall(1)
      trjpgshow.Enabled = False
    End If
  Else
    jpgcom.Visible = True
    jpgcom.Left = Val(jpgcom.Left) - 150
    If Val(jpgcom.Left) <= bigall(2) Then
      trjpgshow.Enabled = False
    End If
  End If
End If
End Sub

Sub trout_Timer()
If timeout = 0 Then
  timeout = Val(timeout) + 1
Else
  outprocess
End If
End Sub
Sub outprocess()
  Form6.Visible = False
'  FormMainMode.atkingnumtot.Caption = -2
  trout.Enabled = False
  If Val(擲骰表單溝通暫時變數(4)) = 1 Then
   Select Case Val(擲骰表單溝通暫時變數(1))
    Case 1
       usatkcom
    Case 2
       comatkus
    End Select
  Else
   Select Case Val(擲骰表單溝通暫時變數(1))
    Case 1
       comatkus
    Case 2
       usatkcom
    End Select
  End If
'  FormMainMode.骰子執行完啟動.Enabled = True
'  Unload Me
End Sub
Sub usatkcom()
     tot = Val(擲骰表單溝通暫時變數(5)) - Val(擲骰表單溝通暫時變數(6))
'======以下為異常狀態檢查及啟動
'formmainmode.技能.蕾_終曲_無盡輪迴的終結_舊   '(階段3)
'atkingck(17, 1) = 3
'技能.帕茉_靜謐之背 '(階段3)
'=========
擲骰表單溝通暫時變數(2) = tot
'擲骰後骰傷害數 = tot
擲骰表單溝通暫時變數(3) = 2
End Sub
Sub comatkus()
 tot = Val(擲骰表單溝通暫時變數(6)) - Val(擲骰表單溝通暫時變數(5))
'======以下為異常狀態檢查及啟動
'formmainmode.異常狀態.不死_使用者  '(階段1)
'=========
擲骰表單溝通暫時變數(2) = tot
'擲骰後骰傷害數 = tot
擲骰表單溝通暫時變數(3) = 1
End Sub
'Sub 異常狀態.不死_使用者_分支_階段一()
'tot = 0
'End Sub
'Sub 技能_蕾_終曲_無盡輪迴的終結_舊_分支_階段三()
'If tot < -7 Then
'   tot = livecom(角色人物對戰人數(2, 2))
'End If
'End Sub
'Function 技能_帕茉_靜謐之背_分支_階段三(turn As Integer)
'tot = tot + turn
'End Function
Private Sub trshow_Timer()
If 擲骰表單溝通暫時變數(4) = 1 Then
  Select Case Val(擲骰表單溝通暫時變數(1))
   Case 1
      showusatk
   Case 2
      showcomatk
  End Select
ElseIf 擲骰表單溝通暫時變數(4) = 2 Then
  Select Case Val(擲骰表單溝通暫時變數(1))
   Case 1
      showcomatk
   Case 2
      showusatk
  End Select
End If
End Sub
Function showusatk()
    adfe.adust = showendus
    adfe.adcomt = showendcom
    adfe.adcge = 1
    adfe.adwait = False
    adfe.Visible = True
    trshow.Enabled = False
    trwait.Enabled = True
End Function
Function showcomatk()
    adfe.adust = showendus
    adfe.adcomt = showendcom
    adfe.adcge = 2
    adfe.adwait = False
    adfe.Visible = True
    trshow.Enabled = False
    trwait.Enabled = True
End Function

Private Sub trwait_Timer()
If adfe.adwait = True Then
    'trshow.Enabled = False
   'hideall = 1
    trwait.Enabled = False
   'trhide.Enabled = True
    timeout = 0
    trout.Enabled = True
End If
End Sub

