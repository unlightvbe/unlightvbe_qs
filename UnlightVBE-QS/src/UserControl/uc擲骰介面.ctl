VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc擲骰介面 
   Appearance      =   0  '平面
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   9915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   ClipControls    =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   11340
   Windowless      =   -1  'True
   Begin VB.Timer trdiceoff_all 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7080
      Top             =   600
   End
   Begin VB.Timer trdiceoff_true 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6360
      Top             =   600
   End
   Begin VB.Timer trdiceon_true 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5640
      Top             =   600
   End
   Begin VB.Timer trdiceoff_tot 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4680
      Top             =   600
   End
   Begin VB.Timer trdiceshow_tot 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   600
   End
   Begin VB.Timer trpersonshowoff 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2400
      Top             =   600
   End
   Begin UnlightVBE.ucMusicPlayer ucMusicPlayer 
      Height          =   735
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
   End
   Begin ImageX.aicAlphaImage dicemini1 
      Height          =   375
      Index           =   0
      Left            =   3720
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Image           =   "uc擲骰介面.ctx":0000
      Scaler          =   1
      Props           =   13
   End
   Begin ImageX.aicAlphaImage dicemini2 
      Height          =   375
      Index           =   0
      Left            =   3720
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Image           =   "uc擲骰介面.ctx":077E
      Scaler          =   1
      Props           =   13
   End
   Begin ImageX.aicAlphaImage dice2 
      Height          =   750
      Index           =   0
      Left            =   3720
      Top             =   4080
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      Image           =   "uc擲骰介面.ctx":1186
      Props           =   13
   End
   Begin ImageX.aicAlphaImage dice1 
      Height          =   750
      Index           =   0
      Left            =   3720
      Top             =   5040
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      Image           =   "uc擲骰介面.ctx":1B8E
      Props           =   13
   End
   Begin ImageX.aicAlphaImage person 
      Height          =   4590
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8096
      Image           =   "uc擲骰介面.ctx":2380
      Scaler          =   3
      Props           =   9
   End
End
Attribute VB_Name = "uc擲骰介面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_diceusStr As String, m_dicecomStr As String
Attribute m_dicecomStr.VB_VarUserMemId = 1073938432
Dim m_diceusTotal As Integer, m_dicecomTotal As Integer
Attribute m_diceusTotal.VB_VarUserMemId = 1073938434
Attribute m_dicecomTotal.VB_VarUserMemId = 1073938434
Dim m_diceusTrue As Integer, m_dicecomTrue As Integer
Attribute m_diceusTrue.VB_VarUserMemId = 1073938436
Attribute m_dicecomTrue.VB_VarUserMemId = 1073938436
Dim m_start As Boolean, m_stop As Boolean
Attribute m_start.VB_VarUserMemId = 1073938438
Attribute m_stop.VB_VarUserMemId = 1073938438
Dim m_dicevoice As Integer  '擲骰子聲音音量
Attribute m_dicevoice.VB_VarUserMemId = 1073938440
Dim m_DiceATKMode As Integer    '1-使用者攻(上DEF下ATK)/2-電腦攻(上ATK下DEF)
Attribute m_DiceATKMode.VB_VarUserMemId = 1073938441
Dim m_PersonImage As String
Attribute m_PersonImage.VB_VarUserMemId = 1073938442
Dim m_PersonImageLeftZero As Boolean    '圖片標記是否對齊邊框
Attribute m_PersonImageLeftZero.VB_VarUserMemId = 1073938443
Dim TrPersonShowOffNum As Integer    '1-出來/2-退去
Attribute TrPersonShowOffNum.VB_VarUserMemId = 1073938444
Dim TrDiceShowTotalNum1 As Integer, TrDiceShowTotalNum2 As Integer    'Num1-第x顆骰子/Num2-骰子透明度進度
Attribute TrDiceShowTotalNum1.VB_VarUserMemId = 1073938445
Attribute TrDiceShowTotalNum2.VB_VarUserMemId = 1073938445
Dim m_DiceInputMode As Integer    '1-讀取骰數串，並輸出總骰數及真實骰數/2-讀取總骰數，並輸出骰數串及真實骰數
Attribute m_DiceInputMode.VB_VarUserMemId = 1073938447
Dim TrDiceOffTotNum1 As Integer, TrDiceOffTotNum2 As Integer  'Num1-時間紀錄/Num2-骰子透明度進度
Attribute TrDiceOffTotNum1.VB_VarUserMemId = 1073938448
Attribute TrDiceOffTotNum2.VB_VarUserMemId = 1073938448
Dim TrDiceOnTrueNum As Integer    'Num-骰子透明度進度
Attribute TrDiceOnTrueNum.VB_VarUserMemId = 1073938450
Dim TrDiceOffTrueNum1 As Integer, TrDiceOffTrueNum2 As Integer    'Num1-第x顆骰子/Num2-骰子透明度進度
Attribute TrDiceOffTrueNum1.VB_VarUserMemId = 1073938451
Attribute TrDiceOffTrueNum2.VB_VarUserMemId = 1073938451
Dim TrDiceOffAllNum1 As Integer, TrDiceOffAllNum2 As Integer, TrDiceOffAllNum3 As Integer    'Num1-時間紀錄/Num2-骰子透明度進度(總)/Num3-骰子透明度進度(正面)
Attribute TrDiceOffAllNum1.VB_VarUserMemId = 1073938453
Attribute TrDiceOffAllNum2.VB_VarUserMemId = 1073938453
Attribute TrDiceOffAllNum3.VB_VarUserMemId = 1073938453
Public Property Get diceusStr() As String
    diceusStr = m_diceusStr
End Property
Public Property Let diceusStr(ByVal New_diceusStr As String)
    m_diceusStr = New_diceusStr
    PropertyChanged "diceusStr"
End Property
Public Property Get dicecomStr() As String
    dicecomStr = m_dicecomStr
End Property
Public Property Let dicecomStr(ByVal New_dicecomStr As String)
    m_dicecomStr = New_dicecomStr
    PropertyChanged "dicecomStr"
End Property
'=========================================================
Public Property Get diceusTotal() As Integer
    diceusTotal = m_diceusTotal
End Property
Public Property Let diceusTotal(ByVal New_diceusTotal As Integer)
    m_diceusTotal = New_diceusTotal
    PropertyChanged "diceusTotal"
End Property
Public Property Get dicecomTotal() As Integer
    dicecomTotal = m_dicecomTotal
End Property
Public Property Let dicecomTotal(ByVal New_dicecomTotal As Integer)
    m_dicecomTotal = New_dicecomTotal
    PropertyChanged "dicecomTotal"
End Property
'==========================================================
Public Property Get diceusTrue() As Integer
    diceusTrue = m_diceusTrue
End Property
Public Property Let diceusTrue(ByVal New_diceusTrue As Integer)
    m_diceusTrue = New_diceusTrue
    PropertyChanged "diceusTrue"
End Property
Public Property Get dicecomTrue() As Integer
    dicecomTrue = m_dicecomTrue
End Property
Public Property Let dicecomTrue(ByVal New_dicecomTrue As Integer)
    m_dicecomTrue = New_dicecomTrue
    PropertyChanged "dicecomTrue"
End Property
'==========================================================
Public Property Get DiceStart() As Boolean
    DiceStart = m_start
End Property
Public Property Let DiceStart(ByVal New_DiceStart As Boolean)
    m_start = New_DiceStart
    PropertyChanged "DiceStart"
    '==================
    If Me.DiceStart = True Then
        Me.DiceStop = False
        Select Case Me.DiceATKMode
            Case 1
                person.Left = -person.Width
                person.Mirror = aiMirrorNone
            Case 2
                person.Left = 11340
                person.Mirror = aiMirrorHorizontal
        End Select
        person.Opacity = 100
        person.Visible = True
        TrPersonShowOffNum = 1
        TrDiceShowTotalNum1 = 0
        TrDiceShowTotalNum2 = 10
        '==================
        Dim i As Integer, nm1 As Integer, nm2 As Integer, k As Integer
        Dim nmstr1 As String, nmstr2 As String
        Select Case Me.DiceInputMode
            Case 1
                For i = 1 To Len(Me.diceusStr)
                    If Mid(Me.diceusStr, i, 1) = "1" Then
                        nm1 = nm1 + 1
                    End If
                Next
                For i = 1 To Len(Me.dicecomStr)
                    If Mid(Me.dicecomStr, i, 1) = "1" Then
                        nm2 = nm2 + 1
                    End If
                Next
                Me.diceusTotal = Len(Me.diceusStr)
                Me.diceusTrue = nm1
                Me.dicecomTotal = Len(Me.dicecomStr)
                Me.dicecomTrue = nm2
            Case 2
                For i = 1 To Me.diceusTotal
                    Randomize Timer
                    k = Int(Rnd() * 6) + 1
                    If k = 1 Or k = 6 Then
                        nm1 = nm1 + 1
                        nmstr1 = nmstr1 & "1"
                    Else
                        nmstr1 = nmstr1 & "0"
                    End If
                Next
                For i = 1 To Me.dicecomTotal
                    Randomize Timer
                    k = Int(Rnd() * 6) + 1
                    If k = 1 Or k = 6 Then
                        nm2 = nm2 + 1
                        nmstr2 = nmstr2 & "1"
                    Else
                        nmstr2 = nmstr2 & "0"
                    End If
                Next
                Me.diceusStr = nmstr1
                Me.diceusTrue = nm1
                Me.dicecomStr = nmstr2
                Me.dicecomTrue = nm2
        End Select
        '==========================
        If Me.diceusTotal > 120 Then
            For i = 1 To 120
                Load dicemini1(i)
                dicemini1(i).Opacity = 0
                dicemini1(i).Visible = False
            Next
        Else
            For i = 1 To Me.diceusTotal
                Load dicemini1(i)
                dicemini1(i).Opacity = 0
                dicemini1(i).Visible = False
            Next
        End If
        For i = 1 To Me.diceusTrue
            Load dice1(i)
            dice1(i).Opacity = 0
            dice1(i).Visible = False
        Next
        If Me.dicecomTotal > 120 Then
            For i = 1 To 120
                Load dicemini2(i)
                dicemini2(i).Opacity = 0
                dicemini2(i).Visible = False
            Next
        Else
            For i = 1 To Me.dicecomTotal
                Load dicemini2(i)
                dicemini2(i).Opacity = 0
                dicemini2(i).Visible = False
            Next
        End If
        For i = 1 To Me.dicecomTrue
            Load dice2(i)
            dice2(i).Opacity = 0
            dice2(i).Visible = False
        Next
        For i = 1 To 4
            Load ucMusicPlayer(i)
            ucMusicPlayer(i).Filepath = App.Path & "\mp3\ulse07.mp3"
            ucMusicPlayer(i).IsLoop = False
        Next
        Call AdjustVolume
        '=========================
        Call 排列骰子版面
        trpersonshowoff.Enabled = True
    End If
End Property
Public Property Get DiceStop() As Boolean
    DiceStop = m_stop
End Property
Public Property Let DiceStop(ByVal New_DiceStop As Boolean)
    m_stop = New_DiceStop
    PropertyChanged "DiceStop"
    '=====================
    If Me.DiceStop = True Then
        Me.DiceStart = False
    End If
End Property
Public Property Get DiceATKMode() As Integer
    DiceATKMode = m_DiceATKMode
End Property
Public Property Let DiceATKMode(ByVal New_DiceATKMode As Integer)
    m_DiceATKMode = New_DiceATKMode
    PropertyChanged "DiceATKMode"
End Property
Public Property Get DiceInputMode() As Integer
    DiceInputMode = m_DiceInputMode
End Property
Public Property Let DiceInputMode(ByVal New_DiceInputMode As Integer)
    m_DiceInputMode = New_DiceInputMode
    PropertyChanged "DiceInputMode"
End Property
Public Property Get dicevoice() As Integer
    dicevoice = m_dicevoice
End Property
Public Property Let dicevoice(ByVal New_dicevoice As Integer)
    If New_dicevoice + 10 <= 100 And New_dicevoice + 10 >= 0 Then
        m_dicevoice = New_dicevoice + 10
    Else
        m_dicevoice = 100
    End If
    PropertyChanged "dicevoice"
End Property
Public Property Get PersonImage() As String
    PersonImage = m_PersonImage
End Property
Public Property Let PersonImage(ByVal New_PersonImage As String)
    m_PersonImage = New_PersonImage
    PropertyChanged "PersonImage"
    '=====================
    person.ClearImage
    person.ScaleMethod = aiActualSize
    person.AutoSize = True
    person.LoadImage_FromFile Me.PersonImage
    person.AutoSize = False
    person.Height = 5190
    person.Top = 960
End Property
Public Property Get PersonImageLeftZero() As Boolean
    PersonImageLeftZero = m_PersonImageLeftZero
End Property
Public Property Let PersonImageLeftZero(ByVal New_PersonImageLeftZero As Boolean)
    m_PersonImageLeftZero = New_PersonImageLeftZero
    PropertyChanged "PersonImageLeftZero"
End Property

Private Sub trdiceoff_all_Timer()
    Dim i As Integer
    If TrDiceOffAllNum1 >= 130 Then
        If TrDiceOffAllNum2 >= 10 And TrDiceOffAllNum3 >= 10 Then
            trdiceoff_all.Enabled = False
            TrPersonShowOffNum = 2
            trpersonshowoff.Enabled = True
        Else
            For i = 1 To dicemini1.UBound
                dicemini1(i).Opacity = dicemini1(i).Opacity - 10
                If dicemini1(i).Opacity = 0 Then
                    dicemini1(i).Visible = False
                End If
            Next
            For i = 1 To dicemini2.UBound
                dicemini2(i).Opacity = dicemini2(i).Opacity - 10
                If dicemini2(i).Opacity = 0 Then
                    dicemini2(i).Visible = False
                End If
            Next
            For i = 1 To Me.diceusTrue
                dice1(i).Opacity = dice1(i).Opacity - 10
                If dice1(i).Opacity = 0 Then
                    dice1(i).Visible = False
                End If
            Next
            For i = 1 To Me.dicecomTrue
                dice2(i).Opacity = dice2(i).Opacity - 10
                If dice2(i).Opacity = 0 Then
                    dice2(i).Visible = False
                End If
            Next
            TrDiceOffAllNum2 = TrDiceOffAllNum2 + 1
            TrDiceOffAllNum3 = TrDiceOffAllNum3 + 1
            TrDiceOffAllNum1 = TrDiceOffAllNum1 + 1
        End If
    Else
        TrDiceOffAllNum1 = TrDiceOffAllNum1 + 1
    End If
End Sub

Private Sub trdiceoff_tot_Timer()
    Dim i As Integer
    If TrDiceOffTotNum1 = 100 Then
        TrDiceOnTrueNum = 0
        trdiceon_true.Enabled = True
        TrDiceOffTotNum1 = TrDiceOffTotNum1 + 1
    ElseIf TrDiceOffTotNum1 > 100 Then
        If TrDiceOffTotNum2 >= 5 Then
            trdiceoff_tot.Enabled = False
        Else
            For i = 1 To dicemini1.UBound
                dicemini1(i).Opacity = dicemini1(i).Opacity - 10
            Next
            For i = 1 To dicemini2.UBound
                dicemini2(i).Opacity = dicemini2(i).Opacity - 10
            Next
            TrDiceOffTotNum2 = TrDiceOffTotNum2 + 1
            TrDiceOffTotNum1 = TrDiceOffTotNum1 + 1
        End If
    Else
        TrDiceOffTotNum1 = TrDiceOffTotNum1 + 1
    End If
End Sub

Private Sub trdiceoff_true_Timer()
    Dim i As Integer
    If TrDiceOffTrueNum2 >= 100 Then
        TrDiceOffTrueNum2 = 0
        TrDiceOffTrueNum1 = TrDiceOffTrueNum1 + 1
    Else
        If TrDiceOffTrueNum1 <= Me.diceusTrue Then
            If TrDiceOffTrueNum2 = 90 Then
                dice1(TrDiceOffTrueNum1).Opacity = dice1(TrDiceOffTrueNum1).Opacity - 10
            Else
                dice1(TrDiceOffTrueNum1).Opacity = dice1(TrDiceOffTrueNum1).Opacity - 15
            End If
        End If
        If TrDiceOffTrueNum1 <= Me.dicecomTrue Then
            If TrDiceOffTrueNum2 = 90 Then
                dice2(TrDiceOffTrueNum1).Opacity = dice2(TrDiceOffTrueNum1).Opacity - 10
            Else
                dice2(TrDiceOffTrueNum1).Opacity = dice2(TrDiceOffTrueNum1).Opacity - 15
            End If
        End If
        TrDiceOffTrueNum2 = TrDiceOffTrueNum2 + 15
    End If
    If TrDiceOffTrueNum1 > Me.diceusTrue Or TrDiceOffTrueNum1 > Me.dicecomTrue Then
        trdiceoff_true.Enabled = False
        TrDiceOffAllNum1 = 0
        TrDiceOffAllNum2 = 1
        TrDiceOffAllNum3 = 0
        trdiceoff_all.Enabled = True
    End If
End Sub

Private Sub trdiceon_true_Timer()
    Dim i As Integer

    If TrDiceOnTrueNum >= 10 Then
        trdiceon_true.Enabled = False
        TrDiceOffTrueNum1 = 1
        TrDiceOffTrueNum2 = 0
        trdiceoff_true.Enabled = True
    Else
        For i = 1 To Me.diceusTrue
            dice1(i).Opacity = dice1(i).Opacity + 20
        Next
        For i = 1 To Me.dicecomTrue
            dice2(i).Opacity = dice2(i).Opacity + 20
        Next
        TrDiceOnTrueNum = TrDiceOnTrueNum + 2
    End If
End Sub

Private Sub trdiceshow_tot_Timer()
    Dim i As Integer
    Dim CurState As Long
    If TrDiceShowTotalNum2 >= 100 Then
        TrDiceShowTotalNum2 = 0
        '====================
        If (TrDiceShowTotalNum1 + 3 <= Me.diceusTotal And Me.diceusTotal >= Me.dicecomTotal) Or _
           (TrDiceShowTotalNum1 + 3 <= Me.dicecomTotal And Me.diceusTotal < Me.dicecomTotal) Or _
           TrDiceShowTotalNum1 = 0 Then
            ucMusicPlayer((Val(TrDiceShowTotalNum1) Mod 4) + 1).MusicStop
            ucMusicPlayer((Val(TrDiceShowTotalNum1) Mod 4) + 1).MusicPlay
        End If
        '====================
        TrDiceShowTotalNum1 = TrDiceShowTotalNum1 + 1
    Else
        If TrDiceShowTotalNum1 <= Me.diceusTotal And TrDiceShowTotalNum1 <= 120 Then
            dicemini1(TrDiceShowTotalNum1).Opacity = dicemini1(TrDiceShowTotalNum1).Opacity + 50
        End If
        If TrDiceShowTotalNum1 <= Me.dicecomTotal And TrDiceShowTotalNum1 <= 120 Then
            dicemini2(TrDiceShowTotalNum1).Opacity = dicemini2(TrDiceShowTotalNum1).Opacity + 50
        End If
        TrDiceShowTotalNum2 = TrDiceShowTotalNum2 + 50
    End If
    If TrDiceShowTotalNum1 > Me.diceusTotal And TrDiceShowTotalNum1 > Me.dicecomTotal Then
        trdiceshow_tot.Enabled = False
        TrDiceOffTotNum1 = 0
        TrDiceOffTotNum2 = 0
        trdiceoff_tot.Enabled = True
    End If
End Sub

Private Sub trpersonshowoff_Timer()
    Dim kp As Integer, bigw As Integer, bigall As Integer
    '============================
    bigw = person.Width / 2
    Select Case Me.DiceATKMode
        Case 1
            If 2580 - bigw < 0 Or Me.PersonImageLeftZero = True Then
                bigall = 0
            Else
                bigall = 2580 - bigw
            End If
            kp = (person.Width + bigall) / 20
        Case 2
            If 8760 - bigw >= (11340 - person.Width) Or Me.PersonImageLeftZero = True Then
                bigall = 11340 - person.Width
            Else
                bigall = 8760 - bigw
            End If
            kp = (11340 - bigall) / 20
    End Select
    '============================
    Select Case TrPersonShowOffNum
        Case 1
            Select Case Me.DiceATKMode
                Case 1
                    If person.Left >= bigall Then
                        GoTo TrOff1
                    ElseIf Abs(bigall - person.Left) < kp Then
                        person.Left = person.Left + Abs(bigall - person.Left)
                    Else
                        person.Left = person.Left + kp
                    End If
                Case 2
                    If person.Left <= bigall Then
                        GoTo TrOff1
                    ElseIf (person.Left - bigall) < kp Then
                        person.Left = person.Left - (person.Left - bigall)
                    Else
                        person.Left = person.Left - kp
                    End If
            End Select
        Case 2
            Select Case Me.DiceATKMode
                Case 1
                    If person.Left <= -person.Width Then
                        GoTo TrOff2
                    Else
                        person.Left = person.Left - kp
                        If person.Opacity > 0 Then
                            person.Opacity = Val(person.Opacity) - 10
                        End If
                    End If
                Case 2
                    If person.Left >= 11340 Then
                        GoTo TrOff2
                    Else
                        person.Left = person.Left + kp
                        If person.Opacity > 0 Then
                            person.Opacity = Val(person.Opacity) - 10
                        End If
                    End If
            End Select
    End Select
    '===========================
    Exit Sub
TrOff1:
    trpersonshowoff.Enabled = False
    trdiceshow_tot.Enabled = True
    '===========================
    Exit Sub
TrOff2:
    trpersonshowoff.Enabled = False
    Call 擲骰物件卸載
    Me.DiceStop = True
End Sub

Private Sub 排列骰子版面()
    Dim i As Integer
    For i = 1 To Me.diceusTotal
        If i Mod 10 = 1 Then
            dicemini1(i).Top = Val(5040) + Val((i \ 10)) * Val(410)
            dicemini1(i).Left = 3720
        Else
            If i Mod 10 = 0 Then
                dicemini1(i).Left = 3720 + 9 * 411
                dicemini1(i).Top = Val(5040) + Val(((i \ 10) - 1)) * Val(410)
            Else
                dicemini1(i).Left = 3720 + ((i Mod 10) - 1) * 411
                dicemini1(i).Top = Val(5040) + Val((i \ 10)) * Val(410)
            End If
        End If
        Select Case Me.DiceATKMode
            Case 1
                dicemini1(i).ClearImage
                Select Case Val(Mid(Me.diceusStr, i, 1))
                    Case 0
                        dicemini1(i).LoadImage_FromFile App.Path & "\gif\system\atknothing.png"
                    Case 1
                        dicemini1(i).LoadImage_FromFile App.Path & "\gif\system\atkshow.png"
                End Select
            Case 2
                dicemini1(i).ClearImage
                Select Case Val(Mid(Me.diceusStr, i, 1))
                    Case 0
                        dicemini1(i).LoadImage_FromFile App.Path & "\gif\system\defnothing.png"
                    Case 1
                        dicemini1(i).LoadImage_FromFile App.Path & "\gif\system\defshow.png"
                End Select
        End Select
        dicemini1(i).ScaleMethod = aiStretch
        dicemini1(i).Width = 375
        dicemini1(i).Height = 375
        dicemini1(i).Opacity = 0
        dicemini1(i).Visible = True
        dicemini1(i).ZOrder
    Next
    For i = 1 To Me.dicecomTotal
        If i Mod 10 = 1 Then
            dicemini2(i).Top = Val(4440) - Val((i \ 10)) * Val(410)
            dicemini2(i).Left = 3720
        Else
            If i Mod 10 = 0 Then
                dicemini2(i).Left = 3720 + 9 * 411
                dicemini2(i).Top = Val(4440) - Val(((i \ 10) - 1)) * Val(410)
            Else
                dicemini2(i).Left = 3720 + ((i Mod 10) - 1) * 411
                dicemini2(i).Top = Val(4440) - Val((i \ 10)) * Val(410)
            End If
        End If
        Select Case Me.DiceATKMode
            Case 2
                dicemini2(i).ClearImage
                Select Case Val(Mid(Me.dicecomStr, i, 1))
                    Case 0
                        dicemini2(i).LoadImage_FromFile App.Path & "\gif\system\atknothing.png"
                    Case 1
                        dicemini2(i).LoadImage_FromFile App.Path & "\gif\system\atkshow.png"
                End Select
            Case 1
                dicemini2(i).ClearImage
                Select Case Val(Mid(Me.dicecomStr, i, 1))
                    Case 0
                        dicemini2(i).LoadImage_FromFile App.Path & "\gif\system\defnothing.png"
                    Case 1
                        dicemini2(i).LoadImage_FromFile App.Path & "\gif\system\defshow.png"
                End Select
        End Select
        dicemini2(i).ScaleMethod = aiStretch
        dicemini2(i).Width = 375
        dicemini2(i).Height = 375
        dicemini2(i).Opacity = 0
        dicemini2(i).Visible = True
        dicemini2(i).ZOrder
    Next
    '======================================
    For i = 1 To Me.diceusTrue
        If i Mod 5 = 1 Then
            dice1(i).Top = Val(5040) + Val((i \ 5)) * Val(840)
            dice1(i).Left = 3720
        Else
            If i Mod 5 = 0 Then
                dice1(i).Left = Val(3720) + Val(4) * Val(840)
                dice1(i).Top = Val(5040) + Val(((i \ 5) - 1)) * Val(840)
            Else
                dice1(i).Left = Val(3720) + Val(((i Mod 5) - 1)) * Val(840)
                dice1(i).Top = Val(5040) + Val((i \ 5)) * Val(840)
            End If
        End If
        Select Case Me.DiceATKMode
            Case 1
                dice1(i).ClearImage
                dice1(i).LoadImage_FromFile App.Path & "\gif\system\atkshow.png"
            Case 2
                dice1(i).ClearImage
                dice1(i).LoadImage_FromFile App.Path & "\gif\system\defshow.png"
        End Select
        dice1(i).Opacity = 0
        dice1(i).Visible = True
        dice1(i).ZOrder
    Next
    For i = 1 To Me.dicecomTrue
        If i Mod 5 = 1 Then
            dice2(i).Top = Val(4080) - Val((i \ 5)) * Val(840)
            dice2(i).Left = 3720
        Else
            If i Mod 5 = 0 Then
                dice2(i).Left = 3720 + 4 * 840
                dice2(i).Top = Val(4080) - Val(((i \ 5) - 1)) * Val(840)
            Else
                dice2(i).Left = 3720 + ((i Mod 5) - 1) * 840
                dice2(i).Top = Val(4080) - Val((i \ 5)) * Val(840)
            End If
        End If
        Select Case Me.DiceATKMode
            Case 2
                dice2(i).ClearImage
                dice2(i).LoadImage_FromFile App.Path & "\gif\system\atkshow.png"
            Case 1
                dice2(i).ClearImage
                dice2(i).LoadImage_FromFile App.Path & "\gif\system\defshow.png"
        End Select
        dice2(i).Opacity = 0
        dice2(i).Visible = True
        dice2(i).ZOrder
    Next
End Sub
Private Sub 擲骰物件卸載()
    Dim i As Integer
    For i = 1 To dicemini1.UBound
        Unload dicemini1(i)
    Next
    For i = 1 To dice1.UBound
        Unload dice1(i)
    Next
    For i = 1 To dicemini2.UBound
        Unload dicemini2(i)
    Next
    For i = 1 To dice2.UBound
        Unload dice2(i)
    Next
    For i = 1 To ucMusicPlayer.UBound
        Unload ucMusicPlayer(i)
    Next
End Sub
Private Sub AdjustVolume()
    Dim i As Integer
    For i = 1 To ucMusicPlayer.UBound
        ucMusicPlayer(i).Volume = Me.dicevoice
    Next
End Sub
