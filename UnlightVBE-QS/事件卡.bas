Attribute VB_Name = "事件卡"
Option Explicit
Public 事件卡記錄暫時數(0 To 2, 1 To 6) As Integer '事件卡使用紀錄暫時變數(0.(1)總共給予回合數,1.使用者/2.電腦,1.總共數值/2.目前處理數值/3.目前階段/4.事件卡牌編號/5.事件分類/6.是否啟動)
Sub 機會_使用者(ByVal num As Integer, ByVal tot As Integer)
Select Case 事件卡記錄暫時數(1, 3)
    Case 1
        目前數(15) = 7
        事件卡記錄暫時數(1, 4) = num
        事件卡記錄暫時數(1, 1) = tot
        事件卡記錄暫時數(1, 5) = 1
        事件卡記錄暫時數(1, 6) = 1
        FormMainMode.對齊完成檢查.Enabled = False
    Case 2
        一般系統類.音效播放 7
        '=============以下是牌移動(收牌)(使用者)
'         戰鬥系統類.座標計算_使用者手牌
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(1, 4)
         目前數(5) = pagecardnum(事件卡記錄暫時數(1, 4), 7)
         pagecardnum(事件卡記錄暫時數(1, 4), 9) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(1, 4), 10) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 6
         FormMainMode.牌移動.Enabled = True
        '================以下是出牌對齊
        目前數(3) = 0
        戰鬥系統類.出牌順序計算_使用者_出牌
        FormMainMode.使用者出牌_出牌對齊_靠右.Enabled = True
        '=====================
        事件卡記錄暫時數(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        If BattleCardNum < 事件卡記錄暫時數(1, 1) Then
           戰鬥系統類.執行動作_洗牌
        End If
    Case 3
         If 事件卡記錄暫時數(1, 2) > 事件卡記錄暫時數(1, 1) Or BattleCardNum <= 0 Then
             turnpageonin = 1
             FormMainMode.PEAFInterface.BnOKStartListen
             事件卡記錄暫時數(1, 6) = 0
             Exit Sub
         End If
         Do Until 事件卡記錄暫時數(1, 2) > 事件卡記錄暫時數(1, 1)
             目前數(15) = 8
             FormMainMode.tr牌組_抽牌_使用者.Enabled = True
             事件卡記錄暫時數(1, 2) = 事件卡記錄暫時數(1, 2) + 1
             Exit Do
         Loop
End Select
End Sub
Sub 機會_電腦(ByVal num As Integer, ByVal tot As Integer)
Select Case 事件卡記錄暫時數(2, 3)
    Case 1
        目前數(15) = 9
        目前數(17) = 2
        事件卡記錄暫時數(2, 4) = num
        事件卡記錄暫時數(2, 1) = tot
        事件卡記錄暫時數(2, 5) = 1
        事件卡記錄暫時數(2, 6) = 1
    Case 2
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Width = 810
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Height = 1260
        FormMainMode.card(事件卡記錄暫時數(2, 4)).cardImage = app_path & "card\" & pagecardnum(事件卡記錄暫時數(2, 4), 8) & ".png"
        FormMainMode.card(事件卡記錄暫時數(2, 4)).CardRotationType = pageonin(事件卡記錄暫時數(2, 4))
        一般系統類.音效播放 7
        等待時間佇列(2).Add 9
        FormMainMode.等待時間_2.Enabled = True
    Case 3
        '=============以下是牌移動(收牌)(電腦)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(2, 4)
         pagecardnum(事件卡記錄暫時數(2, 4), 9) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(2, 4), 10) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 10
         FormMainMode.牌移動.Enabled = True
        '=====================
        事件卡記錄暫時數(2, 2) = 1
        If BattleCardNum < 事件卡記錄暫時數(2, 1) Then
           戰鬥系統類.執行動作_洗牌
        End If
    Case 4
         If 事件卡記錄暫時數(2, 2) > 事件卡記錄暫時數(2, 1) Or BattleCardNum <= 0 Then
             等待時間佇列(2).Add 10
             FormMainMode.等待時間_2.Enabled = True
             事件卡記錄暫時數(2, 6) = 0
             Exit Sub
         End If
         Do Until 事件卡記錄暫時數(2, 2) > 事件卡記錄暫時數(2, 1)
             目前數(15) = 11
             FormMainMode.tr牌組_抽牌_電腦.Enabled = True
             事件卡記錄暫時數(2, 2) = 事件卡記錄暫時數(2, 2) + 1
             Exit Do
         Loop
End Select
End Sub
Sub 詛咒術_使用者(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case 事件卡記錄暫時數(1, 3)
    Case 1
        目前數(15) = 12
        事件卡記錄暫時數(1, 4) = num
        事件卡記錄暫時數(1, 1) = tot
        事件卡記錄暫時數(1, 5) = 2
        事件卡記錄暫時數(1, 6) = 1
        FormMainMode.對齊完成檢查.Enabled = False
    Case 2
        事件卡記錄暫時數(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        '=======================
        Do Until 事件卡記錄暫時數(1, 2) > 事件卡記錄暫時數(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * 公用牌實體卡片分隔紀錄數(1)) + 1
            If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                目前數(17) = 6
                目前數(16) = m
                事件卡記錄暫時數(1, 2) = 事件卡記錄暫時數(1, 2) + 1
                FormMainMode.tr電腦牌_翻牌.Enabled = True
                Exit Sub
            End If
        Loop
        If 事件卡記錄暫時數(1, 2) > 事件卡記錄暫時數(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0 Then
            等待時間佇列(2).Add 12
            FormMainMode.等待時間_2.Enabled = True
        End If
     Case 3
        Do Until 事件卡記錄暫時數(1, 2) > 事件卡記錄暫時數(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * 公用牌實體卡片分隔紀錄數(1)) + 1
            If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                目前數(17) = 6
                目前數(16) = m
                事件卡記錄暫時數(1, 2) = 事件卡記錄暫時數(1, 2) + 1
                FormMainMode.tr電腦牌_翻牌.Enabled = True
                Exit Sub
            End If
        Loop
        If 事件卡記錄暫時數(1, 2) > 事件卡記錄暫時數(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0 Then
            等待時間佇列(2).Add 12
            FormMainMode.等待時間_2.Enabled = True
        End If
     Case 4
        FormMainMode.tr電腦牌_棄牌.Enabled = True
     Case 5
         一般系統類.音效播放 7
        '=============以下是牌移動(收牌)(使用者)
'         戰鬥系統類.座標計算_使用者手牌
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(1, 4)
         目前數(5) = pagecardnum(事件卡記錄暫時數(1, 4), 7)
         pagecardnum(事件卡記錄暫時數(1, 4), 9) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(1, 4), 10) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 13
         FormMainMode.牌移動.Enabled = True
        '================以下是出牌對齊
        目前數(3) = 0
        戰鬥系統類.出牌順序計算_使用者_出牌
        FormMainMode.使用者出牌_出牌對齊_靠右.Enabled = True
        '=====================
        事件卡記錄暫時數(1, 2) = 1
    Case 6
        turnpageonin = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        事件卡記錄暫時數(1, 6) = 0
End Select
End Sub
Sub 詛咒術_電腦(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case 事件卡記錄暫時數(2, 3)
    Case 1
        目前數(15) = 14
        目前數(17) = 2
        事件卡記錄暫時數(2, 4) = num
        事件卡記錄暫時數(2, 1) = tot
        事件卡記錄暫時數(2, 5) = 2
        事件卡記錄暫時數(2, 6) = 1
    Case 2
        事件卡記錄暫時數(2, 2) = 1
        '=======================
        Do Until 事件卡記錄暫時數(2, 2) > 事件卡記錄暫時數(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * 公用牌實體卡片分隔紀錄數(1)) + 1
            If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                目前數(21) = 3
                目前數(20) = m
                事件卡記錄暫時數(2, 2) = 事件卡記錄暫時數(2, 2) + 1
                FormMainMode.tr使用者_棄牌.Enabled = True
                Exit Sub
            End If
        Loop
        If 事件卡記錄暫時數(2, 2) > 事件卡記錄暫時數(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
            等待時間佇列(2).Add 14
            FormMainMode.等待時間_2.Enabled = True
        End If
     Case 3
        Do Until 事件卡記錄暫時數(2, 2) > 事件卡記錄暫時數(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * 公用牌實體卡片分隔紀錄數(1)) + 1
            If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                目前數(21) = 3
                目前數(20) = m
                事件卡記錄暫時數(2, 2) = 事件卡記錄暫時數(2, 2) + 1
                FormMainMode.tr使用者_棄牌.Enabled = True
                Exit Sub
            End If
        Loop
        If 事件卡記錄暫時數(2, 2) > 事件卡記錄暫時數(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
            等待時間佇列(2).Add 14
            FormMainMode.等待時間_2.Enabled = True
        End If
     Case 4
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Width = 810
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Height = 1260
        FormMainMode.card(事件卡記錄暫時數(2, 4)).cardImage = app_path & "card\" & pagecardnum(事件卡記錄暫時數(2, 4), 8) & ".png"
        FormMainMode.card(事件卡記錄暫時數(2, 4)).CardRotationType = pageonin(事件卡記錄暫時數(2, 4))
        一般系統類.音效播放 7
        等待時間佇列(2).Add 15
        FormMainMode.等待時間_2.Enabled = True
     Case 5
        '=============以下是牌移動(收牌)(電腦)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(2, 4)
         pagecardnum(事件卡記錄暫時數(2, 4), 9) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(2, 4), 10) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 15
         FormMainMode.牌移動.Enabled = True
        '=====================
    Case 6
        等待時間佇列(2).Add 10
        FormMainMode.等待時間_2.Enabled = True
        事件卡記錄暫時數(2, 6) = 0
End Select
End Sub
Sub HP回復_使用者(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case 事件卡記錄暫時數(1, 3)
    Case 1
        目前數(15) = 16
        事件卡記錄暫時數(1, 4) = num
        事件卡記錄暫時數(1, 1) = tot
        事件卡記錄暫時數(1, 5) = 3
        事件卡記錄暫時數(1, 6) = 1
        FormMainMode.對齊完成檢查.Enabled = False
    Case 2
        事件卡記錄暫時數(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        '=======================
        戰鬥系統類.回復執行_使用者 Val(事件卡記錄暫時數(1, 1)), 1, 0, True
        等待時間佇列(2).Add 17
        FormMainMode.等待時間_2.Enabled = True
     Case 3
         一般系統類.音效播放 7
        '=============以下是牌移動(收牌)(使用者)
'         戰鬥系統類.座標計算_使用者手牌
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(1, 4)
         目前數(5) = pagecardnum(事件卡記錄暫時數(1, 4), 7)
         pagecardnum(事件卡記錄暫時數(1, 4), 9) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(1, 4), 10) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 17
         FormMainMode.牌移動.Enabled = True
        '================以下是出牌對齊
        目前數(3) = 0
        戰鬥系統類.出牌順序計算_使用者_出牌
        FormMainMode.使用者出牌_出牌對齊_靠右.Enabled = True
        '=====================
    Case 4
        turnpageonin = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        事件卡記錄暫時數(1, 6) = 0
End Select
End Sub
Sub HP回復_電腦(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case 事件卡記錄暫時數(2, 3)
    Case 1
        目前數(15) = 18
        目前數(17) = 2
        事件卡記錄暫時數(2, 4) = num
        事件卡記錄暫時數(2, 1) = tot
        事件卡記錄暫時數(2, 5) = 3
        事件卡記錄暫時數(2, 6) = 1
    Case 2
        事件卡記錄暫時數(2, 2) = 1
        '=======================
        回復執行_電腦 Val(事件卡記錄暫時數(2, 1)), 1, 0, True
        等待時間佇列(2).Add 19
        FormMainMode.等待時間_2.Enabled = True
     Case 3
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Width = 810
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Height = 1260
        FormMainMode.card(事件卡記錄暫時數(2, 4)).cardImage = app_path & "card\" & pagecardnum(事件卡記錄暫時數(2, 4), 8) & ".png"
        FormMainMode.card(事件卡記錄暫時數(2, 4)).CardRotationType = pageonin(事件卡記錄暫時數(2, 4))
        一般系統類.音效播放 7
        等待時間佇列(2).Add 20
        FormMainMode.等待時間_2.Enabled = True
     Case 4
        '=============以下是牌移動(收牌)(電腦)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(2, 4)
         pagecardnum(事件卡記錄暫時數(2, 4), 9) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(2, 4), 10) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 19
         FormMainMode.牌移動.Enabled = True
        '=====================
    Case 5
        等待時間佇列(2).Add 10
        FormMainMode.等待時間_2.Enabled = True
        事件卡記錄暫時數(2, 6) = 0
End Select
End Sub
Sub 聖水_使用者(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case 事件卡記錄暫時數(1, 3)
    Case 1
        目前數(15) = 41
        事件卡記錄暫時數(1, 4) = num
        事件卡記錄暫時數(1, 1) = tot
        事件卡記錄暫時數(1, 5) = 4
        事件卡記錄暫時數(1, 6) = 1
        FormMainMode.對齊完成檢查.Enabled = False
    Case 2
        事件卡記錄暫時數(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        '=======================
        戰鬥系統類.執行動作_清除所有異常狀態_聖水 1, 1
        戰鬥系統類.骰量更新顯示
        FormMainMode.trgoi1.Enabled = True
        等待時間佇列(2).Add 40
        FormMainMode.等待時間_2.Enabled = True
     Case 3
         一般系統類.音效播放 7
        '=============以下是牌移動(收牌)(使用者)
'         戰鬥系統類.座標計算_使用者手牌
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(1, 4)
         目前數(5) = pagecardnum(事件卡記錄暫時數(1, 4), 7)
         pagecardnum(事件卡記錄暫時數(1, 4), 9) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(1, 4), 10) = FormMainMode.card(事件卡記錄暫時數(1, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 17
         FormMainMode.牌移動.Enabled = True
        '================以下是出牌對齊
        目前數(3) = 0
        戰鬥系統類.出牌順序計算_使用者_出牌
        FormMainMode.使用者出牌_出牌對齊_靠右.Enabled = True
        '=====================
    Case 4
        turnpageonin = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        事件卡記錄暫時數(1, 6) = 0
End Select
End Sub
Sub 聖水_電腦(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case 事件卡記錄暫時數(2, 3)
    Case 1
        目前數(15) = 43
        目前數(17) = 2
        事件卡記錄暫時數(2, 4) = num
        事件卡記錄暫時數(2, 1) = tot
        事件卡記錄暫時數(2, 5) = 3
        事件卡記錄暫時數(2, 6) = 1
    Case 2
        事件卡記錄暫時數(2, 2) = 1
        '=======================
        戰鬥系統類.執行動作_清除所有異常狀態_聖水 2, 1
        戰鬥系統類.骰量更新顯示
        FormMainMode.trgoi2.Enabled = True
        等待時間佇列(2).Add 42
        FormMainMode.等待時間_2.Enabled = True
     Case 3
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Width = 810
        FormMainMode.card(事件卡記錄暫時數(2, 4)).Height = 1260
        FormMainMode.card(事件卡記錄暫時數(2, 4)).cardImage = app_path & "card\" & pagecardnum(事件卡記錄暫時數(2, 4), 8) & ".png"
        FormMainMode.card(事件卡記錄暫時數(2, 4)).CardRotationType = pageonin(事件卡記錄暫時數(2, 4))
        一般系統類.音效播放 7
        等待時間佇列(2).Add 43
        FormMainMode.等待時間_2.Enabled = True
     Case 4
        '=============以下是牌移動(收牌)(電腦)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         牌移動暫時變數(1) = 240
         牌移動暫時變數(2) = 960
         牌移動暫時變數(3) = 事件卡記錄暫時數(2, 4)
         pagecardnum(事件卡記錄暫時數(2, 4), 9) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Left  '指定目前Left(座標)
         pagecardnum(事件卡記錄暫時數(2, 4), 10) = FormMainMode.card(事件卡記錄暫時數(2, 4)).Top  '指定目前Top(座標)
         戰鬥系統類.計算牌移動距離單位
         目前數(15) = 44
         FormMainMode.牌移動.Enabled = True
        '=====================
    Case 5
        等待時間佇列(2).Add 10
        FormMainMode.等待時間_2.Enabled = True
        事件卡記錄暫時數(2, 6) = 0
End Select
End Sub

