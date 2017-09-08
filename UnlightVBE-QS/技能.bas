Attribute VB_Name = "技能"
Public atking_sheri_4_tot As Integer   '技能-雪莉-飛刃雨出牌量儲存變數
Public atking_sheri_4_tot_ai As Integer  '技能-AI-雪莉-飛刃雨出牌量儲存變數
Public atking_帕茉_慈悲的藍眼_tot(1 To 2) As Integer  '技能-帕茉-慈悲的藍眼骰子量紀錄暫時變數(1.數值/2.是否啟動)
Public atking_艾茵_十三隻眼_tot(1 To 2) As Integer '技能.艾茵_十三隻眼骰子量紀錄暫時變數(1.數值/2.是否啟動)
Public atking_史塔夏_殺戮模式狀態數(1 To 5) As Integer '史塔夏殺戮模式狀態檢查數(1.狀態執行階段/2.狀態啟動檢查值/3.紀錄數值(原始)/4.紀錄數值(變更後)/5.數值紀錄是否啟動)
Public atking_音音夢_成長模式狀態數(1 To 2) As Integer '音音夢成長模式狀態檢查數(1.狀態執行階段/2.狀態啟動檢查值)
Public atking_蕾_守護模式狀態啟動值 As Boolean '技能-蕾-Ex-協奏曲-加百烈的守護免除直傷模式啟動值
Public atking_蕾_終曲_無盡輪迴的終結紀錄數 As Integer  '技能-蕾-Ex-終曲-無盡輪迴的終結紀錄對手之防禦牌值暫時數
Public atking_羅莎琳_黑霧幻影紀錄狀態數(1 To 106) As Boolean '技能-羅莎琳-黑霧幻影(普、EX)紀錄對手出牌編號數
Public atking_伊芙琳_怠惰的墓表紀錄數(0 To 2) As Integer '技能-伊芙琳-怠惰的墓表紀錄對手牌編號暫時數(0.總共張數值/1~2牌編號)
Public atking_伊芙琳_赤紅石榴階段紀錄數(0 To 106, 1 To 4) As Integer '技能-伊芙琳-赤紅石榴紀錄效果及階段暫時數(0.(1).當前效果/(2).當前效果階段/(3)總共抽牌數量/(4)目前抽/棄牌數量,1~106.(1)牌號選定紀錄值)
Public atking_古魯瓦爾多_精神力吸收紀錄數(0 To 106) As Integer '技能-古魯瓦爾多-精神力吸收紀錄對手牌編號暫時數(0.總共張數值/1~106牌編號選擇值)
Public atking_梅倫_Jackpot紀錄數(1 To 2) As Integer '技能-梅倫-Jackpot抽牌紀錄數(1.總共數/2.目前數)
Public atking_艾伯李斯特_雷擊紀錄數(1 To 2) As Integer '技能-艾伯李斯特-雷擊丟棄對手牌紀錄數(1.總共數/2.目前數)
Public atking_艾伯李斯特_智略紀錄數 As Integer '技能-艾伯李斯特-智略抽牌目前數
Public atking_艾依查庫_神速之劍計算數值紀錄數(1 To 2) As Integer  '技能-艾依查庫-神速之劍計算劍數值紀錄暫時數(1.目前計算數值/2.(廢除))
Public atking_布勞_發條機構紀錄數 As Integer '技能-布勞-發條機構抽牌目前數
Public atking_利恩_反擊的狼煙紀錄數(1 To 2) As Integer '技能-利恩-反擊的狼煙抽牌目前數(1.總共數/2.目前數)
Public atking_夏洛特_大聖堂骰量紀錄數(1 To 3) As Integer '技能-夏洛特-大聖堂擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後結果)
Public atking_瑪格莉特_月光紀錄數(0 To 107) As Integer '技能-瑪格莉特-月光紀錄對手牌編號暫時數(0.目前丟棄張數值/1~106牌編號選擇值/107.總共能丟棄張數值)
Public atking_庫勒尼西_瘋狂眼窩紀錄數 As Integer '技能-庫勒尼西-瘋狂眼窩丟棄對手牌紀錄目前數
Public atking_傑多_因果之刻記錄數(1 To 108) As Integer '技能-傑多-因果之刻紀錄對手出牌編號數(1~106.記錄牌編號/107.總共回張數/108.目前數)
Public atking_傑多_因果之幻骰量紀錄數(1 To 3) As Integer '技能-傑多-因果之幻擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後結果)
Public atking_阿奇波爾多_防護射擊_槍數值紀錄數 As Integer '技能-阿奇波爾多-防護射擊目前累計加槍數值紀錄數
Public atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1 To 2) As Integer '技能-洛洛妮-逆轉戰局的槍響抽牌目前數(1.總共數/2.目前數)
Public atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 As Integer '技能-洛洛妮-貪婪之刃與嗜血之槍搶牌目前數
Public atking_克頓_竊取資料_奪牌紀錄數(1 To 2) As Integer  '技能-克頓-竊取資料奪取對手出牌牌號紀錄數(1.奪牌編號/2.奪牌原方出牌順序)
Public atking_克頓_隱蔽射擊骰量紀錄數(1 To 3) As Integer '技能-克頓-隱蔽射擊擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後總結果)
Public atking_克頓_惡意情報紀錄數(0 To 106) As Integer '技能-克頓-惡意情報紀錄對手牌編號暫時數(0.目前階段/1~106牌編號選擇值)
Public atking_露緹亞_渦騎劍閃計算張數紀錄數 As Integer  '技能-露緹亞-渦騎劍閃計算劍卡張數值紀錄暫時數
Public atking_艾蕾可_王座之炎計算出牌張數紀錄數 As Integer  '技能-艾蕾可-王座之炎計算出牌張數值紀錄暫時數
Public atking_艾蕾可_聖王威光紀錄數(1 To 2) As Integer  '技能-艾蕾可-聖王威光紀錄暫時數(1.對手當回合防禦力/2.對手當回合出牌數/3.使用者當回合攻擊力)
Public atking_梅莉_綿羊幻夢_抽牌紀錄數(1 To 2) As Integer '技能-梅莉-綿羊幻夢抽牌目前數(1.總共數/2.目前數)
Public atking_貝琳達_雪光_抽牌紀錄數(1 To 2) As Integer '技能-貝琳達-雪光抽牌目前數(1.總共數/2.目前數)
Public atking_貝琳達_水晶幻鏡紀錄狀態數(1 To 106) As Boolean '技能-貝琳達-水晶幻鏡紀錄對手出牌編號數
Public atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1 To 2) As Boolean '技能-貝琳達-溶魂之雨攻擊力加成暫時紀錄數(1.是否10張已+10/2.是否15張已+15)
Public atking_尤莉卡_超載目前階段紀錄數(1 To 4)  As Integer  '技能-尤莉卡-超載執行目前階段數值紀錄暫時數(1.紀錄數值(原始)/2.紀錄數值(變更後)/3.目前執行階段(總)/4.超載3時攻防骰量加倍是否啟動)
Sub 雪莉_巨大黑犬()
Dim i As Integer, j As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "巨大黑犬" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(4, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
   Select Case atkingck(4, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(1, 1) >= 3 And atkingck(4, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(4, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 1) < 3 And atkingck(4, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(4, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\雪莉\雪莉_巨大黑犬_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9810
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 4
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
    Case 3
         Do
            atkingck(4, 2) = 0
            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                If 人物異常狀態資料庫(2, j, 3) = 5 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                 FormMainMode.personcomspe(j).person_num = 4
                 FormMainMode.personcomspe(j).person_turn = 3
                 人物異常狀態資料庫(2, j, 1) = 4
                 人物異常狀態資料庫(2, j, 2) = 3
                 Exit Do
                End If
            Next
           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 5, app_path & "gif\異常狀態\defdown.gif", 4, 3
                 異常狀態檢查數(5, 1) = 1
                 異常狀態檢查數(5, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If

End Sub
Sub 雪莉_自殺傾向(ByVal Index As Integer)
If FormMainMode.personatk(1).Caption = "自殺傾向" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(1, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
 Select Case atkingck(1, 1)
    Case 1
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2)) * 5
               If atkingck(1, 2) = 0 And atkingpagetot(1, 4) > 0 Then
                  戰鬥系統類.人物技能欄燈開關 True, 1
                  atkingck(1, 2) = 1
                  atkingtrn(1) = Val(atkingtrn(1)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 2)) * 5
               If atkingck(1, 2) = 1 And atkingpagetot(1, 4) = 0 Then
                  戰鬥系統類.人物技能欄燈開關 False, 1
                  atkingck(1, 2) = 0
                  atkingtrn(1) = Val(atkingtrn(1)) - 1
               End If
        End If
        FormMainMode.trgoi1.Enabled = True
    Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2)) * 5
               If atkingck(1, 2) = 0 And atkingpagetot(1, 4) > 0 Then
                  戰鬥系統類.人物技能欄燈開關 True, 1
                  atkingck(1, 2) = 1
                  atkingtrn(1) = Val(atkingtrn(1)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 4)) * 5
               If atkingck(1, 2) = 1 And atkingpagetot(1, 4) = 0 Then
                  戰鬥系統類.人物技能欄燈開關 False, 1
                  atkingck(1, 2) = 0
                  atkingtrn(1) = Val(atkingtrn(1)) - 1
               End If
        End If
        FormMainMode.trgoi1.Enabled = True
    Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_自殺傾向_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 0
                atkingno(i, 6) = 0
                atkingno(i, 7) = 1
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
        戰鬥系統類.人物技能欄燈開關 False, 1
       '-------------
    Case 4
       戰鬥系統類.傷害執行_技能直傷_使用者 Val(atkingpagetot(1, 4)), 1
       atkingck(1, 2) = 0
  End Select
End If
End Sub
Sub 雪莉_VBE_自殺傾向(ByVal Index As Integer)
Dim bloodnum As Integer '暫時變數
If FormMainMode.personatk(1).Caption = "VBE自殺傾向" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(42, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
 Select Case atkingck(42, 1)
    Case 1
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2)) * 10
               If atkingck(42, 2) = 0 And atkingpagetot(1, 4) > 0 Then
                  戰鬥系統類.人物技能欄燈開關 True, 1
                  atkingck(42, 2) = 1
                  atkingtrn(1) = Val(atkingtrn(1)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 2)) * 10
               If atkingck(42, 2) = 1 And atkingpagetot(1, 4) = 0 Then
                  戰鬥系統類.人物技能欄燈開關 False, 1
                  atkingck(42, 2) = 0
                  atkingtrn(1) = Val(atkingtrn(1)) - 1
               End If
        End If
        FormMainMode.trgoi1.Enabled = True
    Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2)) * 10
               If atkingck(42, 2) = 0 And atkingpagetot(1, 4) > 0 Then
                  戰鬥系統類.人物技能欄燈開關 True, 1
                  atkingck(42, 2) = 1
                  atkingtrn(1) = Val(atkingtrn(1)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 4)) * 10
               '顯示列1.goi1 = Val(顯示列1.goi1) - Val(pagecardnum(Index, 4)) * 5
               If atkingck(42, 2) = 1 And atkingpagetot(1, 4) = 0 Then
                  戰鬥系統類.人物技能欄燈開關 False, 1
                  atkingck(42, 2) = 0
                  atkingtrn(1) = Val(atkingtrn(1)) - 1
               End If
        End If
        FormMainMode.trgoi1.Enabled = True
    Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_自殺傾向_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 0
                atkingno(i, 6) = 0
                atkingno(i, 7) = 42
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
        戰鬥系統類.人物技能欄燈開關 False, 1
'        atkingck(42, 2) = 0
       '-------------
    Case 4
       bloodnum = Val(atkingpagetot(1, 4)) \ 2
       If bloodnum >= liveus(角色人物對戰人數(1, 2)) Then
           bloodnum = liveus(角色人物對戰人數(1, 2)) - 1
       End If
       戰鬥系統類.傷害執行_技能直傷_使用者 bloodnum, 1
       atkingck(42, 2) = 0
  End Select
End If
End Sub
Sub 雪莉_VBE_異質者()
Dim i, j, rrr As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "VBE異質者" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(43, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
    Select Case atkingck(43, 1)
         Case 1
             If atkingpagetot(1, 4) >= 3 And atkingck(43, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(43, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 4) < 3 And atkingck(43, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(43, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
         Case 2
           For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_異質者_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 10110
                atkingno(i, 7) = 43
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
           Next
           戰鬥系統類.人物技能欄燈開關 False, 2
           戰鬥系統類.自動捲軸捲動
     Case 3
'          atkingck(43, 2) = 0
          If Val(擲骰表單溝通暫時變數(3)) - Val(擲骰表單溝通暫時變數(2)) >= liveus(角色人物對戰人數(1, 2)) And 異常狀態檢查數(14, 2) = 0 Then
           For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, j, 2) = 0 Then
                    戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 6, 3
                    異常狀態檢查數(7, 1) = 1
                    異常狀態檢查數(7, 2) = 1
                    戰鬥系統類.人物異常狀態表設定_初設 1, j + 1, 14, app_path & "gif\異常狀態\不死.gif", 0, 3
                    異常狀態檢查數(14, 1) = 1
                    異常狀態檢查數(14, 2) = 1
                    戰鬥系統類.人物異常狀態表設定_初設 1, j + 2, 15, app_path & "gif\異常狀態\自壞.gif", 0, 3
                    異常狀態檢查數(15, 1) = 1
                    異常狀態檢查數(15, 2) = 1
                    atkingck(43, 2) = 0
                    Exit For
                 End If
           Next
         End If
    Case 4
        Do
            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, j, 2) >= 1 And 人物異常狀態資料庫(1, j, 3) = 14 Then
                     Exit Do
                 End If
            Next
           For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
              If 人物異常狀態資料庫(1, j, 2) = 0 Then
                    戰鬥系統類.人物異常狀態表設定_初設 1, j, 14, app_path & "gif\異常狀態\不死.gif", 0, 1
                    異常狀態檢查數(14, 1) = 1
                    異常狀態檢查數(14, 2) = 1
                 Exit Do
             End If
           Next
        Loop
        atkingck(43, 2) = 0
    End Select
End If
End Sub
Sub 雪莉_VBE_巨大黑犬()
Dim i As Integer, j As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "VBE巨大黑犬" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(44, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
   Select Case atkingck(44, 1)
        Case 1
'          If movecp < 3 Then
            If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 3) >= 2 And atkingck(44, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(44, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 3) < 2) And atkingck(44, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(44, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
'          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
'             atkingck(44, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\雪莉\雪莉_巨大黑犬_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9810
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 44
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
     Case 3
         Do
            atkingck(44, 2) = 0
            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                If 人物異常狀態資料庫(2, j, 3) = 5 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                 FormMainMode.personcomspe(j).person_num = 8
                 FormMainMode.personcomspe(j).person_turn = 6
                 人物異常狀態資料庫(2, j, 1) = 8
                 人物異常狀態資料庫(2, j, 2) = 6
                 Exit Do
                End If
            Next
           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 5, app_path & "gif\異常狀態\defdown.gif", 8, 6
                 異常狀態檢查數(5, 1) = 1
                 異常狀態檢查數(5, 2) = 1
                 Exit Do
             End If
           Next
        Loop
        If movecp = 1 Then
            戰鬥系統類.傷害執行_技能直傷_電腦 livecom(角色人物對戰人數(2, 2)), 1
        End If
   End Select
End If

End Sub
Sub 雪莉_VBE_飛刃雨()
Dim ttt As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "VBE飛刃雨" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(45, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
    Select Case atkingck(45, 1)
         Case 1
               If atkingck(45, 2) = 0 Then
                  For i = 1 To 106
                       If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          戰鬥系統類.人物技能欄燈開關 True, 4
                          atkingck(45, 2) = 1
                          atkingtrn(1) = Val(atkingtrn(1)) + 1
                          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(FormMainMode.pageusqlead) * movecp * 2
                          atkingck(45, 1) = 2
                          atking_sheri_4_tot = Val(FormMainMode.pageusqlead)
                          Exit For
                       End If
                  Next
               End If
               FormMainMode.trgoi1.Enabled = True
         Case 2
                  If atkingpagetot(1, 3) = 0 Then
                     戰鬥系統類.人物技能欄燈開關 False, 4
                     atkingck(45, 2) = 0
                     atkingtrn(1) = Val(atkingtrn(1)) - 1
                     atkingck(45, 1) = 1
                     If Val(FormMainMode.pageusqlead) = atking_sheri_4_tot Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(FormMainMode.pageusqlead) * movecp * 2
                     Else
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(FormMainMode.pageusqlead) * movecp * 2 - movecp * 2
                     End If
                     atking_sheri_4_tot = 0
                  ElseIf atkingpagetot(1, 3) > 1 Then
                     For i = 1 To 106
                       If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          ttt = ttt + 1
                       End If
                     Next
                     If ttt = 0 Then
                       戰鬥系統類.人物技能欄燈開關 False, 4
                       atkingck(45, 2) = 0
                       atkingtrn(1) = Val(atkingtrn(1)) - 1
                       atkingck(45, 1) = 1
                       If Val(FormMainMode.pageusqlead) = atking_sheri_4_tot Then
                          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(FormMainMode.pageusqlead) * movecp * 2
                       Else
                          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(FormMainMode.pageusqlead) * movecp * 2 - movecp * 2
                       End If
                       atking_sheri_4_tot = 0
                     End If
                  End If
                  If atkingck(45, 2) = 1 Then
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (Val(FormMainMode.pageusqlead) - Val(atking_sheri_4_tot)) * movecp * 2
                     atking_sheri_4_tot = Val(FormMainMode.pageusqlead)
                  End If
                  FormMainMode.trgoi1.Enabled = True
         Case 3
           For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_飛刃雨_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 9690
                atkingno(i, 7) = 0
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
           Next
           戰鬥系統類.人物技能欄燈開關 False, 4
           atkingck(45, 2) = 0
           戰鬥系統類.自動捲軸捲動
    End Select
End If
End Sub

Sub 雪莉_飛刃雨()
Dim ttt As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "飛刃雨" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(3, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
    Select Case atkingck(3, 1)
         Case 1
               If atkingck(3, 2) = 0 And movecp = 3 Then
                  For i = 1 To 106
                       If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          戰鬥系統類.人物技能欄燈開關 True, 4
                          atkingck(3, 2) = 1
                          atkingtrn(1) = Val(atkingtrn(1)) + 1
                          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + pageqlead(1) * 2
                          atkingck(3, 1) = 2
                          atking_sheri_4_tot = pageqlead(1)
                          Exit For
                       End If
                  Next
               End If
               FormMainMode.trgoi1.Enabled = True
         Case 2
                  If atkingpagetot(1, 3) = 0 Then
                     戰鬥系統類.人物技能欄燈開關 False, 4
                     atkingck(3, 2) = 0
                     atkingtrn(1) = Val(atkingtrn(1)) - 1
                     atkingck(3, 1) = 1
                     If pageqlead(1) = atking_sheri_4_tot Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 2
                     Else
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 2 - 2
                     End If
                     atking_sheri_4_tot = 0
                  ElseIf atkingpagetot(1, 3) > 1 Then
                     For i = 1 To 106
                       If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          ttt = ttt + 1
                       End If
                     Next
                     If ttt = 0 Then
                       戰鬥系統類.人物技能欄燈開關 False, 4
                       atkingck(3, 2) = 0
                       atkingtrn(1) = Val(atkingtrn(1)) - 1
                       atkingck(3, 1) = 1
                       If pageqlead(1) = atking_sheri_4_tot Then
                          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 2
                       Else
                          攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 2 - 2
                       End If
                       atking_sheri_4_tot = 0
                     End If
                  End If
                  If atkingck(3, 2) = 1 Then
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (pageqlead(1) - Val(atking_sheri_4_tot)) * 2
                     atking_sheri_4_tot = pageqlead(1)
                  End If
                  FormMainMode.trgoi1.Enabled = True
         Case 3
           For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_飛刃雨_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 9690
                atkingno(i, 7) = 0
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
           Next
           戰鬥系統類.人物技能欄燈開關 False, 4
           atkingck(3, 2) = 0
           戰鬥系統類.自動捲軸捲動
    End Select
End If
End Sub
Sub 雪莉_異質者()
Dim i, j, rrr As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "異質者" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(10, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "雪莉" Then
    Select Case atkingck(10, 1)
         Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
'                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) >= 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr = rrr + 1
                End If
             Next
             If rrr >= 1 And atkingck(10, 2) = 0 Then
                atkingck(10, 2) = 1
                戰鬥系統類.人物技能欄燈開關 True, 2
                atkingtrn(1) = Val(atkingtrn(1)) + 1
             End If
             If rrr < 1 And atkingck(10, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 2
                atkingck(10, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
         Case 2
           For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_異質者_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 10110
                atkingno(i, 7) = 10
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
           Next
           戰鬥系統類.人物技能欄燈開關 False, 2
           戰鬥系統類.自動捲軸捲動
     Case 3
          atkingck(10, 2) = 0
          If Val(擲骰表單溝通暫時變數(3)) - Val(擲骰表單溝通暫時變數(2)) >= liveus(角色人物對戰人數(1, 2)) And 異常狀態檢查數(14, 2) = 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 6
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 1) = 6
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 6, 3
                          異常狀態檢查數(7, 1) = 1
                          異常狀態檢查數(7, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '==================================
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 14 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 0
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 1) = 0
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 14, app_path & "gif\異常狀態\不死.gif", 0, 3
                          異常狀態檢查數(14, 1) = 1
                          異常狀態檢查數(14, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '===============================
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 15 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 0
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 1) = 0
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 15, app_path & "gif\異常狀態\自壞.gif", 0, 3
                          異常狀態檢查數(15, 1) = 1
                          異常狀態檢查數(15, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
         End If
    End Select
End If
End Sub
Sub 古魯瓦爾多_必殺架勢()
Dim i As Integer, j As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "必殺架勢" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(12, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "古魯瓦爾多" Then
   Select Case atkingck(12, 1)
        Case 1
            If atkingpagetot(1, 4) >= 2 And atkingpagetot(1, 3) = 0 And atkingck(12, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(12, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 4) < 2 Or atkingpagetot(1, 3) <> 0) And atkingck(12, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(12, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\古魯瓦爾多-必殺架勢1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 240
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9345
                   atkingno(i, 7) = 12
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
         Do
            atkingck(12, 2) = 0
            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                     FormMainMode.personusspe(j).person_num = 5
                     FormMainMode.personusspe(j).person_turn = 1
                     人物異常狀態資料庫(1, j, 1) = 5
                     人物異常狀態資料庫(1, j, 2) = 1
                     Exit Do
                 End If
            Next
           For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
              If 人物異常狀態資料庫(1, j, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 5, 1
                 異常狀態檢查數(7, 1) = 1
                 異常狀態檢查數(7, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub 古魯瓦爾多_血之恩賜()
Dim bloodtot As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "血之恩賜" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(60, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "古魯瓦爾多" Then
   Select Case atkingck(60, 1)
        Case 1
             If atkingpagetot(1, 2) >= 3 And atkingpagetot(1, 4) >= 2 And atkingck(60, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(60, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(60, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
            ElseIf (atkingpagetot(1, 2) < 3 Or atkingpagetot(1, 4) < 2) And atkingck(60, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(60, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(60, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\Grunwaldatking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6915
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 60
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(FormMainMode.顯示列1.goi2) <= 0 Then
                atkingck(60, 2) = 0
            End If
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) <= 0 Then
                bloodtot = Abs(Val(擲骰表單溝通暫時變數(2)))
                If Formsetting.checktest.Value = 1 Then Debug.Print "古魯瓦爾多-血之恩賜回復量:" & bloodtot
                戰鬥系統類.回復執行_使用者 bloodtot, 1
            End If
            '=============
            atkingck(60, 2) = 0
   End Select
End If
End Sub


Sub 古魯瓦爾多_猛擊()
Dim rrr As Integer '暫時變數
If FormMainMode.personatk(1).Caption = "猛擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(6, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "古魯瓦爾多" Then
   Select Case atkingck(6, 1)
      Case 1
          If movecp = 1 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr = rrr + 1
                End If
             Next
          End If
          If rrr >= 2 And atkingck(6, 2) = 0 Then
             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
             atkingck(6, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 1
             atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
          If rrr < 2 And atkingck(6, 2) = 1 Then
             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(6, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
           End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(6, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\古魯瓦爾多_猛擊_1.jpeg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 10305
                   atkingno(i, 6) = 8925
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 古魯瓦爾多_精神力吸收()
Dim rrr(1 To 3) As Integer '牌判斷暫時變數
If FormMainMode.personatk(4).Caption = "精神力吸收" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(61, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "古魯瓦爾多" Then
   Select Case atkingck(61, 1)
        Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                    If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(2) = rrr(2) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(3) = rrr(3) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingck(61, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingck(61, 2) = 0 Then
                戰鬥系統類.人物技能欄燈開關 True, 4
                atkingck(61, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingck(61, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingck(61, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 4
                atkingck(61, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
              End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
              For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\Grunwaldatking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1080
                   atkingno(i, 5) = 8025
                   atkingno(i, 6) = 9525
                   atkingno(i, 7) = 61
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
            cardpn = 0
'            Erase cardp
            Erase atking_古魯瓦爾多_精神力吸收紀錄數
            '=====================
            For i = 1 To 106
                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                    If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                         atking_古魯瓦爾多_精神力吸收紀錄數(i) = 1
                         atking_古魯瓦爾多_精神力吸收紀錄數(0) = atking_古魯瓦爾多_精神力吸收紀錄數(0) + 1
                     End If
                End If
            Next
            If atking_古魯瓦爾多_精神力吸收紀錄數(0) > 0 Then
                atking_古魯瓦爾多_精神力吸收紀錄數(0) = 0
                For i = 1 To 106
                    If atking_古魯瓦爾多_精神力吸收紀錄數(i) = 1 Then
                        atking_古魯瓦爾多_精神力吸收紀錄數(0) = Val(atking_古魯瓦爾多_精神力吸收紀錄數(0)) + 1
                        目前數(16) = i
                        atking_古魯瓦爾多_精神力吸收紀錄數(i) = 0
                        FormMainMode.tr電腦牌_翻牌.Enabled = True
                        Exit Sub
                    End If
                Next
            Else
               目前數(22) = 16
               FormMainMode.等待時間.Enabled = True
            End If
        Case 4
            FormMainMode.tr電腦牌_偷牌.Enabled = True
            目前數(17) = 5
        Case 5
            If atking_古魯瓦爾多_精神力吸收紀錄數(0) > 0 Then
                For i = 1 To 106
                    If atking_古魯瓦爾多_精神力吸收紀錄數(i) = 1 And atking_古魯瓦爾多_精神力吸收紀錄數(0) < 3 Then
                        atking_古魯瓦爾多_精神力吸收紀錄數(0) = Val(atking_古魯瓦爾多_精神力吸收紀錄數(0)) + 1
                        目前數(16) = i
                        atking_古魯瓦爾多_精神力吸收紀錄數(i) = 0
                        FormMainMode.tr電腦牌_翻牌.Enabled = True
                        Exit Sub
                    End If
                Next
                If i = 107 Then
                    atkingck(61, 2) = 0
                    戰鬥系統類.執行動作_技能手動結束
                 End If
            Else
               目前數(22) = 16
               FormMainMode.等待時間.Enabled = True
            End If
        Case 6
            If atking_古魯瓦爾多_精神力吸收紀錄數(0) = 0 Then
                atking_古魯瓦爾多_精神力吸收紀錄數(0) = 99
               目前數(22) = 16
               FormMainMode.等待時間.Enabled = True
            ElseIf atking_古魯瓦爾多_精神力吸收紀錄數(0) > 0 Then
               atkingck(61, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            End If
   End Select
End If

End Sub
Sub 艾茵_一顆心()
Dim cardnum(1 To 2) As Integer '暫時變數
If FormMainMode.personatk(1).Caption = "一顆心" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(37, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾茵" Then
   Select Case atkingck(37, 1)
        Case 1
           If movecp = 2 Then
                 If atkingpagetot(1, 4) >= 3 And atkingck(37, 2) = 0 Then
'                 If pageqlead(1) >= 1 And atkingck(37, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingck(37, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf atkingpagetot(1, 4) < 3 And atkingck(37, 2) = 1 Then
'                ElseIf pageqlead(1) < 1 And atkingck(37, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(37, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
           End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
              For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_一顆心.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8925
                   atkingno(i, 6) = 9105
                   atkingno(i, 7) = 37
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            '=====================
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                  If Val(pagecardnum(i, 2)) > cardnum(1) Then
                      cardnum(1) = pagecardnum(i, 2)
                      cardnum(2) = i
                  End If
                  If Val(pagecardnum(i, 4)) > cardnum(1) Then
                      cardnum(1) = pagecardnum(i, 4)
                      cardnum(2) = i
                  End If
               End If
            Next
            目前數(16) = cardnum(2)
            FormMainMode.tr電腦牌_翻牌.Enabled = True
        Case 4
'            電腦牌_偷牌 目前數(16)
            FormMainMode.tr電腦牌_偷牌.Enabled = True
            目前數(17) = 2
            atkingck(37, 2) = 0
   End Select
End If
End Sub
Sub 艾茵_兩個身體()
Dim bloodtot As Single  '暫時變數
Dim num As Integer
If FormMainMode.personatk(2).Caption = "兩個身體" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(32, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾茵" Then
   Select Case atkingck(32, 1)
        Case 1
             If atkingpagetot(1, 3) >= 1 And atkingck(32, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(32, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 3) < 1 And atkingck(32, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(32, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_兩個身體_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6330
                   atkingno(i, 6) = 9285
                   atkingno(i, 7) = 32
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(32, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                bloodtot = Val(擲骰表單溝通暫時變數(2)) \ Val(2)
                Do
                    Randomize
                    num = Int(Rnd() * 3) + 1
                    If livecom(角色待機人物紀錄數(2, num)) > 0 Then
                        戰鬥系統類.傷害執行_技能直傷_電腦 bloodtot, num
                        Exit Do
                    End If
                Loop
            End If
   End Select
End If
End Sub

Sub 艾茵_九個靈魂()
Dim bloodtot As Single  '暫時變數
Dim pic As Integer 'RND暫時變數
If FormMainMode.personatk(3).Caption = "九個靈魂" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(26, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾茵" Then
   Select Case atkingck(26, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(1, 2) >= 5 And atkingpagetot(1, 4) >= 1 And atkingck(26, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(26, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(26, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 9
            ElseIf (atkingpagetot(1, 2) < 5 Or atkingpagetot(1, 4) < 1) And atkingck(26, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(26, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(26, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 9
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_九個靈魂\艾茵_九個靈魂main.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6330
                   atkingno(i, 6) = 9510
                   atkingno(i, 7) = 26
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 11) = 0
                   '=================
                   Randomize
                   pic = Int(Rnd() * 8) + 1
                   atkingno(i, 10) = app_path & "gif\艾茵\艾茵_九個靈魂\艾茵_九個靈魂" & pic & ".jpg"
                   Exit For
                 End If
             Next
        Case 3
            bloodtot = Int(atkingpagetot(1, 4) / 2 + 0.5)
            '=============
            If Val(liveus(角色人物對戰人數(1, 2))) < Val(liveusmax(角色人物對戰人數(1, 2))) Then
                戰鬥系統類.回復執行_使用者 bloodtot, 1
            End If
            atkingck(26, 2) = 0
   End Select
End If
End Sub
Sub 艾茵_十三隻眼()
Dim rrr(1 To 2) As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "十三隻眼" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(16, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾茵" Then
   Select Case atkingck(16, 1)
        Case 1
           If movecp < 3 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr(1) = rrr(1) + 1
                End If
                If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr(2) = rrr(2) + 1
                End If
             Next
           End If
          If rrr(1) >= 1 And rrr(2) >= 1 And atkingck(16, 2) = 0 Then
'          If rrr(1) >= 1 And atkingck(16, 2) = 0 Then
             atkingck(16, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 4
             atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
          If (rrr(1) < 1 Or rrr(2) < 1) And atkingck(16, 2) = 1 Then
'          If rrr(1) < 1 And atkingck(16, 2) = 1 Then
             戰鬥系統類.人物技能欄燈開關 False, 4
             atkingck(16, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
           End If
        Case 2
            If atking_艾茵_十三隻眼_tot(2) = 0 Then
                atking_艾茵_十三隻眼_tot(1) = 攻擊防禦骰子總數(1)
                atking_艾茵_十三隻眼_tot(2) = 1
                攻擊防禦骰子總數(1) = 13
                atkingck(16, 1) = 1
            ElseIf atking_艾茵_十三隻眼_tot(2) = 1 Then
                atking_艾茵_十三隻眼_tot(1) = atking_艾茵_十三隻眼_tot(1) + (攻擊防禦骰子總數(1) - 13)
                攻擊防禦骰子總數(1) = 13
                atkingck(16, 1) = 1
            End If
        Case 3
           atking_艾茵_十三隻眼_tot(1) = atking_艾茵_十三隻眼_tot(1) + (攻擊防禦骰子總數(1) - 13)
           攻擊防禦骰子總數(1) = atking_艾茵_十三隻眼_tot(1)
           atking_艾茵_十三隻眼_tot(1) = 0
           atking_艾茵_十三隻眼_tot(2) = 0
           atkingck(16, 1) = 1
        Case 4
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             atkingck(16, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_十三隻眼.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7980
                   atkingno(i, 6) = 9015
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             Erase atking_艾茵_十三隻眼_tot
             '===============
            戰鬥系統類.直接寫入顯示列數值 1, 13
'            攻擊防禦骰子總數(1) = FormMainMode.顯示列1.goi1
            戰鬥系統類.直接寫入顯示列數值 2, 0
'            攻擊防禦骰子總數(2) = FormMainMode.顯示列1.goi2
        Case 5
            攻擊防禦骰子總數(2) = 0
   End Select
End If
End Sub
Sub 蕾_輪旋曲_琉璃色的微風()
If FormMainMode.personatk(1).Caption = "輪旋曲-琉璃色的微風" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(13, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(13, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(1, 1) >= 4 And atkingck(13, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(13, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
            ElseIf atkingpagetot(1, 1) < 4 And atkingck(13, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(13, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
            End If
          End If
        Case 2
'            戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 3
        Case 3
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(13, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-輪旋曲-琉璃色的微風.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7230
                   atkingno(i, 6) = 9000
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
'             戰鬥系統類.直接寫入顯示列數值 2, 攻擊防禦骰子總數(2)
             戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 3
        Case 4
'            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
   End Select
End If
End Sub
Sub 蕾_EX_輪旋曲_琉璃色的微風()
If FormMainMode.personatk(1).Caption = "Ex輪旋曲-琉璃色的微風" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(19, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(19, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(1, 1) >= 5 And atkingck(19, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(19, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 8
            ElseIf atkingpagetot(1, 1) < 5 And atkingck(19, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(19, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 8
            End If
          End If
        Case 2
'            戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 6
'            formmainmode.trgoi1.Enabled = True
        Case 3
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(19, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-輪旋曲-琉璃色的微風.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7320
                   atkingno(i, 6) = 9000
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
'             戰鬥系統類.直接寫入顯示列數值 2, 攻擊防禦骰子總數(2)
             戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 6
        Case 4
'            戰鬥系統類.直接寫入顯示列數值 2, 攻擊防禦骰子總數(2)
'            戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 6
'            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
   End Select
End If
End Sub
Sub 蕾_協奏曲_加百烈的守護()
Dim i As Integer, j As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "協奏曲-加百烈的守護" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(11, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(11, 1)
        Case 1
            If atkingpagetot(1, 4) >= 2 And atkingpagetot(1, 3) >= 1 And atkingck(11, 2) = 0 Then
'            If atkingpagetot(1, 3) >= 1 And atkingck(11, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(11, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 4) < 2 Or atkingpagetot(1, 3) < 1) And atkingck(11, 2) = 1 Then
'            ElseIf atkingpagetot(1, 3) < 1 And atkingck(11, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(11, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-協奏曲-加百烈的守護.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6180
                   atkingno(i, 6) = 9000
                   atkingno(i, 7) = 11
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
          Do
            atkingck(11, 2) = 0
            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, j, 1) >= 10 And 人物異常狀態資料庫(1, j, 3) = 8 Then
                     FormMainMode.personusspe(j).person_turn = 3
                     人物異常狀態資料庫(1, j, 2) = 3
                     Exit Do
                 End If
                 If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                     FormMainMode.personusspe(j).person_num = 人物異常狀態資料庫(1, j, 1) + 1
                     FormMainMode.personusspe(j).person_turn = 3
                     人物異常狀態資料庫(1, j, 1) = 人物異常狀態資料庫(1, j, 1) + 1
                     人物異常狀態資料庫(1, j, 2) = 3
                     '========DEF+1立即生效
'                         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 1
                         戰鬥系統類.直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) + 1
                    '===============
                     Exit Do
                 End If
            Next
           For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
              If 人物異常狀態資料庫(1, j, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 1, j, 8, app_path & "gif\異常狀態\defup.gif", 3, 3
                 異常狀態檢查數(8, 1) = 1
                 異常狀態檢查數(8, 2) = 1
                  '========DEF+3立即生效
'                         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
                         戰鬥系統類.直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) + 3
                  '===============
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub 蕾_EX_協奏曲_加百烈的守護()
If FormMainMode.personatk(2).Caption = "Ex協奏曲-加百烈的守護" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(38, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(38, 1)
        Case 1
            If atkingpagetot(1, 4) >= 3 And atkingpagetot(1, 3) >= 1 And atkingck(38, 2) = 0 Then
'            If atkingpagetot(1, 3) >= 1 And atkingck(38, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(38, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
            ElseIf (atkingpagetot(1, 4) < 3 Or atkingpagetot(1, 3) < 1) And atkingck(38, 2) = 1 Then
'            ElseIf atkingpagetot(1, 3) < 1 And atkingck(38, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(38, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-協奏曲-加百烈的守護.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6540
                   atkingno(i, 6) = 9420
                   atkingno(i, 7) = 38
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
          '==========================
          atking_蕾_守護模式狀態啟動值 = True
    Case 3
          atking_蕾_守護模式狀態啟動值 = False
          atkingck(38, 2) = 0
    Case 4
          If Formsetting.checktest.Value = 1 Then Debug.Print "蕾-Ex協奏曲-加百烈的守護擲骰表單溝通暫時變數(2)前:" & 擲骰表單溝通暫時變數(2)
          擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) - 5
          擲骰後骰傷害數 = 擲骰後骰傷害數 - 5
          If Formsetting.checktest.Value = 1 Then Debug.Print "蕾-Ex協奏曲-加百烈的守護擲骰表單溝通暫時變數(2)後:" & 擲骰表單溝通暫時變數(2)
   End Select
End If
End Sub

Sub 蕾_安魂曲_死神的鎮魂歌()
Dim rrr As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "安魂曲-死神的鎮魂歌" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(14, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(14, 1)
        Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr = rrr + 1
                End If
             Next
          If rrr >= 1 And atkingck(14, 2) = 0 Then
             atkingck(14, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 3
             atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
          If rrr < 1 And atkingck(14, 2) = 1 Then
             戰鬥系統類.人物技能欄燈開關 False, 3
             atkingck(14, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
           End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-安魂曲-死神的鎮魂歌.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = -360
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8835
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(14, 2) = 0
             If liveus(角色人物對戰人數(1, 2)) <= 0 Then
                 For i = 2 To 3
                     If FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption > 0 Then
                        Do
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                                  If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 5
                                      FormMainMode.personusspe(j).person_turn = 3
                                      人物異常狀態資料庫(1, j, 1) = 5
                                      人物異常狀態資料庫(1, j, 2) = 3
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 5, 3
                                  異常狀態檢查數(7, 1) = 1
                                  異常狀態檢查數(7, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        Do
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                                  If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 5
                                      FormMainMode.personusspe(j).person_turn = 3
                                      人物異常狀態資料庫(1, j, 1) = 5
                                      人物異常狀態資料庫(1, j, 2) = 3
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 8, app_path & "gif\異常狀態\defup.gif", 5, 3
                                  異常狀態檢查數(8, 1) = 1
                                  異常狀態檢查數(8, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub 蕾_EX_安魂曲_死神的鎮魂歌()
Dim rrr As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "Ex安魂曲-死神的鎮魂歌" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(62, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(62, 1)
        Case 1
             For i = 1 To 106
               If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                  rrr = rrr + 1
               End If
            Next
          If rrr >= 1 And atkingck(62, 2) = 0 Then
             atkingck(62, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 3
             atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
          If rrr < 1 And atkingck(62, 2) = 1 Then
             戰鬥系統類.人物技能欄燈開關 False, 3
             atkingck(62, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
           End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-安魂曲-死神的鎮魂歌.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9675
                   atkingno(i, 6) = 10155
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(62, 2) = 0
             If liveus(角色人物對戰人數(1, 2)) <= 0 Then
                 For i = 2 To 3
                     If FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption > 0 Then
                        Do
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                                  If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 9
                                      FormMainMode.personusspe(j).person_turn = 2
                                      人物異常狀態資料庫(1, j, 1) = 9
                                      人物異常狀態資料庫(1, j, 2) = 2
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 9, 2
                                  異常狀態檢查數(7, 1) = 1
                                  異常狀態檢查數(7, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        Do
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                                  If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 9
                                      FormMainMode.personusspe(j).person_turn = 2
                                      人物異常狀態資料庫(1, j, 1) = 9
                                      人物異常狀態資料庫(1, j, 2) = 2
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 8, app_path & "gif\異常狀態\defup.gif", 9, 2
                                  異常狀態檢查數(8, 1) = 1
                                  異常狀態檢查數(8, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub 蕾_終曲_無盡輪迴的終結()
Dim num(1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.personatk(4).Caption = "終曲-無盡輪迴的終結" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(15, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(15, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(1, 4) >= 4 And atkingck(15, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(15, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 16
            ElseIf atkingpagetot(1, 4) < 4 And atkingck(15, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(15, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 16
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-終曲-無盡輪迴的終結.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8655
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 15
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(15, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                 num(1) = 1
                 num(2) = livecom(角色人物對戰人數(2, 2))
                 For i = 2 To 3
                    If livecom(角色待機人物紀錄數(2, i)) > 0 And livecom(角色待機人物紀錄數(2, i)) < num(2) Then
                        num(1) = i
                        num(2) = livecom(角色待機人物紀錄數(2, i))
                    End If
                Next
                戰鬥系統類.傷害執行_技能直傷_電腦 Val(擲骰表單溝通暫時變數(2)), num(1)
            End If
            擲骰表單溝通暫時變數(2) = 0
            擲骰後骰傷害數 = 0
   End Select
End If
End Sub
Sub 蕾_EX_終曲_無盡輪迴的終結()
Dim num(1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.personatk(4).Caption = "Ex終曲-無盡輪迴的終結" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(161, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(161, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(1, 4) >= 6 And atkingck(161, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(161, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 18
            ElseIf atkingpagetot(1, 4) < 6 And atkingck(161, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(161, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 18
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             atking_蕾_終曲_無盡輪迴的終結紀錄數 = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-終曲-無盡輪迴的終結.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8655
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 161
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=============
             atking_蕾_終曲_無盡輪迴的終結紀錄數 = atkingpagetot(2, 2)
        Case 3
             atkingck(161, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                 num(1) = 1
                 num(2) = livecom(角色人物對戰人數(2, 2))
                 For i = 2 To 3
                    If livecom(角色待機人物紀錄數(2, i)) > 0 And livecom(角色待機人物紀錄數(2, i)) < num(2) Then
                        num(1) = i
                        num(2) = livecom(角色待機人物紀錄數(2, i))
                    End If
                Next
                戰鬥系統類.傷害執行_技能直傷_電腦 Val(擲骰表單溝通暫時變數(2)), num(1)
            End If
            '=================
            戰鬥系統類.傷害執行_技能直傷_電腦 Val(atking_蕾_終曲_無盡輪迴的終結紀錄數), 1
            擲骰表單溝通暫時變數(2) = 0
            擲骰後骰傷害數 = 0
   End Select
End If
End Sub

Sub 蕾_終曲_無盡輪迴的終結_舊()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "終曲-無盡輪迴的終結" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(15, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾" Then
   Select Case atkingck(15, 1)
        Case 1
           If movecp < 3 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr(1) = rrr(1) + 1
                End If
                If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr(2) = rrr(2) + 1
                End If
                If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr(3) = rrr(3) + 1
                End If
             Next
           End If
          If rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1 And atkingck(15, 2) = 0 Then
          'If rrr(1) >= 1 And atkingck(15, 2) = 0 Then
             atkingck(15, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 4
             atkingtrn(1) = Val(atkingtrn(1)) + 1
             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 10
          End If
          If (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingck(15, 2) = 1 Then
          'If rrr(1) < 1 And atkingck(15, 2) = 1 Then
             戰鬥系統類.人物技能欄燈開關 False, 4
             atkingck(15, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 10
           End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             atkingck(15, 1) = 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-終曲-無盡輪迴的終結.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8655
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If atkingck(15, 2) = 1 Then
                 atkingck(15, 2) = 0
                 Form6.技能_蕾_終曲_無盡輪迴的終結_舊_分支_階段三
             End If
   End Select
End If
End Sub
Sub 伊芙琳_怠惰的墓表()
Dim cardp(1 To 106) As Boolean '紀錄暫時變數
Dim cardpn As Integer '紀錄牌總數暫時變數
If FormMainMode.personatk(1).Caption = "怠惰的墓表" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(56, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "伊芙琳" Then
   Select Case atkingck(56, 1)
        Case 1
           If movecp < 3 Then
                 If atkingpagetot(1, 4) >= 2 And atkingck(56, 2) = 0 Then
'                 If pageqlead(1) >= 1 And atkingck(56, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingck(56, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf atkingpagetot(1, 4) < 2 And atkingck(56, 2) = 1 Then
'                ElseIf pageqlead(1) < 1 And atkingck(56, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(56, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
           End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
              For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6345
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 56
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
            Do
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, i, 2) >= 9 And 人物異常狀態資料庫(1, i, 3) = 24 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(1, i, 3) = 24 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 9 Then
                     FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) + 1
                     人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 1, i, 24, app_path & "gif\異常狀態\能力低下.gif", 0, 1
                     異常狀態檢查數(24, 1) = 1
                     異常狀態檢查數(24, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
            '=====================
            cardpn = 0
            Erase cardp
            Erase atking_伊芙琳_怠惰的墓表紀錄數
            '=====================
            Do
               Randomize
               i = Int(Rnd() * 106) + 1
               If cardp(i) = False Then
                    cardp(i) = True
                    cardpn = cardpn + 1
                    If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                      Select Case movecp
                         Case 1
                             If pagecardnum(i, 1) = a1a Or pagecardnum(i, 3) = a1a Then
                                  atking_伊芙琳_怠惰的墓表紀錄數(atking_伊芙琳_怠惰的墓表紀錄數(0) + 1) = i
                                  atking_伊芙琳_怠惰的墓表紀錄數(0) = atking_伊芙琳_怠惰的墓表紀錄數(0) + 1
                              End If
                         Case Is > 1
                             If pagecardnum(i, 1) = a5a Or pagecardnum(i, 3) = a5a Then
                                  atking_伊芙琳_怠惰的墓表紀錄數(atking_伊芙琳_怠惰的墓表紀錄數(0) + 1) = i
                                  atking_伊芙琳_怠惰的墓表紀錄數(0) = atking_伊芙琳_怠惰的墓表紀錄數(0) + 1
                              End If
                        End Select
                    End If
               End If
               If atking_伊芙琳_怠惰的墓表紀錄數(0) >= 2 Then
                   Exit Do
               End If
            Loop While cardpn < 106
            If atking_伊芙琳_怠惰的墓表紀錄數(0) > 0 Then
                目前數(16) = atking_伊芙琳_怠惰的墓表紀錄數(1)
                FormMainMode.tr電腦牌_翻牌.Enabled = True
            Else
'               atkingck(56, 2) = 0
               目前數(22) = 1
               FormMainMode.等待時間.Enabled = True
            End If
        Case 4
            FormMainMode.tr電腦牌_偷牌.Enabled = True
            目前數(17) = 3
            If atking_伊芙琳_怠惰的墓表紀錄數(0) = 3 Then
                atkingck(56, 1) = 6
            End If
        Case 5
             If atking_伊芙琳_怠惰的墓表紀錄數(0) < 2 Then
                目前數(22) = 1
                FormMainMode.等待時間.Enabled = True
            Else
                目前數(16) = atking_伊芙琳_怠惰的墓表紀錄數(2)
                atking_伊芙琳_怠惰的墓表紀錄數(0) = 3
                FormMainMode.tr電腦牌_翻牌.Enabled = True
            End If
        Case 6
            If atking_伊芙琳_怠惰的墓表紀錄數(0) = 0 Then
               atking_伊芙琳_怠惰的墓表紀錄數(0) = 3
               目前數(22) = 1
               FormMainMode.等待時間.Enabled = True
            ElseIf atking_伊芙琳_怠惰的墓表紀錄數(0) > 0 Then
               atkingck(56, 2) = 0
               執行動作_技能手動結束
            End If
   End Select
End If

End Sub
Sub 伊芙琳_慟哭之歌()
If FormMainMode.personatk(2).Caption = "慟哭之歌" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(57, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "伊芙琳" Then
   Select Case atkingck(57, 1)
        Case 1
            If movecp > 1 Then
                If atkingpagetot(1, 2) >= 3 And atkingpagetot(1, 4) >= 1 And atkingck(57, 2) = 0 Then
    '             If atkingpagetot(1, 2) >= 1 And atkingck(57, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingck(57, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf (atkingpagetot(1, 2) < 3 Or atkingpagetot(1, 4) < 1) And atkingck(57, 2) = 1 Then
    '            ElseIf atkingpagetot(1, 2) < 1 And atkingck(57, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(57, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6195
                   atkingno(i, 6) = 8730
                   atkingno(i, 7) = 57
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
             戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) \ 2
'             攻擊防禦骰子總數(2) = FormMainMode.顯示列1.goi2
        Case 3
            atkingck(57, 2) = 0
            Do
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, i, 2) >= 9 And 人物異常狀態資料庫(1, i, 3) = 24 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(1, i, 3) = 24 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 9 Then
                     FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) + 1
                     人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 1, i, 24, app_path & "gif\異常狀態\能力低下.gif", 0, 1
                     異常狀態檢查數(24, 1) = 1
                     異常狀態檢查數(24, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
   End Select
End If
End Sub
Sub 伊芙琳_紅蓮車輪()
Dim bloodtot As Integer '暫時變數
Dim num(1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.personatk(3).Caption = "紅蓮車輪" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(58, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "伊芙琳" Then
   Select Case atkingck(58, 1)
        Case 1
            If movecp < 3 Then
                 If atkingpagetot(1, 1) >= 2 And atkingpagetot(1, 5) >= 2 And atkingpagetot(1, 4) >= 1 And atkingck(58, 2) = 0 Then
    '             If atkingpagetot(1, 1) >= 1 And atkingck(34, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingck(58, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 13
                ElseIf (atkingpagetot(1, 1) < 2 Or atkingpagetot(1, 5) < 2 Or atkingpagetot(1, 4) < 1) And atkingck(58, 2) = 1 Then
    '            ElseIf atkingpagetot(1, 1) < 1 And atkingck(34, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(58, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 13
                End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking3-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 8250
                   atkingno(i, 6) = 10275
                   atkingno(i, 7) = 58
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\伊芙琳\Evelynatking3-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            Do
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, i, 2) >= 9 And 人物異常狀態資料庫(1, i, 3) = 24 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(1, i, 3) = 24 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 9 Then
                     FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) + 1
                     人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 1, i, 24, app_path & "gif\異常狀態\能力低下.gif", 0, 1
                     異常狀態檢查數(24, 1) = 1
                     異常狀態檢查數(24, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
        Case 4
            bloodtot = Val(FormMainMode.顯示列1.goi1) \ 10
            num(2) = 999
            For i = 1 To 3
               If livecom(角色待機人物紀錄數(2, i)) < num(2) And livecom(角色待機人物紀錄數(2, i)) > 0 Then
                   num(1) = i
                   num(2) = livecom(角色待機人物紀錄數(2, i))
               End If
            Next
            戰鬥系統類.傷害執行_技能直傷_電腦 bloodtot, num(1)
            atkingck(58, 2) = 0
   End Select
End If
End Sub
Sub 伊芙琳_赤紅石榴()
Dim mkp As Integer '暫時變數
Dim cardp(1 To 106) As Boolean '紀錄暫時變數
Dim cardpn(1 To 2) As Integer '紀錄牌總數暫時變數(1.牌紀錄目前總數/2.牌選定目前總數)
If FormMainMode.personatk(4).Caption = "赤紅石榴" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(59, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "伊芙琳" Then
   If Formsetting.checktest.Value = 1 Then Debug.Print "經過赤紅石榴主名字判斷"
   Select Case atkingck(59, 1)
        Case 1
            If movecp = 3 Then
                 If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 _
                    And atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 5) >= 1 And atkingck(59, 2) = 0 Then
'                 If atkingpagetot(1, 1) >= 1 And atkingck(59, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingck(59, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1 _
                   Or atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 5) < 1) And atkingck(59, 2) = 1 Then
'                ElseIf atkingpagetot(1, 1) < 1 And atkingck(59, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(59, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             '==================================
             Erase atking_伊芙琳_赤紅石榴階段紀錄數
             Randomize
             mkp = Int(Rnd() * 16) + 1
             atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = mkp
             If Formsetting.checktest.Value = 1 Then Debug.Print "技能 - 伊芙琳 - 赤紅石榴效果值" & mkp
             '===================================
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9330
                   atkingno(i, 6) = 9165
                   atkingno(i, 7) = 59
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                    If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) <= 9 Or atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) >= 13 Then
                       atkingno(i, 11) = 0
                    Else
                       atkingno(i, 11) = 1
                    End If
                   Exit For
                 End If
             Next
        Case 3
            '======================
               執行動作_清除所有異常狀態_使用者
            '======================
            Select Case atking_伊芙琳_赤紅石榴階段紀錄數(0, 1)
                Case 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 1, 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 1, 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 1, 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 1, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 1, 2
                    戰鬥系統類.傷害執行_技能直傷_電腦 1, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員1點傷害。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, 2
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員3點傷害。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 3
                    戰鬥系統類.傷害執行_技能直傷_使用者 5, 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 5, 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 5, 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 5, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 5, 2
                    戰鬥系統類.傷害執行_技能直傷_電腦 5, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員5點傷害。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                '==============================================
                Case 4
                    回復執行_使用者 1, 1
                    回復執行_使用者 1, 2
                    回復執行_使用者 1, 3
                    回復執行_電腦 1, 1
                    回復執行_電腦 1, 2
                    回復執行_電腦 1, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員HP回復1點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 5
                    回復執行_使用者 3, 1
                    回復執行_使用者 3, 2
                    回復執行_使用者 3, 3
                    回復執行_電腦 3, 1
                    回復執行_電腦 3, 2
                    回復執行_電腦 3, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員HP回復3點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 6
                    回復執行_使用者 5, 1
                    回復執行_使用者 5, 2
                    回復執行_使用者 5, 3
                    回復執行_電腦 5, 1
                    回復執行_電腦 5, 2
                    回復執行_電腦 5, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員HP回復5點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                '===============================================
                Case 7
                    戰鬥系統類.傷害執行_技能直傷_使用者 Val(liveus(角色人物對戰人數(1, 2))) - 1, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 Val(livecom(角色人物對戰人數(2, 2))) - 1, 1
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己與對方的HP變為1點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                '============================================
                Case 8
                    If Val(liveus(角色人物對戰人數(1, 2))) > 5 Then
                        戰鬥系統類.傷害執行_技能直傷_使用者 Val(liveus(角色人物對戰人數(1, 2))) - 5, 1
                    Else
                        回復執行_使用者 5 - Val(liveus(角色人物對戰人數(1, 2))), 1
                    End If
                    If Val(livecom(角色人物對戰人數(2, 2))) > 5 Then
                        戰鬥系統類.傷害執行_技能直傷_電腦 Val(livecom(角色人物對戰人數(2, 2))) - 5, 1
                    Else
                        回復執行_電腦 5 - Val(livecom(角色人物對戰人數(2, 2))), 1
                    End If
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己與對方的HP變為5點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                '===============================================
                Case 9
                    回復執行_使用者 Val(liveusmax(角色人物對戰人數(1, 2))) - Val(liveus(角色人物對戰人數(1, 2))), 1
                    回復執行_電腦 Val(livecommax(角色人物對戰人數(2, 2))) - Val(livecom(角色人物對戰人數(2, 2))), 1
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己與對方的HP完全恢復。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                '===============================================
                Case 10
                    目前數(20) = 1
                    atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1
                '==========使用者棄牌階段
                    Do
                        If Val(pagecardnum(目前數(20), 5)) = 1 And Val(pagecardnum(目前數(20), 6)) = 1 Then
                            atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                            目前數(21) = 2
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                        目前數(20) = 目前數(20) + 1
                    Loop Until 目前數(20) > 106
                    If 目前數(20) > 106 Then
                        GoTo 效果10_使用者棄牌階段直接跳過
                    End If
                '============================================
                Case 11
                    目前數(20) = 1
                    '========使用者牌數判斷及選擇
                    If Val(FormMainMode.pageusglead) > 8 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pageusglead) - 8 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                                atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(21) = 2
                                FormMainMode.tr使用者_棄牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(20) = 目前數(20) + 1
                        Loop Until 目前數(20) > 106
                    ElseIf Val(FormMainMode.pageusglead) < 8 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 3
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 8
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        目前數(15) = 3
                        If Val(FormMainMode.pageul) < 8 - Val(FormMainMode.pageusglead) Then
                            戰鬥系統類.執行動作_洗牌
                        End If
                        FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                    Else
                        GoTo 效果11_移至電腦判斷
                    End If
                '============================================
                Case 12
                    目前數(20) = 1
                    '========使用者牌數判斷及選擇
                    If Val(FormMainMode.pageusglead) > 15 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pageusglead) - 15 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                                atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(21) = 2
                                FormMainMode.tr使用者_棄牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(20) = 目前數(20) + 1
                        Loop Until 目前數(20) > 106
                    ElseIf Val(FormMainMode.pageusglead) < 15 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 3
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 15
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        目前數(15) = 3
                        If Val(FormMainMode.pageul) < 15 - Val(FormMainMode.pageusglead) Then
                            戰鬥系統類.執行動作_洗牌
                        End If
                        FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                    Else
                        GoTo 效果12_移至電腦判斷
                    End If
                '===============================================
                Case 13
                    執行動作_距離變更 (1)
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  距離變為近距離。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 14
                    執行動作_距離變更 (2)
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  距離變為中距離。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 15
                    執行動作_距離變更 (3)
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  距離變為遠距離。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
                Case 16
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  什麼都沒有發生。"
                    戰鬥系統類.自動捲軸捲動
                    atkingck(59, 2) = 0
            End Select
        '=====================================================
       Case 4
             If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 10 Then
                    '==========使用者棄牌階段2
                    If 目前數(20) <= 106 Then
                        Do
                            If Val(pagecardnum(目前數(20), 5)) = 1 And Val(pagecardnum(目前數(20), 6)) = 1 Then
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(21) = 2
                                FormMainMode.tr使用者_棄牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(20) = 目前數(20) + 1
                        Loop Until 目前數(20) > 106
                    End If
效果10_使用者棄牌階段直接跳過:
                    If 目前數(20) > 106 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1 Then
                        目前數(16) = 1
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2
                        '=============電腦方棄牌階段1
                        Do
                            If Val(pagecardnum(目前數(16), 5)) = 2 And Val(pagecardnum(目前數(16), 6)) = 1 Then
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(17) = 4
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                        If 目前數(16) > 106 Then
                            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                                GoTo 效果結束實行_手牌變化類
                            Else
                                目前數(22) = 33
                                FormMainMode.等待時間.Enabled = True
                            End If
                        End If
                    End If
                    If 目前數(16) <= 106 And atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2 Then
                        '==============電腦方棄牌階段2
                        Do
                            If Val(pagecardnum(目前數(16), 5)) = 2 And Val(pagecardnum(目前數(16), 6)) = 1 Then
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(17) = 4
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                        If 目前數(16) > 106 Then
                            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                                GoTo 效果結束實行_手牌變化類
                            Else
                                目前數(22) = 33
                                FormMainMode.等待時間.Enabled = True
                            End If
                        End If
                    End If
            End If
        '=====================================================
        Case 5
            FormMainMode.tr電腦牌_棄牌.Enabled = True
        '=====================================================
        Case 6
            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
                    Do
                        If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                            atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                            atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                            目前數(21) = 2
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                        目前數(20) = 目前數(20) + 1
                    Loop Until 目前數(20) > 106
                    '=========電腦牌數判斷及選擇
效果11_移至電腦判斷:
                    目前數(16) = 1
                    If Val(FormMainMode.pagecomglead) > 8 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2
                        Erase cardp
                        Erase cardpn
                        For i = 1 To 106
                            atking_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 0
                        Next
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pagecomglead) - 8 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                                目前數(17) = 4
                                atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                    ElseIf Val(FormMainMode.pagecomglead) < 8 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 4
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 8
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        If Val(FormMainMode.pageul) < 8 - Val(FormMainMode.pagecomglead) Then
                            戰鬥系統類.執行動作_洗牌
                        End If
                        目前數(15) = 3
                        FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Else
                        If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                            GoTo 效果結束實行_手牌變化類
                        Else
                            目前數(22) = 33
                            FormMainMode.等待時間.Enabled = True
                        End If
                    End If
            End If
        '=====================================================
        Case 7
            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
                 Do
                     If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                         目前數(17) = 4
                         atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                         atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                         FormMainMode.tr電腦牌_翻牌.Enabled = True
                         Exit Sub
                     End If
                     目前數(16) = 目前數(16) + 1
                 Loop Until 目前數(16) > 106
                 If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                        GoTo 效果結束實行_手牌變化類
                 Else
                        目前數(22) = 33
                        FormMainMode.等待時間.Enabled = True
                 End If
             End If
        '=====================================================
        Case 8
            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
                Select Case atking_伊芙琳_赤紅石榴階段紀錄數(0, 2)
                    Case 3
                        If Val(FormMainMode.pageusglead) < atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 3
                           atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                        End If
                        If Val(FormMainMode.pageusglead) >= atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           GoTo 效果11_移至電腦判斷
                        End If
                    Case 4
                        If Val(FormMainMode.pagecomglead) < atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 3
                           atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                        End If
                        If Val(FormMainMode.pagecomglead) >= atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                                GoTo 效果結束實行_手牌變化類
                            Else
                                目前數(22) = 33
                                FormMainMode.等待時間.Enabled = True
                            End If
                        End If
                 End Select
            End If
        '=====================================================
        Case 9
            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
                    Do
                        If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                            atking_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                            atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                            目前數(21) = 2
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                        目前數(20) = 目前數(20) + 1
                    Loop Until 目前數(20) > 106
                    '=========電腦牌數判斷及選擇
效果12_移至電腦判斷:
                    目前數(16) = 1
                    If Val(FormMainMode.pagecomglead) > 15 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2
                        Erase cardp
                        Erase cardpn
                        For i = 1 To 106
                            atking_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 0
                        Next
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pagecomglead) - 15 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                                目前數(17) = 4
                                atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                                atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                    ElseIf Val(FormMainMode.pagecomglead) < 15 Then
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 4
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 15
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        If Val(FormMainMode.pageul) < 15 - Val(FormMainMode.pagecomglead) Then
                            戰鬥系統類.執行動作_洗牌
                        End If
                        目前數(15) = 3
                        FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Else
                        If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                            GoTo 效果結束實行_手牌變化類
                        Else
                            目前數(22) = 33
                            FormMainMode.等待時間.Enabled = True
                        End If
                    End If
            End If
        '=====================================================
        Case 10
            If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
                Do
                    If atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                        目前數(17) = 4
                        atking_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                        atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        FormMainMode.tr電腦牌_翻牌.Enabled = True
                        Exit Sub
                    End If
                    目前數(16) = 目前數(16) + 1
                Loop Until 目前數(16) > 106
                If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                    GoTo 效果結束實行_手牌變化類
                Else
                    目前數(22) = 33
                    FormMainMode.等待時間.Enabled = True
                End If
            End If
        '=====================================================
       Case 11
           If atking_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
                Select Case atking_伊芙琳_赤紅石榴階段紀錄數(0, 2)
                    Case 3
                        If Val(FormMainMode.pageusglead) < atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 3
                           atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                        End If
                        If Val(FormMainMode.pageusglead) >= atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           GoTo 效果12_移至電腦判斷
                        End If
                    Case 4
                        If Val(FormMainMode.pagecomglead) < atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 3
                           atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                        End If
                        If Val(FormMainMode.pagecomglead) >= atking_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           If atking_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                                GoTo 效果結束實行_手牌變化類
                            Else
                                目前數(22) = 33
                                FormMainMode.等待時間.Enabled = True
                            End If
                        End If
                 End Select
            End If
        Case 12
效果結束實行_手牌變化類:
            '==============結束技能實行(手牌變化類)
            Select Case atking_伊芙琳_赤紅石榴階段紀錄數(0, 1)
                 Case 10
                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為0張。"
                        戰鬥系統類.自動捲軸捲動
                        atkingck(59, 2) = 0
                        執行動作_技能手動結束
                 Case 11
                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為8張。"
                        戰鬥系統類.自動捲軸捲動
                        atkingck(59, 2) = 0
                        執行動作_技能手動結束
                 Case 12
                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為15張。"
                        戰鬥系統類.自動捲軸捲動
                        atkingck(59, 2) = 0
                        執行動作_技能手動結束
            End Select
   End Select
End If
End Sub
Sub CC_滅菌空間()
If FormMainMode.personatk(1).Caption = "滅菌空間" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(33, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "C.C." Then
   Select Case atkingck(33, 1)
        Case 1
             If atkingpagetot(1, 4) >= 1 And atkingck(33, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(33, 2) = 1
            ElseIf atkingpagetot(1, 4) < 1 And atkingck(33, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(33, 2) = 0
            End If
        Case 2
            atkingtrn(1) = Val(atkingtrn(1)) + 1
        Case 3
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_滅菌空間_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7275
                   atkingno(i, 6) = 9480
                   atkingno(i, 7) = 33
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            For i = 1 To 3
                回復執行_使用者 1, i
            Next
            atkingck(33, 2) = 0
            '======================
               戰鬥系統類.執行動作_清除所有異常狀態_使用者
           '======================
   End Select
End If
End Sub
Sub CC_白銀戰機()
Dim bloodntot As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "白銀戰機" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(34, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "C.C." Then
   Select Case atkingck(34, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(1, 1) >= 2 And atkingpagetot(1, 5) >= 2 And atkingck(34, 2) = 0 Then
'             If atkingpagetot(1, 1) >= 1 And atkingck(34, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(34, 2) = 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
            ElseIf (atkingpagetot(1, 1) < 2 Or atkingpagetot(1, 5) < 2) And atkingck(34, 2) = 1 Then
'            ElseIf atkingpagetot(1, 1) < 1 And atkingck(34, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(34, 2) = 0
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
            End If
          End If
        Case 2
            atkingtrn(1) = Val(atkingtrn(1)) + 1
        Case 3
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_白銀戰機_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -720
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 34
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            atkingck(34, 2) = 0
            For i = 1 To 3
                If i = 1 Then
                    Randomize
                    bloodntot = Int(Rnd() * 3) + 0
                    If livecom(角色人物對戰人數(2, 2)) > 1 And bloodntot < livecom(角色人物對戰人數(2, 2)) Then
                       戰鬥系統類.傷害執行_技能直傷_電腦 bloodntot, 1
                    ElseIf livecom(角色人物對戰人數(2, 2)) = 2 And bloodntot = 2 Then
                       bloodntot = 1
                       戰鬥系統類.傷害執行_技能直傷_電腦 bloodntot, 1
                    End If
                Else
                    Randomize
                    bloodntot = Int(Rnd() * 3) + 0
                    戰鬥系統類.傷害執行_技能直傷_電腦 bloodntot, i
                End If
            Next
   End Select
End If
End Sub
Sub CC_原子之心()
If FormMainMode.personatk(3).Caption = "原子之心" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(36, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "C.C." Then
   Select Case atkingck(36, 1)
        Case 1
             If atkingpagetot(1, 2) >= 2 And atkingpagetot(1, 4) >= 2 And atkingck(36, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(36, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(36, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 2
            ElseIf (atkingpagetot(1, 2) < 2 Or atkingpagetot(1, 4) < 2) And atkingck(36, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(36, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(36, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 2
            End If
        Case 2
            '===========將所有技能無效化-電腦方(階段1)
            atkingtrn(2) = 0
            For i = 1 To UBound(atkingckai)
                 atkingckai(i, 2) = 0
            Next
        Case 3
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_原子之心_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6945
                   atkingno(i, 6) = 10050
                   atkingno(i, 7) = 36
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
              '=================更改數值為原骰數值
              FormMainMode.顯示列1.goi1 = 攻擊防禦骰子總數(3) + 2
              FormMainMode.顯示列1.goi2 = 攻擊防禦骰子總數(4)
              '===============
              Erase atking_AI_音音夢_成長模式狀態數
              atking_AI_蕾_守護模式狀態啟動值 = False
        Case 4
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 1)
            FormMainMode.personcomminijpg.小人物影子圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 2)
            FormMainMode.顯示列1.電腦方小人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 4)
            FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
            FormMainMode.personcomminijpg.小人物影子Left = Val(VBEPerson(2, 角色人物對戰人數(2, 2), 2, 1, 5))
            FormMainMode.personcomminijpg.小人物影子top差 = Val(VBEPerson(2, 角色人物對戰人數(2, 2), 2, 1, 6))
            FormMainMode.personcomminijpg.Visible = True
            Form6.jpgcom.大人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 3)
'            Select Case 角色人物對戰人數(2, 2)
'                Case 1
'                    FormMainMode.personcomminijpg.Visible = False
'                    FormMainMode.personcomminijpg.小人物圖片 = formsettingpersoncom.personmini.Text
'                    FormMainMode.personcomminijpg.小人物影子圖片 = formsettingpersoncom.personsmalldown.Text
'                    FormMainMode.顯示列1.電腦方小人物圖片 = formsettingpersoncom.personf.Text
'                    FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
'                    FormMainMode.personcomminijpg.小人物影子Left = Val(formsettingpersoncom.smalldownleft.Text)
'                    FormMainMode.personcomminijpg.小人物影子top差 = Val(formsettingpersoncom.smalldowntop.Text)
'                    FormMainMode.personcomminijpg.Visible = True
'                    Form6.jpgcom.大人物圖片 = formsettingpersoncom.personbig.Text
'                Case 2
'                    FormMainMode.personcomminijpg.Visible = False
'                    FormMainMode.personcomminijpg.小人物圖片 = formsettingpersoncom2.personmini.Text
'                    FormMainMode.personcomminijpg.小人物影子圖片 = formsettingpersoncom2.personsmalldown.Text
'                    FormMainMode.顯示列1.電腦方小人物圖片 = formsettingpersoncom2.personf.Text
'                    FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
'                    FormMainMode.personcomminijpg.小人物影子Left = Val(formsettingpersoncom2.smalldownleft.Text)
'                    FormMainMode.personcomminijpg.小人物影子top差 = Val(formsettingpersoncom2.smalldowntop.Text)
'                    FormMainMode.personcomminijpg.Visible = True
'                    Form6.jpgcom.大人物圖片 = formsettingpersoncom2.personbig.Text
'                Case 3
'                    FormMainMode.personcomminijpg.Visible = False
'                    FormMainMode.personcomminijpg.小人物圖片 = formsettingpersoncom3.personmini.Text
'                    FormMainMode.personcomminijpg.小人物影子圖片 = formsettingpersoncom3.personsmalldown.Text
'                    FormMainMode.顯示列1.電腦方小人物圖片 = formsettingpersoncom3.personf.Text
'                    FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
'                    FormMainMode.personcomminijpg.小人物影子Left = Val(formsettingpersoncom3.smalldownleft.Text)
'                    FormMainMode.personcomminijpg.小人物影子top差 = Val(formsettingpersoncom3.smalldowntop.Text)
'                    FormMainMode.personcomminijpg.Visible = True
'                    Form6.jpgcom.大人物圖片 = formsettingpersoncom3.personbig.Text
'            End Select
            戰鬥系統類.執行動作_距離變更 movecp
            atkingck(36, 2) = 0
   End Select
End If
End Sub
Sub CC_高頻電磁手術刀()
If FormMainMode.personatk(4).Caption = "高頻電磁手術刀" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(35, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "C.C." Then
   Select Case atkingck(35, 1)
        Case 1
            If movecp = 1 Then
                If atkingpagetot(1, 4) >= 6 And atkingck(35, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingck(35, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 24
                ElseIf atkingpagetot(1, 4) < 6 And atkingck(35, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(35, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 24
                End If
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             atkingck(35, 1) = 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_高頻電磁手術刀_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9630
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
          '===========雙方中異常狀態
            Do
                For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 3) = 16 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                      FormMainMode.personusspe(i).person_turn = 3
                      人物異常狀態資料庫(1, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                   If 人物異常狀態資料庫(1, i, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 1, i, 16, app_path & "gif\異常狀態\moveerr.gif", 0, 3
                      異常狀態檢查數(16, 1) = 1
                      異常狀態檢查數(16, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            '===============
            Do
                For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 3) = 17 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                      FormMainMode.personcomspe(i).person_turn = 3
                      人物異常狀態資料庫(2, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                   If 人物異常狀態資料庫(2, i, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 2, i, 17, app_path & "gif\異常狀態\moveerr.gif", 0, 3
                      異常狀態檢查數(17, 1) = 1
                      異常狀態檢查數(17, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            atkingck(35, 2) = 0
   End Select
End If
End Sub
Sub 帕茉_憤怒之爪()
If FormMainMode.personatk(1).Caption = "憤怒之爪" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(7, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "帕茉" Then
   Select Case atkingck(7, 1)
      Case 1
            If atkingpagetot(1, 4) >= 1 And atkingck(7, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(7, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 4) < 1 And atkingck(7, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(7, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_憤怒之爪_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 1440
                   atkingno(i, 5) = 7050
                   atkingno(i, 6) = 9090
                   atkingno(i, 7) = 7
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
    Case 3
        Do
           atkingck(7, 2) = 0
           For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
             If 人物異常狀態資料庫(1, i, 2) >= 9 And 人物異常狀態資料庫(1, i, 3) = 13 Then
                Exit Do
             End If
             If 人物異常狀態資料庫(1, i, 3) = 13 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 9 Then
                 FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) + 1
                 人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 1
                 Exit Do
             End If
           Next
           For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
              If 人物異常狀態資料庫(1, i, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 1, i, 13, app_path & "gif\異常狀態\聖痕.gif", 0, 1
                 異常狀態檢查數(13, 1) = 1
                 異常狀態檢查數(13, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub 帕茉_靜謐之背()
If FormMainMode.personatk(2).Caption = "靜謐之背" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(17, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "帕茉" Then
   Select Case atkingck(17, 1)
      Case 1
         If movecp < 3 Then
            If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 2 And atkingck(17, 2) = 0 Then
'            If atkingpagetot(1, 1) >= 1 And atkingck(17, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(17, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 2) And atkingck(17, 2) = 1 Then
'            ElseIf atkingpagetot(1, 1) < 1 And atkingck(17, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(17, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             atkingck(17, 1) = 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_靜謐之背.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8340
                   atkingno(i, 6) = 8520
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
      Case 3
         If atkingck(17, 2) = 1 Then
             For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, i, 3) = 13 And 人物異常狀態資料庫(1, i, 2) >= 1 Then
                     Form6.技能_帕茉_靜謐之背_分支_階段三 (人物異常狀態資料庫(1, i, 2))
                     atkingck(17, 1) = 4
                     Exit For
                 End If
             Next
         End If
      Case 4
           atkingck(17, 2) = 0
           For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
             If 人物異常狀態資料庫(1, i, 3) = 13 And 人物異常狀態資料庫(1, i, 2) >= 1 Then
                 FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) - 1
                 人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
                 If 人物異常狀態資料庫(1, i, 2) = 0 Then
                     '===繼承下一狀態資料
                     戰鬥系統類.異常狀態繼承_使用者
                 End If
                 Exit For
             End If
           Next
   End Select
End If
End Sub
Sub 帕茉_慈悲的藍眼()
If FormMainMode.personatk(3).Caption = "慈悲的藍眼" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(9, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "帕茉" Then
   Select Case atkingck(9, 1)
      Case 1
          If movecp > 1 Then
             If atkingpagetot(1, 1) >= 6 And atkingck(9, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(9, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 1) < 6 And atkingck(9, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(9, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
          End If
      Case 2
          atking_帕茉_慈悲的藍眼_tot(1) = atking_帕茉_慈悲的藍眼_tot(1) + 攻擊防禦骰子總數(1)
          攻擊防禦骰子總數(1) = 0
          atking_帕茉_慈悲的藍眼_tot(2) = 1
          atkingck(9, 1) = 1
      Case 3
          atking_帕茉_慈悲的藍眼_tot(1) = atking_帕茉_慈悲的藍眼_tot(1) + 攻擊防禦骰子總數(1)
          攻擊防禦骰子總數(1) = atking_帕茉_慈悲的藍眼_tot(1)
          atking_帕茉_慈悲的藍眼_tot(1) = 0
          atking_帕茉_慈悲的藍眼_tot(2) = 0
          atkingck(9, 1) = 1
      Case 4
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_慈悲的藍眼_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6945
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 9
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 5
            Do
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, i, 2) >= 9 And 人物異常狀態資料庫(1, i, 3) = 13 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(1, i, 3) = 13 And 人物異常狀態資料庫(1, i, 2) = 8 Then
                     FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) + 1
                     人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 1
                     Exit Do
                 ElseIf 人物異常狀態資料庫(1, i, 3) = 13 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) <= 7 Then
'                 If 人物異常狀態資料庫(1, i, 3) = 13 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) <= 97 Then
                     FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2) + 2
                     人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 2
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 1, i, 13, app_path & "gif\異常狀態\聖痕.gif", 0, 2
                     異常狀態檢查數(13, 1) = 1
                     異常狀態檢查數(13, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
            '================
            回復執行_使用者 2, 1
            '================
            atkingck(9, 2) = 0
            atkingck(9, 1) = 0
            Erase atking_帕茉_慈悲的藍眼_tot
   End Select
End If
End Sub
Sub 帕茉_戰慄的狼牙()
Dim rrr As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "戰慄的狼牙" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(18, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "帕茉" Then
   Select Case atkingck(18, 1)
      Case 1
         If movecp = 1 Then
            If atkingpagetot(1, 1) >= 6 And atkingck(18, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(18, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 1) < 6 And atkingck(18, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(18, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_戰慄的狼牙_1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6645
                   atkingno(i, 6) = 9330
                   atkingno(i, 7) = 18
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
     Case 3
           For rrr = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
             If 人物異常狀態資料庫(1, rrr, 3) = 13 Then
                回復執行_使用者 人物異常狀態資料庫(1, rrr, 2), 1
                戰鬥系統類.傷害執行_技能直傷_電腦 人物異常狀態資料庫(1, rrr, 2), 1
                Exit For
             End If
           Next
            '=====================
               執行動作_清除所有異常狀態_使用者
               執行動作_清除所有異常狀態_電腦
           '======================
           atkingck(18, 2) = 0
   End Select
End If
End Sub
Sub 史塔夏_殺戮器官()
If FormMainMode.personatk(1).Caption = "殺戮器官" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(21, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "史塔夏" Then
   Select Case atkingck(21, 1)
        Case 1
            If pageqlead(1) >= 3 And atkingck(21, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(21, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf pageqlead(1) < 3 And atkingck(21, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(21, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_殺戮器官.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 21
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciamini1.png"
            FormMainMode.personusminijpg.小人物影子圖片 = app_path & "gif\史塔夏\殺戮\Staciaminidown1.png"
            FormMainMode.personusminijpg.小人物影子Left = -90
            FormMainMode.personusminijpg.小人物影子top差 = -60
            Form6.jpgus.大人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaperson1.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaf1.png"
            atking_史塔夏_殺戮模式狀態數(2) = 1
            atkingck(21, 2) = 0
'            formsettingpersonus.smalldownleft = -90
'            formsettingpersonus.smalldowntop = -60
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personusminijpg.Visible = True
   End Select
End If
End Sub
Sub 史塔夏_愚者之手()
Dim apn As Integer
If FormMainMode.personatk(2).Caption = "愚者之手" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(23, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "史塔夏" Then
   Select Case atkingck(23, 1)
        Case 1
            If movecp < 3 Then
             For i = 1 To 3
                 If livecom(i) > 0 Then
                     apn = apn + 1
                 End If
             Next
             If atkingpagetot(1, 1) >= 6 And atkingck(23, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(23, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + apn * 4
            ElseIf atkingpagetot(1, 1) < 6 And atkingck(23, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(23, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - apn * 4
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_愚者之手.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 23
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingck(23, 1) = 3
        Case 3
            For i = 1 To 3
                 If livecom(i) > 0 Then
                     apn = apn + 1
                 End If
            Next
            If atking_史塔夏_殺戮模式狀態數(2) = 1 Then
                戰鬥系統類.傷害執行_技能直傷_使用者 apn, 1
            End If
            atkingck(23, 2) = 0
   End Select
End If
End Sub

Sub 史塔夏_時間種子()
Dim bloodtot As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "時間種子" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(24, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "史塔夏" Then
   Select Case atkingck(24, 1)
        Case 1
            If movecp < 3 Then
             If atkingpagetot(1, 2) >= 2 And atkingpagetot(1, 4) >= 2 And atkingck(24, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(24, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(24, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 2) < 2 Or atkingpagetot(1, 4) < 2) And atkingck(24, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(24, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(24, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_時間種子.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 24
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            bloodtot = Val(FormMainMode.turni) \ 2
            If bloodtot > 4 Then bloodtot = 4
            '=============
            If Val(liveus(角色人物對戰人數(1, 2))) < Val(liveusmax(角色人物對戰人數(1, 2))) Then
               Select Case atking_史塔夏_殺戮模式狀態數(2)
                   Case 0
                        回復執行_使用者 bloodtot, 1
                   Case 1
                        回復執行_使用者 bloodtot \ 2, 1
                  End Select
            End If
            atkingck(24, 2) = 0
   End Select
End If
End Sub
Sub 史塔夏_命運的鐵門()
Dim num(1 To 2, 1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.personatk(4).Caption = "命運的鐵門" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(25, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "史塔夏" Then
   Select Case atkingck(25, 1)
        Case 1
         If movecp = 3 Then
             If atkingpagetot(1, 1) >= 9 And atkingck(25, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(25, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 1) < 9 And atkingck(25, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(25, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
         End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_命運的鐵門.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6720
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 25
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(liveus(角色人物對戰人數(1, 2))) < Val(liveusmax(角色人物對戰人數(1, 2))) Then
               If atking_史塔夏_殺戮模式狀態數(2) = 1 Then
                  回復執行_使用者 3, 1
               End If
            End If
            atkingck(25, 1) = 4
        Case 4
           num(1, 2) = 999 '目的取最低HP數
           num(2, 2) = 999
           For i = 2 To 3
               If Val(FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption) < num(1, 2) And Val(FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption) > 0 Then
                   num(1, 1) = i
                   num(1, 2) = FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption
               End If
            Next
            For i = 1 To 3
               If livecom(角色待機人物紀錄數(2, i)) < num(2, 2) And livecom(角色待機人物紀錄數(2, i)) > 0 Then
                   num(2, 1) = i
                   num(2, 2) = livecom(角色待機人物紀錄數(2, i))
               End If
           Next
           If num(1, 2) < num(2, 2) Or num(1, 2) = num(2, 2) Then
               戰鬥系統類.傷害執行_立即死亡_使用者 num(1, 1)
           Else
               戰鬥系統類.傷害執行_立即死亡_電腦 num(2, 1)
           End If
           atkingck(25, 2) = 0
   End Select
End If
End Sub
Sub 羅莎琳_黑霧幻影()
If FormMainMode.personatk(1).Caption = "黑霧幻影" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(54, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "羅莎琳" Then
   Select Case atkingck(54, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(1, 2) >= 3 And atkingpagetot(1, 4) >= 2 And atkingck(54, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(54, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
            ElseIf (atkingpagetot(1, 2) < 3 Or atkingpagetot(1, 4) < 2) And atkingck(54, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(54, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_黑霧幻影_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 54
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_羅莎琳_黑霧幻影紀錄狀態數(i) = True
'                   目前數(18) = 目前數(18) + 1
               End If
            Next
            目前數(18) = 1
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) <= 0 Then
                Do
                    If atking_羅莎琳_黑霧幻影紀錄狀態數(目前數(18)) = True Then
                        目前數(16) = 目前數(18)
                        目前數(15) = 21
'                        tr牌組_翻牌.Enabled = True
                        FormMainMode.tr牌組_回牌_使用者.Enabled = True
                        atking_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
                        Exit Do
                    End If
                    目前數(18) = 目前數(18) + 1
                Loop Until 目前數(18) >= 106
            End If
            If 目前數(18) >= 106 Or Val(擲骰表單溝通暫時變數(2)) > 0 Then
                atkingck(54, 1) = 6
                FormMainMode.骰子執行完啟動.Enabled = True
            End If
        Case 5
'            tr牌組_回牌_使用者.Enabled = True
'            atking_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
        Case 6
            atkingck(54, 2) = 0
            Erase atking_羅莎琳_黑霧幻影紀錄狀態數
   End Select
End If
End Sub
Sub 羅莎琳_EX_黑霧幻影()
If FormMainMode.personatk(1).Caption = "Ex黑霧幻影" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(55, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "羅莎琳" Then
   Select Case atkingck(55, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(1, 2) >= 4 And atkingpagetot(1, 4) >= 2 And atkingck(55, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(55, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 9
            ElseIf (atkingpagetot(1, 2) < 4 Or atkingpagetot(1, 4) < 2) And atkingck(55, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(55, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 9
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_Ex-黑霧幻影_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 55
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_羅莎琳_黑霧幻影紀錄狀態數(i) = True
'                   目前數(18) = 目前數(18) + 1
               End If
            Next
            目前數(18) = 1
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) <= 0 Then
                Do
                    If atking_羅莎琳_黑霧幻影紀錄狀態數(目前數(18)) = True Then
                        目前數(16) = 目前數(18)
                        目前數(15) = 21
'                        tr牌組_翻牌.Enabled = True
                        FormMainMode.tr牌組_回牌_使用者.Enabled = True
                        atking_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
                        Exit Do
                    End If
                    目前數(18) = 目前數(18) + 1
                Loop Until 目前數(18) >= 106
            End If
            If 目前數(18) >= 106 Or Val(擲骰表單溝通暫時變數(2)) > 0 Then
                atkingck(55, 1) = 6
                FormMainMode.骰子執行完啟動.Enabled = True
            End If
        Case 5
'            tr牌組_回牌_使用者.Enabled = True
'            atking_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
        Case 6
            atkingck(55, 2) = 0
            Erase atking_羅莎琳_黑霧幻影紀錄狀態數
   End Select
End If
End Sub

Sub 羅莎琳_染血之刃()
If FormMainMode.personatk(2).Caption = "染血之刃" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(51, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "羅莎琳" Then
   Select Case atkingck(51, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 3) >= 1 And atkingck(51, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(51, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
            ElseIf (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 3) < 1) And atkingck(51, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(51, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_染血之刃_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 51
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            回復執行_使用者 1, 1
        Case 4
            atkingck(51, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                回復執行_使用者 1, 1
            End If
   End Select
End If
End Sub
Sub 羅莎琳_咀咒的刻印()
If FormMainMode.personatk(3).Caption = "咀咒的刻印" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(53, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "羅莎琳" Then
   Select Case atkingck(53, 1)
        Case 1
            If movecp > 1 Then
                If atkingpagetot(1, 2) >= 5 And atkingpagetot(1, 4) >= 1 And atkingck(53, 2) = 0 Then
    '             If atkingpagetot(1, 2) >= 1 And atkingck(24, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingck(53, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf (atkingpagetot(1, 2) < 5 Or atkingpagetot(1, 4) < 1) And atkingck(53, 2) = 1 Then
    '            ElseIf atkingpagetot(1, 2) < 1 And atkingck(24, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(53, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             atkingck(53, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_咀咒的刻印_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6975
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 53
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
             If atkingpagetot(2, 4) >= 1 Then
                 戰鬥系統類.直接寫入顯示列數值 2, Int(Val(FormMainMode.顯示列1.goi2) / 3 + 0.9)
             Else
                 戰鬥系統類.直接寫入顯示列數值 2, Int(Val(FormMainMode.顯示列1.goi2) / 2 + 0.9)
             End If
'             攻擊防禦骰子總數(2) = FormMainMode.顯示列1.goi2
   End Select
End If
End Sub
Sub 羅莎琳_黑霧的纏繞()
Dim m As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "黑霧的纏繞" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(52, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "羅莎琳" Then
   Select Case atkingck(52, 1)
        Case 1
            If atkingpagetot(1, 4) >= 2 And atkingck(52, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(52, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
            ElseIf atkingpagetot(1, 4) < 2 And atkingck(52, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(52, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_黑霧的纏繞_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = -240
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6390
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 52
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(52, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                       Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 21 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 2
                                  人物異常狀態資料庫(2, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 21, app_path & "gif\異常狀態\damage.gif", 0, 2
                                  異常狀態檢查數(21, 1) = 1
                                  異常狀態檢查數(21, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 17 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 2
                                  人物異常狀態資料庫(2, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 17, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                                  異常狀態檢查數(17, 1) = 1
                                  異常狀態檢查數(17, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 23 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 2
                                  人物異常狀態資料庫(2, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 23, app_path & "gif\異常狀態\atkingerr.gif", 0, 2
                                  異常狀態檢查數(23, 1) = 1
                                  異常狀態檢查數(23, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
            End If
   End Select
End If
End Sub
Sub 梅倫_High_hand()
If FormMainMode.personatk(1).Caption = "High hand" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(63, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅倫" Then
   Select Case atkingck(63, 1)
        Case 1
             If atkingpagetot(1, 4) >= 2 And atkingck(63, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(63, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + pageqlead(2) * 2
            ElseIf atkingpagetot(1, 4) < 2 And atkingck(63, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(63, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(2) * 2
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             atkingck(63, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\High hand_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7770
                   atkingno(i, 6) = 10020
                   atkingno(i, 7) = 63
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 梅倫_Jackpot()
Dim m As Integer
If FormMainMode.personatk(2).Caption = "Jackpot" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(64, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅倫" Then
   Select Case atkingck(64, 1)
        Case 1
            If movecp = 2 Then
                If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 2) >= 1 And atkingck(64, 2) = 0 Then
'                If atkingpagetot(1, 2) >= 1 And atkingck(64, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingck(64, 2) = 1
                ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 5) < 1 Or atkingpagetot(1, 2) < 1) And atkingck(64, 2) = 1 Then
'                ElseIf atkingpagetot(1, 2) < 1 And atkingck(64, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(64, 2) = 0
                End If
            End If
        Case 2
             atking_梅倫_Jackpot紀錄數(1) = pageqlead(1) * 2
             atking_梅倫_Jackpot紀錄數(2) = 1
        Case 3
             atkingtrn(1) = Val(atkingtrn(1)) + 1
        Case 4
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\Jackpot_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6480
                   atkingno(i, 6) = 10020
                   atkingno(i, 7) = 64
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
             If Val(FormMainMode.pageul.Caption) < atking_梅倫_Jackpot紀錄數(1) And atking_梅倫_Jackpot紀錄數(2) = 1 Then
               戰鬥系統類.執行動作_洗牌
             End If
             If atking_梅倫_Jackpot紀錄數(2) > atking_梅倫_Jackpot紀錄數(1) Then
                 atkingck(64, 2) = 0
                 戰鬥系統類.執行動作_技能手動結束
            Else
                目前數(15) = 22
                FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                atking_梅倫_Jackpot紀錄數(2) = atking_梅倫_Jackpot紀錄數(2) + 1
            End If
   End Select
End If
End Sub
Sub 梅倫_Lowball()
If FormMainMode.personatk(3).Caption = "Lowball" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(65, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅倫" Then
   Select Case atkingck(65, 1)
        Case 1
             If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 _
                And atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 5) >= 1 And atkingck(65, 2) = 0 Then
'            If atkingpagetot(1, 1) >= 1 And atkingck(65, 2) = 0 Then
                    戰鬥系統類.人物技能欄燈開關 True, 3
                    atkingck(65, 2) = 1
                    atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1 _
               Or atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 5) < 1) And atkingck(65, 2) = 1 Then
'            ElseIf atkingpagetot(1, 1) < 1 And atkingck(65, 2) = 1 Then
                    戰鬥系統類.人物技能欄燈開關 False, 3
                    atkingck(65, 2) = 0
                    atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\Lowball_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7020
                   atkingno(i, 6) = 9555
                   atkingno(i, 7) = 65
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) + 5
             擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
             atkingck(65, 2) = 0
   End Select
End If
End Sub
Sub 梅倫_Gamble()
If FormMainMode.personatk(4).Caption = "Gamble" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(66, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅倫" Then
   Select Case atkingck(66, 1)
        Case 1
            If movecp = 1 Then
                 If pageqlead(1) >= 3 And atkingck(66, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingck(66, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf pageqlead(1) < 3 And atkingck(66, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(66, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\Gamble_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6960
                   atkingno(i, 6) = 9780
                   atkingno(i, 7) = 66
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(66, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) = 1 Then
                 戰鬥系統類.傷害執行_立即死亡_電腦 1
             End If
   End Select
End If
End Sub
Sub 音音夢_美味牛奶()
If FormMainMode.personatk(1).Caption = "美味牛奶" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(67, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "音音夢" Then
   Select Case atkingck(67, 1)
        Case 1
            If pageqlead(1) >= 2 And atkingck(67, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(67, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf pageqlead(1) < 2 And atkingck(67, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(67, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\音音夢\atking1-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6360
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 67
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\音音夢\atking1-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(67, 2) = 0
            '==============================
            For k = 2 To 3
                傷害執行_技能直傷_使用者 1, k
            Next
            '==============================
            atking_音音夢_成長模式狀態數(2) = 1
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = app_path & "gif\音音夢\成長\Nenemmini1.png"
            FormMainMode.personusminijpg.小人物影子圖片 = app_path & "gif\音音夢\成長\Nenemminidown1.png"
            FormMainMode.personusminijpg.小人物影子Left = 20
            FormMainMode.personusminijpg.小人物影子top差 = -90
            Form6.jpgus.大人物圖片 = app_path & "gif\音音夢\成長\Nenemperson1.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\音音夢\成長\Nenemf1.png"
'            formsettingpersonus.smalldownleft = 20
'            formsettingpersonus.smalldowntop = -90
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personusminijpg.Visible = True
   End Select
End If
End Sub
Sub 音音夢_溫柔注射()
Dim n(1 To 2) As Integer
If FormMainMode.personatk(2).Caption = "溫柔注射" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(68, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "音音夢" Then
   Select Case atkingck(68, 1)
        Case 1
            If movecp < 3 Then
             If atkingpagetot(1, 1) >= 2 And atkingpagetot(1, 2) >= 2 And atkingck(68, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(68, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(68, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
            ElseIf (atkingpagetot(1, 1) < 2 Or atkingpagetot(1, 2) < 2) And atkingck(68, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(68, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(68, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\音音夢\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6165
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 68
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(68, 2) = 0
            '=======================
            n(1) = 999 '取最小HP值
            n(2) = 0
            For i = 2 To 3
                If liveus(角色待機人物紀錄數(1, i)) > 0 And liveus(角色待機人物紀錄數(1, i)) < n(1) Then
                    n(1) = liveus(角色待機人物紀錄數(1, i))
                    n(2) = i
                End If
            Next
            If n(2) > 0 Then
                If liveus(角色人物對戰人數(1, 2)) >= n(1) Then
                    戰鬥系統類.回復執行_使用者 liveus(角色人物對戰人數(1, 2)) - n(1), n(2)
                Else
                    戰鬥系統類.傷害執行_技能直傷_使用者 n(1) - liveus(角色人物對戰人數(1, 2)), n(2)
                End If
            End If
   End Select
End If
End Sub
Sub 音音夢_愉快抽血(ByVal Index As Integer)
Dim n(1 To 2) As Integer
If FormMainMode.personatk(3).Caption = "愉快抽血" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(69, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "音音夢" Then
 Select Case atkingck(69, 1)
    Case 1
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2)) * 5
               If atkingck(69, 2) = 0 And atkingpagetot(1, 4) > 0 Then
                  戰鬥系統類.人物技能欄燈開關 True, 3
                  atkingck(69, 2) = 1
                  atkingtrn(1) = Val(atkingtrn(1)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 2)) * 5
               If atkingck(69, 2) = 1 And atkingpagetot(1, 4) = 0 Then
                  戰鬥系統類.人物技能欄燈開關 False, 3
                  atkingck(69, 2) = 0
                  atkingtrn(1) = Val(atkingtrn(1)) - 1
               End If
        End If
        FormMainMode.trgoi1.Enabled = True
    Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Val(pagecardnum(Index, 2)) * 5
               If atkingck(69, 2) = 0 And atkingpagetot(1, 4) > 0 Then
                  戰鬥系統類.人物技能欄燈開關 True, 3
                  atkingck(69, 2) = 1
                  atkingtrn(1) = Val(atkingtrn(1)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - Val(pagecardnum(Index, 4)) * 5
               If atkingck(69, 2) = 1 And atkingpagetot(1, 4) = 0 Then
                  戰鬥系統類.人物技能欄燈開關 False, 3
                  atkingck(69, 2) = 0
                  atkingtrn(1) = Val(atkingtrn(1)) - 1
               End If
        End If
        FormMainMode.trgoi1.Enabled = True
    Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\音音夢\atking3_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6645
                atkingno(i, 6) = 9555
                atkingno(i, 7) = 69
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
        戰鬥系統類.人物技能欄燈開關 False, 3
       '-------------
    Case 4
       atkingck(69, 2) = 0
        n(1) = 999 '取最小HP值
        n(2) = 0
        For i = 2 To 3
            If liveus(角色待機人物紀錄數(1, i)) > 0 And liveus(角色待機人物紀錄數(1, i)) < n(1) Then
                n(1) = liveus(角色待機人物紀錄數(1, i))
                n(2) = i
            End If
        Next
        If n(2) > 0 Then
            戰鬥系統類.傷害執行_技能直傷_使用者 Val(atkingpagetot(1, 4)), n(2)
        Else
            戰鬥系統類.傷害執行_技能直傷_使用者 Val(atkingpagetot(1, 4)), 1
        End If
  End Select
End If
End Sub
Sub 音音夢_秘密苦藥()
If FormMainMode.personatk(4).Caption = "秘密苦藥" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(70, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "音音夢" Then
   Select Case atkingck(70, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 2) >= 1 And atkingck(70, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(70, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(70, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 2) < 1) And atkingck(70, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(70, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(70, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\音音夢\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6750
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 70
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(70, 2) = 0
            '=======================
            For i = 2 To 3
                戰鬥系統類.回復執行_使用者 10, i
            Next
            戰鬥系統類.傷害執行_立即死亡_使用者 1
            '=======================
            If atking_音音夢_成長模式狀態數(2) = 1 Then
                牌總階段數(1) = 牌總階段數(1) + 1
            End If
   End Select
End If
End Sub
Sub 艾伯李斯特_精密射擊()
If FormMainMode.personatk(1).Caption = "精密射擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(71, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾伯李斯特" Then
   Select Case atkingck(71, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 2 And atkingck(71, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
                   atkingck(71, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 5) < 2 And atkingck(71, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(71, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(71, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6525
                   atkingno(i, 6) = 10110
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 艾伯李斯特_雷擊()
If FormMainMode.personatk(2).Caption = "雷擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(72, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾伯李斯特" Then
   Select Case atkingck(72, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 4) >= 2 And atkingck(72, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
                   atkingck(72, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 2 And atkingck(72, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(72, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -240
                   atkingno(i, 5) = 9795
                   atkingno(i, 6) = 10215
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(擲骰表單溝通暫時變數(2)) > 0 And Val(FormMainMode.pagecomglead.Caption) > 0 Then
                 atking_艾伯李斯特_雷擊紀錄數(1) = Val(擲骰表單溝通暫時變數(2))
                 atking_艾伯李斯特_雷擊紀錄數(2) = 1
                 '==========================
                  Do Until atking_艾伯李斯特_雷擊紀錄數(2) > atking_艾伯李斯特_雷擊紀錄數(1) Or Val(FormMainMode.pagecomglead.Caption) <= 0
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                            目前數(17) = 7
                            目前數(16) = m
                            atking_艾伯李斯特_雷擊紀錄數(2) = atking_艾伯李斯特_雷擊紀錄數(2) + 1
                            FormMainMode.tr電腦牌_翻牌.Enabled = True
                            Exit Sub
                        End If
                   Loop
             Else
                 atkingck(72, 1) = 5
                 FormMainMode.骰子執行完啟動.Enabled = True
             End If
        Case 4
             Do Until atking_艾伯李斯特_雷擊紀錄數(2) > atking_艾伯李斯特_雷擊紀錄數(1) Or Val(FormMainMode.pagecomglead.Caption) <= 0
                 Randomize
                 m = Int(Rnd() * 106) + 1
                 If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                     目前數(17) = 7
                     目前數(16) = m
                     atking_艾伯李斯特_雷擊紀錄數(2) = atking_艾伯李斯特_雷擊紀錄數(2) + 1
                     FormMainMode.tr電腦牌_翻牌.Enabled = True
                     Exit Sub
                 End If
            Loop
            If atking_艾伯李斯特_雷擊紀錄數(2) > atking_艾伯李斯特_雷擊紀錄數(1) Or Val(FormMainMode.pagecomglead.Caption) <= 0 Then
                atkingck(72, 1) = 5
                目前數(24) = 22
                FormMainMode.等待時間_2.Enabled = True
            End If
        Case 5
            atkingck(72, 2) = 0
            Erase atking_艾伯李斯特_雷擊紀錄數
        Case 6
            FormMainMode.tr電腦牌_棄牌.Enabled = True
   End Select
End If
End Sub
Sub 艾伯李斯特_茨林()
If FormMainMode.personatk(3).Caption = "茨林" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(73, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾伯李斯特" Then
   Select Case atkingck(73, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 4) >= 2 And atkingpagetot(1, 2) >= 2 And atkingck(73, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 7
                   atkingck(73, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 4) < 2 Or atkingpagetot(1, 2) < 2) And atkingck(73, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 7
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(73, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking3-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1200
                   atkingno(i, 5) = 6705
                   atkingno(i, 6) = 10245
                   atkingno(i, 7) = 73
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾伯李斯特\atking3-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(FormMainMode.顯示列1.goi2) <= 0 Then
                atkingck(73, 2) = 0
            End If
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) < 0 Then
                戰鬥系統類.傷害執行_技能直傷_電腦 Abs(擲骰表單溝通暫時變數(2)), 1
            End If
            atkingck(73, 2) = 0
   End Select
End If
End Sub
Sub 艾伯李斯特_智略()
If FormMainMode.personatk(4).Caption = "智略" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(74, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾伯李斯特" Then
   Select Case atkingck(74, 1)
      Case 1
            If pageqlead(1) >= 3 And atkingck(74, 2) = 0 Then
               atkingck(74, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If pageqlead(1) < 3 And atkingck(74, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(74, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6060
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 74
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(FormMainMode.pageul.Caption) < 2 And atking_艾伯李斯特_智略紀錄數 = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_艾伯李斯特_智略紀錄數 = atking_艾伯李斯特_智略紀錄數 + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_艾伯李斯特_智略紀錄數 > 2
                    目前數(15) = 23
                    FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_艾伯李斯特_智略紀錄數 > 2 Or Val(FormMainMode.pageul.Caption) <= 0 Then
               atking_艾伯李斯特_智略紀錄數 = 0
               atkingck(74, 2) = 0
            End If
   End Select
End If
End Sub
Sub 艾依查庫_連射()
If FormMainMode.personatk(1).Caption = "連射" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(78, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾依查庫" Then
   Select Case atkingck(78, 1)
      Case 1
           If movecp > 1 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr = rrr + 1
                End If
             Next
          End If
          If rrr >= 2 And atkingck(78, 2) = 0 Then
             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
             atkingck(78, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 1
             atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
          If rrr < 2 And atkingck(78, 2) = 1 Then
             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(78, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
           End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(78, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking1-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6780
                   atkingno(i, 6) = 10185
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾依查庫\atking1-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 艾依查庫_神速之劍(ByVal Index As Integer)
Dim aw As Integer
If FormMainMode.personatk(2).Caption = "神速之劍" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(79, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾依查庫" Then
   Select Case atkingck(79, 1)
      Case 1
             If movecp > 1 Then
                 If atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 1) >= 2 And atkingck(79, 2) = 0 Then
                     aw = Int(atkingpagetot(1, 1) / 2 + 0.5)
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_艾依查庫_神速之劍計算數值紀錄數(1))
                     atking_艾依查庫_神速之劍計算數值紀錄數(1) = aw
                     戰鬥系統類.人物技能欄燈開關 True, 2
                     atkingck(79, 2) = 1
                     atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
            End If
      Case 2
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 And atkingck(79, 2) = 1 Then
                   aw = Int(atkingpagetot(1, 1) / 2 + 0.5)
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_艾依查庫_神速之劍計算數值紀錄數(1))
                   atking_艾依查庫_神速之劍計算數值紀錄數(1) = aw
            End If
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 And atkingck(79, 2) = 1 Then
                   If atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 1) >= 2 And atkingck(79, 2) = 1 Then
                        aw = Int(atkingpagetot(1, 1) / 2 + 0.5)
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - (atking_艾依查庫_神速之劍計算數值紀錄數(1) - aw)
                        atking_艾依查庫_神速之劍計算數值紀錄數(1) = aw
                   ElseIf (atkingpagetot(1, 5) < 1 Or atkingpagetot(1, 1) < 2) And atkingck(79, 2) = 1 Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - atking_艾依查庫_神速之劍計算數值紀錄數(1)
                        戰鬥系統類.人物技能欄燈開關 False, 2
                        atkingck(79, 2) = 0
                        atkingck(79, 1) = 1
                        atkingtrn(1) = Val(atkingtrn(1)) - 1
                        Erase atking_艾依查庫_神速之劍計算數值紀錄數
                    End If
            End If
            FormMainMode.trgoi1.Enabled = True
    Case 3
        If Val(pagecardnum(Index, 5)) = 1 And atkingck(79, 2) = 1 Then
               If atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 1) >= 2 Then
                    aw = Int(atkingpagetot(1, 1) / 2 + 0.5)
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_艾依查庫_神速之劍計算數值紀錄數(1))
                    atking_艾依查庫_神速之劍計算數值紀錄數(1) = aw
               ElseIf (atkingpagetot(1, 5) < 1 Or atkingpagetot(1, 1) < 2) Then
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - atking_艾依查庫_神速之劍計算數值紀錄數(1)
                    戰鬥系統類.人物技能欄燈開關 False, 2
                    atkingck(79, 2) = 0
                    atkingck(79, 1) = 1
                    atkingtrn(1) = Val(atkingtrn(1)) - 1
                    Erase atking_艾依查庫_神速之劍計算數值紀錄數
                End If
        End If
        FormMainMode.trgoi1.Enabled = True
      Case 4
             戰鬥系統類.人物技能欄燈開關 False, 2
             atkingck(79, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking2-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 9165
                   atkingno(i, 6) = 10350
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾依查庫\atking2-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             Erase atking_艾依查庫_神速之劍計算數值紀錄數
   End Select
End If
End Sub
Sub 艾依查庫_憤怒一擊()
Dim ape As Integer
If FormMainMode.personatk(3).Caption = "憤怒一擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(80, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾依查庫" Then
   Select Case atkingck(80, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 4) >= 3 And atkingck(80, 2) = 0 Then
                   ape = (liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2))) * 2
                   If ape > 16 Then ape = 16
                   atkingck(80, 2) = 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + ape
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 3 And atkingck(80, 2) = 1 Then
                   ape = (liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2))) * 2
                   If ape > 16 Then ape = 16
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - ape
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(80, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             atkingck(80, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking3-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6615
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾依查庫\atking3-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 艾依查庫_不屈之心()
If FormMainMode.personatk(4).Caption = "不屈之心" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(81, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾依查庫" Then
   Select Case atkingck(81, 1)
      Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr = rrr + 1
                End If
             Next
          If rrr >= 2 And atkingck(81, 2) = 0 Then
             atkingck(81, 2) = 1
             戰鬥系統類.人物技能欄燈開關 True, 4
             atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
          If rrr < 2 And atkingck(81, 2) = 1 Then
             戰鬥系統類.人物技能欄燈開關 False, 4
             atkingck(81, 2) = 0
             atkingtrn(1) = Val(atkingtrn(1)) - 1
           End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5760
                   atkingno(i, 6) = 9450
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(81, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) >= liveus(角色人物對戰人數(1, 2)) Then
                 擲骰表單溝通暫時變數(2) = liveus(角色人物對戰人數(1, 2)) - 1
                 擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
             End If
   End Select
End If
End Sub
Sub 布勞_發條機構()
Dim tn As Integer
If FormMainMode.personatk(1).Caption = "發條機構" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(82, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "布勞" Then
   Select Case atkingck(82, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 4) >= 2 And atkingck(82, 2) = 0 Then
                   atkingck(82, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 2 And atkingck(82, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(82, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7455
                   atkingno(i, 6) = 9075
                   atkingno(i, 7) = 82
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
              '=====================================分派下一張事件卡
                tn = Val(FormMainMode.turni) + 1
                If tn <= 18 Then
                    If tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgreus.Value = 0 Then
                        If pageeventnum(1, tn, 1) <> "" Then
                            ay = Split(一般系統類.事件卡資料庫(pageeventnum(1, tn, 1), 3), "=")
                            pagecardnum(70 + tn, 1) = ay(0)
                            pagecardnum(70 + tn, 2) = ay(1)
                            pagecardnum(70 + tn, 3) = ay(2)
                            pagecardnum(70 + tn, 4) = ay(3)
                            pagecardnum(70 + tn, 5) = 1
                            pagecardnum(70 + tn, 6) = 1
                            pagecardnum(70 + tn, 8) = pageeventnum(1, tn, 2)
                            pagecardnum(70 + tn, 11) = 0
                            FormMainMode.card(70 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(1, tn, 2) & "-1.bmp")
                            pageonin(70 + tn) = 1
                        End If
                    End If
                End If
             '=====================================
             If Val(FormMainMode.turni) < 18 And (tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgreus.Value = 0) Then
                目前數(16) = 70 + Val(FormMainMode.turni) + 1
                atking_布勞_發條機構紀錄數 = 1
                目前數(15) = 24
                FormMainMode.tr牌組_回牌_使用者.Enabled = True
            Else
                atkingck(82, 2) = 0
            End If
        Case 4
            If Val(FormMainMode.turni) + atking_布勞_發條機構紀錄數 < 18 And atking_布勞_發條機構紀錄數 < 2 And _
               (tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgreus.Value = 0) Then
               '=====================================分派下一張事件卡
                tn = Val(FormMainMode.turni) + 2
                If tn <= 18 Then
                    If tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgreus.Value = 0 Then
                        If pageeventnum(1, tn, 1) <> "" Then
                            ay = Split(一般系統類.事件卡資料庫(pageeventnum(1, tn, 1), 3), "=")
                            pagecardnum(70 + tn, 1) = ay(0)
                            pagecardnum(70 + tn, 2) = ay(1)
                            pagecardnum(70 + tn, 3) = ay(2)
                            pagecardnum(70 + tn, 4) = ay(3)
                            pagecardnum(70 + tn, 5) = 1
                            pagecardnum(70 + tn, 6) = 1
                            pagecardnum(70 + tn, 8) = pageeventnum(1, tn, 2)
                            pagecardnum(70 + tn, 11) = 0
                            FormMainMode.card(70 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(1, tn, 2) & "-1.bmp")
                            pageonin(70 + tn) = 1
                        End If
                    End If
                End If
                atking_布勞_發條機構紀錄數 = atking_布勞_發條機構紀錄數 + 1
                目前數(16) = 70 + Val(FormMainMode.turni) + 2
                目前數(15) = 24
                FormMainMode.tr牌組_回牌_使用者.Enabled = True
            Else
                FormMainMode.turni = Val(FormMainMode.turni) + atking_布勞_發條機構紀錄數
                turn = Val(FormMainMode.turni)
                atking_布勞_發條機構紀錄數 = 0
                atkingck(82, 2) = 0
            End If
   End Select
End If
End Sub
Sub 布勞_時間追獵()
If FormMainMode.personatk(2).Caption = "時間追獵" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(83, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "布勞" Then
   Select Case atkingck(83, 1)
        Case 1
             If movecp < 3 Then
                 If atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 2) >= 1 And atkingck(83, 2) = 0 Then
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingck(83, 2) = 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                ElseIf (atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 2) < 1) And atkingck(83, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(83, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                End If
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6510
                   atkingno(i, 6) = 9690
                   atkingno(i, 7) = 83
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=====================
            atkingck(83, 2) = 0
            戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - Val(FormMainMode.turni)
'            攻擊防禦骰子總數(2) = FormMainMode.顯示列1.goi2
   End Select
End If
End Sub
Sub 布勞_時間爆彈()
Dim tn As Integer
If FormMainMode.personatk(3).Caption = "時間爆彈" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(84, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "布勞" Then
   Select Case atkingck(84, 1)
        Case 1
             If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 3 And atkingck(84, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(84, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 7
            ElseIf (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 3) And atkingck(84, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(84, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 7
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking3-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7125
                   atkingno(i, 6) = 9330
                   atkingno(i, 7) = 84
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\布勞\atking3-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(84, 2) = 0
            tn = Val(FormMainMode.turni)
            If tn = 2 Or tn = 3 Or tn = 5 Or tn = 7 Or tn = 11 Or tn = 13 Or tn = 17 Then
               戰鬥系統類.傷害執行_技能直傷_電腦 3, 1
            End If
   End Select
End If
End Sub
Sub 布勞_夜幕時分()
Dim tn(1 To 3) As Boolean
If FormMainMode.personatk(4).Caption = "夜幕時分" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(85, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "布勞" Then
   Select Case atkingck(85, 1)
        Case 1
             If pageqlead(1) >= 3 And atkingck(85, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(85, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf pageqlead(1) < 3 And atkingck(85, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(85, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6870
                   atkingno(i, 6) = 10365
                   atkingno(i, 7) = 85
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
            atkingck(85, 2) = 0
            '======================
            For i = 1 To 3
                If VBEPerson(1, 角色待機人物紀錄數(1, i), 1, 2, 1) = "R" Then
                     tn(i) = True
                Else
                     tn(i) = False
                End If
                 If tn(i) = True Then
                     Select Case Val(VBEPerson(1, 角色待機人物紀錄數(1, i), 1, 2, 2))
                         Case Is <= 2
                              戰鬥系統類.回復執行_使用者 1, i
                         Case Is > 2, Is <= 4
                              戰鬥系統類.回復執行_使用者 2, i
                         Case 5
                              戰鬥系統類.回復執行_使用者 3, i
                     End Select
                 End If
            Next
            '=============================
   End Select
End If
End Sub
Sub 阿貝爾_霸王閃擊()
If FormMainMode.personatk(1).Caption = "霸王閃擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(86, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿貝爾" Then
   Select Case atkingck(86, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 1) >= 3 And atkingck(86, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
                   atkingck(86, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 1) < 3 And atkingck(86, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(86, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(86, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 10290
                   atkingno(i, 6) = 8490
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 阿貝爾_閃電旋風刺()
If FormMainMode.personatk(2).Caption = "閃電旋風刺" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(87, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿貝爾" Then
   Select Case atkingck(87, 1)
      Case 1
           If movecp = 2 Then
                If atkingpagetot(1, 3) >= 1 And atkingck(87, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
                   atkingck(87, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 3) < 1 And atkingck(87, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(87, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6555
                   atkingno(i, 6) = 8625
                   atkingno(i, 7) = 87
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(87, 2) = 0
             If movecp > 1 Then
                 戰鬥系統類.執行動作_距離變更 movecp - 1
             End If
   End Select
End If
End Sub
Sub 阿貝爾_幻影劍舞()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "幻影劍舞" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(88, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿貝爾" Then
   Select Case atkingck(88, 1)
      Case 1
            If movecp = 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                        If pagecardnum(i, 1) = a1a And pagecardnum(i, 2) = 1 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a1a And pagecardnum(i, 2) = 2 Then
                           rrr(2) = rrr(2) + 1
                        End If
                        If pagecardnum(i, 1) = a1a And pagecardnum(i, 2) = 3 Then
                           rrr(3) = rrr(3) + 1
                        End If
                    End If
                 Next
            End If
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingck(88, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingck(88, 2) = 0 Then
                戰鬥系統類.人物技能欄燈開關 True, 3
                atkingck(88, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 9
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingck(88, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingck(88, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 3
                atkingck(88, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 9
              End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             atkingck(88, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking3-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8520
                   atkingno(i, 6) = 8280
                   atkingno(i, 7) = 88
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\阿貝爾\atking3-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 阿貝爾_抽刀斷水計()
If FormMainMode.personatk(4).Caption = "抽刀斷水計" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(89, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿貝爾" Then
   Select Case atkingck(89, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 4) >= 3 And atkingck(89, 2) = 0 Then
                   atkingck(89, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 3 And atkingck(89, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(89, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 8745
                   atkingno(i, 6) = 10200
                   atkingno(i, 7) = 89
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(89, 2) = 0
             '======================
               Do
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 3) = 23 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                          FormMainMode.personcomspe(i).person_turn = 1
                          人物異常狀態資料庫(2, i, 2) = 1
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, i, 23, app_path & "gif\異常狀態\atkingerr.gif", 0, 1
                          異常狀態檢查數(23, 1) = 1
                          異常狀態檢查數(23, 2) = 1
                          Exit Do
                       End If
                    Next
               Loop
   End Select
End If
End Sub
Sub 利恩_劫影攻擊()
If FormMainMode.personatk(1).Caption = "劫影攻擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(90, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "利恩" Then
   Select Case atkingck(90, 1)
      Case 1
            If atkingpagetot(1, 4) >= 1 And atkingck(90, 2) = 0 Then
               atkingck(90, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 1 And atkingck(90, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(90, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9705
                   atkingno(i, 6) = 9090
                   atkingno(i, 7) = 90
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(90, 2) = 0
             '======================
             If 擲骰後骰傷害數 > 0 Then
                    Do
                         For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                           If 人物異常狀態資料庫(2, i, 3) = 17 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                               FormMainMode.personcomspe(i).person_turn = 2
                               人物異常狀態資料庫(2, i, 2) = 2
                               Exit Do
                           End If
                         Next
                         For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                            If 人物異常狀態資料庫(2, i, 2) = 0 Then
                               戰鬥系統類.人物異常狀態表設定_初設 2, i, 17, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                               異常狀態檢查數(17, 1) = 1
                               異常狀態檢查數(17, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
              End If
   End Select
End If
End Sub
Sub 利恩_毒牙()
If FormMainMode.personatk(2).Caption = "毒牙" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(91, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "利恩" Then
   Select Case atkingck(91, 1)
      Case 1
            If movecp = 1 Then
                If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 4) >= 3 And atkingck(91, 2) = 0 Then
                   atkingck(91, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
                End If
                If (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 4) < 3) And atkingck(91, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(91, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
                 End If
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6090
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 91
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(91, 2) = 0
             '======================
             If 擲骰後骰傷害數 > 0 Then
                    Do
                         For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                           If 人物異常狀態資料庫(2, i, 3) = 21 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                               FormMainMode.personcomspe(i).person_turn = 3
                               人物異常狀態資料庫(2, i, 2) = 3
                               Exit Do
                           End If
                         Next
                         For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                            If 人物異常狀態資料庫(2, i, 2) = 0 Then
                               戰鬥系統類.人物異常狀態表設定_初設 2, i, 21, app_path & "gif\異常狀態\damage.gif", 0, 3
                               異常狀態檢查數(21, 1) = 1
                               異常狀態檢查數(21, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
             End If
   End Select
End If
End Sub
Sub 利恩_反擊的狼煙()
If FormMainMode.personatk(3).Caption = "反擊的狼煙" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(92, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "利恩" Then
   Select Case atkingck(92, 1)
        Case 1
            If atkingpagetot(1, 4) >= 1 And atkingck(92, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(92, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 4) < 1 And atkingck(92, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(92, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -840
                   atkingno(i, 5) = 8250
                   atkingno(i, 6) = 10155
                   atkingno(i, 7) = 92
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If 擲骰後骰傷害數 > 0 And liveus(角色人物對戰人數(1, 2)) > 0 Then
                atking_利恩_反擊的狼煙紀錄數(1) = 擲骰後骰傷害數 + 1
                If Val(FormMainMode.pageul.Caption) < atking_利恩_反擊的狼煙紀錄數(1) And atking_利恩_反擊的狼煙紀錄數(2) = 0 Then
                   戰鬥系統類.執行動作_洗牌
                End If
                atking_利恩_反擊的狼煙紀錄數(2) = atking_利恩_反擊的狼煙紀錄數(2) + 1
                If Val(FormMainMode.pageul.Caption) > 0 Then
                    Do Until atking_利恩_反擊的狼煙紀錄數(2) > atking_利恩_反擊的狼煙紀錄數(1)
                        目前數(15) = 25
                        FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                        Exit Sub
                    Loop
                End If
            End If
            If atking_利恩_反擊的狼煙紀錄數(2) > atking_利恩_反擊的狼煙紀錄數(1) Or 擲骰後骰傷害數 <= 0 _
                Or Val(FormMainMode.pageul.Caption) <= 0 Or liveus(角色人物對戰人數(1, 2)) <= 0 Then
                目前數(24) = 22
                atkingck(92, 1) = 4
                FormMainMode.等待時間_2.Enabled = True
            End If
        Case 4
            atkingck(92, 2) = 0
            Erase atking_利恩_反擊的狼煙紀錄數
   End Select
End If
End Sub
Sub 利恩_背刺()
If FormMainMode.personatk(4).Caption = "背刺" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(93, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "利恩" Then
   Select Case atkingck(93, 1)
      Case 1
            If movecp = 3 Then
                If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 3 And atkingck(93, 2) = 0 Then
                   If 執行動作_檢查是否有指定異常狀態(2, 17) = True Then
                        atkingck(93, 2) = 1
                        戰鬥系統類.人物技能欄燈開關 True, 4
                        atkingtrn(1) = Val(atkingtrn(1)) + 1
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 12
                    End If
                End If
                If (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 3) And atkingck(93, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(93, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 12
                 End If
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6795
                   atkingno(i, 6) = 9405
                   atkingno(i, 7) = 93
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=============================
             atkingck(93, 2) = 0
             If 執行動作_檢查是否有指定異常狀態(2, 17) = False Then
                 直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) - 12
                 攻擊防禦骰子總數(1) = Val(FormMainMode.顯示列1.goi1)
             End If
   End Select
End If
End Sub
Sub 夏洛特_冬之夢()
If FormMainMode.personatk(2).Caption = "冬之夢" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(95, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "夏洛特" Then
   Select Case atkingck(95, 1)
      Case 1
            If movecp < 3 Then
                If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 3 And atkingck(95, 2) = 0 Then
                   atkingck(95, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                End If
                If (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 3) And atkingck(95, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(95, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                 End If
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9375
                   atkingno(i, 7) = 95
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(95, 2) = 0
             '========================
             For i = 18 To (turn + 3) Step -1
                  pageeventnum(1, i, 1) = pageeventnum(1, i - 2, 1)
                  pageeventnum(1, i, 2) = pageeventnum(1, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 2)
                  pageeventnum(1, i, 1) = "劍5/槍5"
                  pageeventnum(1, i, 2) = 一般系統類.事件卡資料庫("劍5/槍5", 2)
             Next
   End Select
End If
End Sub
Sub 夏洛特_大聖堂()
Dim p, i, j As Integer
If FormMainMode.personatk(1).Caption = "大聖堂" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(94, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "夏洛特" Then
   Select Case atkingck(94, 1)
      Case 1
            If atkingpagetot(1, 4) >= 2 And atkingck(94, 2) = 0 Then
               atkingck(94, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 2 And atkingck(94, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(94, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6705
                   atkingno(i, 6) = 10185
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingck(94, 1) = 3
        Case 3
             atking_夏洛特_大聖堂骰量紀錄數(1) = 擲骰後骰傷害數
             擲骰表單溝通暫時變數(2) = 0
             擲骰表單溝通暫時變數(3) = 0
             '========================================
                For p = 1 To Val(FormMainMode.顯示列1.goi1)
                   Randomize Timer
                   i = Int(Rnd() * 6) + 1
                   If i = 1 Or i = 6 Then 擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) + 1
                Next
                For p = 1 To Val(FormMainMode.顯示列1.goi2)
                   Randomize Timer
                   j = Int(Rnd() * 6) + 1
                   If j = 1 Or j = 6 Then 擲骰表單溝通暫時變數(3) = Val(擲骰表單溝通暫時變數(3)) + 1
                Next
                '=============================
                技能動畫顯示階段數 = 1
                atkingck(94, 1) = 4
                FormMainMode.骰子執行完啟動.Enabled = False
                目前數(22) = 12
                FormMainMode.等待時間.Enabled = True
          Case 4
                atking_夏洛特_大聖堂骰量紀錄數(2) = 擲骰後骰傷害數
                '==========================
                If atking_夏洛特_大聖堂骰量紀錄數(1) > atking_夏洛特_大聖堂骰量紀錄數(2) Then
                    擲骰表單溝通暫時變數(2) = atking_夏洛特_大聖堂骰量紀錄數(2)
                Else
                    擲骰表單溝通暫時變數(2) = atking_夏洛特_大聖堂骰量紀錄數(1)
                End If
                擲骰後骰傷害數 = Val(擲骰表單溝通暫時變數(2))
                atkingck(94, 2) = 0
                Erase atking_夏洛特_大聖堂骰量紀錄數
   End Select
End If
End Sub
Sub 夏洛特_夜未央()
If FormMainMode.personatk(3).Caption = "夜未央" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(96, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "夏洛特" Then
   Select Case atkingck(96, 1)
      Case 1
            If movecp < 3 Then
                If atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 And atkingck(96, 2) = 0 Then
                   atkingck(96, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1) And atkingck(96, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(96, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7575
                   atkingno(i, 6) = 9660
                   atkingno(i, 7) = 96
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(96, 2) = 0
             '========================
             戰鬥系統類.回復執行_使用者 1, 1
             '========================
             For i = 18 To (turn + 3) Step -1
                  pageeventnum(1, i, 1) = pageeventnum(1, i - 2, 1)
                  pageeventnum(1, i, 2) = pageeventnum(1, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 2)
                  pageeventnum(1, i, 1) = "HP回復3"
                  pageeventnum(1, i, 2) = 一般系統類.事件卡資料庫("HP回復3", 2)
             Next
   End Select
End If
End Sub
Sub 夏洛特_幸福的理由()
If FormMainMode.personatk(4).Caption = "幸福的理由" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(97, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "夏洛特" Then
   Select Case atkingck(97, 1)
      Case 1
            If atkingpagetot(1, 4) >= 3 And atkingck(97, 2) = 0 Then
               atkingck(97, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 3 And atkingck(97, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(97, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 600
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6915
                   atkingno(i, 6) = 9690
                   atkingno(i, 7) = 97
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(97, 2) = 0
             '========================
             If 牌總階段數(1) > 0 Then
                 牌總階段數(1) = 牌總階段數(1) - 1
             End If
             '========================
             For i = 18 To (turn + 4) Step -1
                  pageeventnum(1, i, 1) = pageeventnum(1, i - 2, 1)
                  pageeventnum(1, i, 2) = pageeventnum(1, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 3)
                  pageeventnum(1, i, 1) = "機會5"
                  pageeventnum(1, i, 2) = 一般系統類.事件卡資料庫("機會5", 2)
             Next
   End Select
End If
End Sub
Sub 泰瑞爾_Rud_913()
If FormMainMode.personatk(1).Caption = "Rud-913" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(116, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "泰瑞爾" Then
   Select Case atkingck(116, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 3) >= 1 And atkingck(116, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(116, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 3) < 1) And atkingck(116, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(116, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6675
                   atkingno(i, 6) = 9105
                   atkingno(i, 7) = 116
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(116, 2) = 0
            '================
            戰鬥系統類.執行動作_距離變更 3
   End Select
End If
End Sub
Sub 泰瑞爾_Von_541()
If FormMainMode.personatk(2).Caption = "Von-541" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(117, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "泰瑞爾" Then
   Select Case atkingck(117, 1)
      Case 1
            If atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 2) >= 1 And atkingck(117, 2) = 0 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
               atkingck(117, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If (atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 2) < 1) And atkingck(117, 2) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(117, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7140
                   atkingno(i, 6) = 9645
                   atkingno(i, 7) = 117
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(117, 2) = 0
            '================
            If 擲骰後骰傷害數 >= 10 Then
                戰鬥系統類.傷害執行_技能直傷_電腦 擲骰後骰傷害數, 1
                擲骰後骰傷害數 = 0
                擲骰表單溝通暫時變數(2) = 0
            End If
   End Select
End If
End Sub
Sub 泰瑞爾_Chr_799()
Dim m As Integer
If FormMainMode.personatk(3).Caption = "Chr-799" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(118, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "泰瑞爾" Then
   Select Case atkingck(118, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 2 And atkingpagetot(1, 4) >= 2 And atkingck(118, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(118, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 5) < 2 Or atkingpagetot(1, 4) < 2) And atkingck(118, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(118, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 120
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6750
                   atkingno(i, 6) = 9255
                   atkingno(i, 7) = 118
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(118, 2) = 0
            '================
            m = Int(Rnd() * 3) + 1
            Select Case m
                Case 1
                       Do
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                                  If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 3
                                      FormMainMode.personusspe(j).person_turn = 5
                                      人物異常狀態資料庫(1, j, 1) = 3
                                      人物異常狀態資料庫(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 3, 5
                                  異常狀態檢查數(7, 1) = 1
                                  異常狀態檢查數(7, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                                If 人物異常狀態資料庫(2, j, 3) = 4 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 3
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 人物異常狀態資料庫(2, j, 1) = 3
                                 人物異常狀態資料庫(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 4, app_path & "gif\異常狀態\atkdown.gif", 3, 5
                                 異常狀態檢查數(4, 1) = 1
                                 異常狀態檢查數(4, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
                Case 2
                        Do
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                                  If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 3
                                      FormMainMode.personusspe(j).person_turn = 5
                                      人物異常狀態資料庫(1, j, 1) = 3
                                      人物異常狀態資料庫(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 8, app_path & "gif\異常狀態\defup.gif", 3, 5
                                  異常狀態檢查數(8, 1) = 1
                                  異常狀態檢查數(8, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                                If 人物異常狀態資料庫(2, j, 3) = 5 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 3
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 人物異常狀態資料庫(2, j, 1) = 3
                                 人物異常狀態資料庫(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 5, app_path & "gif\異常狀態\defdown.gif", 3, 5
                                 異常狀態檢查數(5, 1) = 1
                                 異常狀態檢查數(5, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
                Case 3
                        Do
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                                  If 人物異常狀態資料庫(1, j, 3) = 9 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 1
                                      FormMainMode.personusspe(j).person_turn = 5
                                      人物異常狀態資料庫(1, j, 1) = 1
                                      人物異常狀態資料庫(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 9, app_path & "gif\異常狀態\movup.gif", 1, 5
                                  異常狀態檢查數(9, 1) = 1
                                  異常狀態檢查數(9, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                                If 人物異常狀態資料庫(2, j, 3) = 6 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 1
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 人物異常狀態資料庫(2, j, 1) = 1
                                 人物異常狀態資料庫(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 6, app_path & "gif\異常狀態\movdown.gif", 1, 5
                                 異常狀態檢查數(6, 1) = 1
                                 異常狀態檢查數(6, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
            End Select
   End Select
End If
End Sub
Sub 泰瑞爾_Wil_846()
If FormMainMode.personatk(4).Caption = "Wil-846" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(119, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "泰瑞爾" Then
   Select Case atkingck(119, 1)
      Case 1
           If movecp = 3 Then
                If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 3 And atkingck(119, 2) = 0 Then
                   atkingck(119, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 3) And atkingck(119, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(119, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6720
                   atkingno(i, 6) = 10320
                   atkingno(i, 7) = 119
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(119, 2) = 0
            '================
            戰鬥系統類.傷害執行_技能直傷_電腦 2, 1
            '================
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          人物異常狀態資料庫(1, j, 1) = 9
                      End If
                      If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          人物異常狀態資料庫(1, j, 1) = 9
                      End If
                      If 人物異常狀態資料庫(1, j, 3) = 9 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          人物異常狀態資料庫(1, j, 1) = 9
                      End If
                      If 人物異常狀態資料庫(1, j, 3) = 10 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          人物異常狀態資料庫(1, j, 1) = 9
                      End If
                      If 人物異常狀態資料庫(1, j, 3) = 11 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          人物異常狀態資料庫(1, j, 1) = 9
                      End If
                      If 人物異常狀態資料庫(1, j, 3) = 12 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          人物異常狀態資料庫(1, j, 1) = 9
                      End If
                 Next
                 '==========================
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                    If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        人物異常狀態資料庫(2, j, 1) = 9
                    End If
                    If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        人物異常狀態資料庫(2, j, 1) = 9
                    End If
                    If 人物異常狀態資料庫(2, j, 3) = 3 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        人物異常狀態資料庫(2, j, 1) = 9
                    End If
                    If 人物異常狀態資料庫(2, j, 3) = 4 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        人物異常狀態資料庫(2, j, 1) = 9
                    End If
                    If 人物異常狀態資料庫(2, j, 3) = 5 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        人物異常狀態資料庫(2, j, 1) = 9
                    End If
                    If 人物異常狀態資料庫(2, j, 3) = 6 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        人物異常狀態資料庫(2, j, 1) = 9
                    End If
                Next
   End Select
End If
End Sub
Sub 瑪格莉特_月光()
Dim m As Integer
If FormMainMode.personatk(1).Caption = "月光" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(122, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "瑪格莉特" Then
   Select Case atkingck(122, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(1, 4) >= 1 And atkingck(122, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
                   atkingck(122, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 1 And atkingck(122, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(122, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -240
                   atkingno(i, 5) = 6195
                   atkingno(i, 6) = 10350
                   atkingno(i, 7) = 122
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
             Erase atking_瑪格莉特_月光紀錄數
             '========================
             For i = 1 To 106
                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                    If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                         atking_瑪格莉特_月光紀錄數(i) = 1
                         atking_瑪格莉特_月光紀錄數(107) = atking_瑪格莉特_月光紀錄數(107) + 1
                     End If
                End If
            Next
            If atking_瑪格莉特_月光紀錄數(107) > 2 Then
                atking_瑪格莉特_月光紀錄數(107) = 2
            End If
            '=========================
            If atking_瑪格莉特_月光紀錄數(107) > 0 Then
                Do
                    m = Int(Rnd() * 106) + 1
                    If atking_瑪格莉特_月光紀錄數(m) = 1 Then
                        目前數(16) = m
                        atking_瑪格莉特_月光紀錄數(m) = 0
                        atking_瑪格莉特_月光紀錄數(0) = atking_瑪格莉特_月光紀錄數(0) + 1
                        FormMainMode.tr電腦牌_翻牌.Enabled = True
                        Exit Sub
                    End If
                Loop
            Else
               目前數(22) = 23
               FormMainMode.等待時間.Enabled = True
            End If
        Case 4
            FormMainMode.tr電腦牌_棄牌.Enabled = True
            目前數(17) = 8
        Case 5
            If atking_瑪格莉特_月光紀錄數(107) > 1 And atking_瑪格莉特_月光紀錄數(0) < atking_瑪格莉特_月光紀錄數(107) Then
                Do
                    m = Int(Rnd() * 106) + 1
                    If atking_瑪格莉特_月光紀錄數(m) = 1 Then
                        目前數(16) = m
                        atking_瑪格莉特_月光紀錄數(m) = 0
                        atking_瑪格莉特_月光紀錄數(0) = atking_瑪格莉特_月光紀錄數(0) + 1
                        FormMainMode.tr電腦牌_翻牌.Enabled = True
                        Exit Sub
                    End If
                Loop
            ElseIf atking_瑪格莉特_月光紀錄數(0) >= 2 Then
               目前數(24) = 24
               FormMainMode.等待時間_2.Enabled = True
            Else
               目前數(24) = 23
               FormMainMode.等待時間_2.Enabled = True
            End If
        Case 6
            If atking_瑪格莉特_月光紀錄數(107) = 0 Then
                atking_瑪格莉特_月光紀錄數(107) = 99
               目前數(22) = 23
               FormMainMode.等待時間.Enabled = True
            ElseIf atking_瑪格莉特_月光紀錄數(107) > 0 And atking_瑪格莉特_月光紀錄數(0) = 0 Then
               atkingck(122, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            ElseIf atking_瑪格莉特_月光紀錄數(107) > 0 And atking_瑪格莉特_月光紀錄數(0) = 1 Then
               戰鬥系統類.傷害執行_技能直傷_電腦 atking_瑪格莉特_月光紀錄數(0), 1
               目前數(24) = 24
               FormMainMode.等待時間_2.Enabled = True
            End If
        Case 7
            If atking_瑪格莉特_月光紀錄數(107) > 0 And atking_瑪格莉特_月光紀錄數(0) = 1 Then
               atkingck(122, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            ElseIf atking_瑪格莉特_月光紀錄數(0) >= 2 Then
               戰鬥系統類.傷害執行_技能直傷_電腦 atking_瑪格莉特_月光紀錄數(0), 1
               atkingck(122, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            End If
   End Select
End If
End Sub
Sub 瑪格莉特_恍惚()
If FormMainMode.personatk(2).Caption = "恍惚" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(123, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "瑪格莉特" Then
   Select Case atkingck(123, 1)
        Case 1
            If movecp = 1 Then
             If atkingpagetot(1, 2) >= 3 And atkingpagetot(1, 3) >= 1 And atkingck(123, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(123, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(123, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
            ElseIf (atkingpagetot(1, 2) < 3 Or atkingpagetot(1, 3) < 1) And atkingck(123, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(123, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(123, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5580
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 123
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(123, 2) = 0
            '===============
            If 擲骰後骰傷害數 <= 0 Then
                Do
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 3) = 17 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                          FormMainMode.personcomspe(i).person_turn = 2
                          人物異常狀態資料庫(2, i, 2) = 2
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, i, 17, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                          異常狀態檢查數(17, 1) = 1
                          異常狀態檢查數(17, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub 瑪格莉特_末日幻影()
Dim m As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "末日幻影" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(124, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "瑪格莉特" Then
   Select Case atkingck(124, 1)
        Case 1
            If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 3) = 0 And atkingck(124, 2) = 0 Then
'            If atkingpagetot(1, 3) >= 1 And atkingck(124, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(124, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 5) < 1 Or atkingpagetot(1, 3) > 0) And atkingck(124, 2) = 1 Then
'            ElseIf atkingpagetot(1, 3) < 1 And atkingck(124, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(124, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 10455
                   atkingno(i, 7) = 124
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
                atkingck(124, 2) = 0
                Select Case movecp
                    Case 1
                       Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 30 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 3
                                  人物異常狀態資料庫(2, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 30, app_path & "gif\異常狀態\恐怖.gif", 0, 3
                                  異常狀態檢查數(30, 1) = 1
                                  異常狀態檢查數(30, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 21 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 3
                                  人物異常狀態資料庫(2, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 21, app_path & "gif\異常狀態\damage.gif", 0, 3
                                  異常狀態檢查數(21, 1) = 1
                                  異常狀態檢查數(21, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 28 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 3
                                  人物異常狀態資料庫(2, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 28, app_path & "gif\異常狀態\狂戰士.gif", 0, 3
                                  異常狀態檢查數(28, 1) = 1
                                  異常狀態檢查數(28, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
   End Select
End If
End Sub
Sub 瑪格莉特_地獄獵心獸()
Dim m As Integer
If FormMainMode.personatk(4).Caption = "地獄獵心獸" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(125, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "瑪格莉特" Then
   Select Case atkingck(125, 1)
        Case 1
             If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 3 And atkingck(125, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(125, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(125, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 3) And atkingck(125, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(125, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(125, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6630
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 125
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(125, 2) = 0
            '===============
            m = (atkingpagetot(1, 1) + atkingpagetot(1, 5)) \ 5
            戰鬥系統類.傷害執行_技能直傷_電腦 m, 1
   End Select
End If
End Sub
Sub 庫勒尼西_沙漠中的海市蜃樓()
If FormMainMode.personatk(1).Caption = "沙漠中的海市蜃樓" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(128, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "庫勒尼西" Then
   Select Case atkingck(128, 1)
      Case 1
           If movecp = 3 Then
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 2
                atkingck(128, 2) = 1
                戰鬥系統類.人物技能欄燈開關 True, 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(128, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 960
                   atkingno(i, 4) = 1560
                   atkingno(i, 5) = 6270
                   atkingno(i, 6) = 9645
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 庫勒尼西_瘋狂眼窩()
If FormMainMode.personatk(2).Caption = "瘋狂眼窩" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(129, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "庫勒尼西" Then
   Select Case atkingck(129, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 4) >= 1 And atkingck(129, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 2
                   atkingck(129, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 1 And atkingck(129, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 2
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(129, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
            End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -720
                   atkingno(i, 5) = 8505
                   atkingno(i, 6) = 10140
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(FormMainMode.pagecomglead.Caption) > 0 Then
                 atking_庫勒尼西_瘋狂眼窩紀錄數 = 1
                 '==========================
                  Do Until atking_庫勒尼西_瘋狂眼窩紀錄數 > 3 Or Val(FormMainMode.pagecomglead.Caption) <= 0
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                            目前數(17) = 9
                            目前數(16) = m
                            atking_庫勒尼西_瘋狂眼窩紀錄數 = atking_庫勒尼西_瘋狂眼窩紀錄數 + 1
                            FormMainMode.tr電腦牌_翻牌.Enabled = True
                            Exit Sub
                        End If
                   Loop
             Else
                 atkingck(129, 1) = 5
                 FormMainMode.骰子執行完啟動.Enabled = True
             End If
        Case 4
             Do Until atking_庫勒尼西_瘋狂眼窩紀錄數 > 3 Or Val(FormMainMode.pagecomglead.Caption) <= 0
                 Randomize
                 m = Int(Rnd() * 106) + 1
                 If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                     目前數(17) = 9
                     目前數(16) = m
                     atking_庫勒尼西_瘋狂眼窩紀錄數 = atking_庫勒尼西_瘋狂眼窩紀錄數 + 1
                     FormMainMode.tr電腦牌_翻牌.Enabled = True
                     Exit Sub
                 End If
            Loop
            If atking_庫勒尼西_瘋狂眼窩紀錄數 > 3 Or Val(FormMainMode.pagecomglead.Caption) <= 0 Then
                atkingck(129, 1) = 5
                目前數(24) = 22
                FormMainMode.等待時間_2.Enabled = True
            End If
        Case 5
            atkingck(129, 2) = 0
        Case 6
            FormMainMode.tr電腦牌_棄牌.Enabled = True
   End Select
End If
End Sub
Sub 庫勒尼西_深淵()
Dim m As Integer
If FormMainMode.personatk(3).Caption = "深淵" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(130, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "庫勒尼西" Then
   Select Case atkingck(130, 1)
        Case 1
             If atkingpagetot(1, 4) >= 3 And atkingck(130, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(130, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingck(130, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(1, 4) < 3 And atkingck(130, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(130, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(130, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8970
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 130
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(130, 2) = 0
            '===============
            m = Int(atkingpagetot(1, 4) / 2 + 0.9)
            戰鬥系統類.傷害執行_技能直傷_電腦 m, 1
   End Select
End If
End Sub
Sub 庫勒尼西_黑暗漩渦()
Dim m As Integer
If FormMainMode.personatk(4).Caption = "黑暗漩渦" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(131, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "庫勒尼西" Then
   Select Case atkingck(131, 1)
        Case 1
             If atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 And atkingck(131, 2) = 0 Then
'             If atkingpagetot(1, 2) >= 1 And atkingck(131, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingck(131, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
            ElseIf (atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1) And atkingck(131, 2) = 1 Then
'            ElseIf atkingpagetot(1, 2) < 1 And atkingck(131, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(131, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6480
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 131
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(131, 2) = 0
            '===============
            m = movecp + 1
            If m > 3 Then m = 3
            戰鬥系統類.執行動作_距離變更 m
   End Select
End If
End Sub
Sub 蕾格烈芙_CTL()
Dim i As Integer
If FormMainMode.personatk(1).Caption = "C.T.L" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(135, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾格烈芙" Then
   Select Case atkingck(135, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 4 And atkingpagetot(1, 4) >= 1 And atkingck(135, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(135, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                   For i = 1 To 106
                       If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                          If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                              攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                              目前數(27) = 1
                              Exit For
                          End If
                       End If
                   Next
                End If
                If (atkingpagetot(1, 5) < 4 Or atkingpagetot(1, 4) < 1) And atkingck(135, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   If 目前數(27) = 1 Then
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                       目前數(27) = 0
                   End If
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(135, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             atkingck(135, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking1-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6540
                   atkingno(i, 6) = 9990
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\蕾格烈芙\atking1-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             If 目前數(27) = 1 Then
                 For i = 1 To 106
                       If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                          If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                              Exit For
                          End If
                       End If
                  Next
                  If i = 107 Then
                      攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                      直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) - 6
                      For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, i, 3) = 31 Then
                              攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                              直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) - 6
                          End If
                      Next
                  End If
                  目前數(27) = 0
             End If
   End Select
End If
End Sub
Sub 蕾格烈芙_BPA()
If FormMainMode.personatk(2).Caption = "B.P.A" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(136, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾格烈芙" Then
   Select Case atkingck(136, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(1, 4) >= 3 And atkingck(136, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
                   atkingck(136, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 3 And atkingck(136, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(136, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6030
                   atkingno(i, 6) = 10530
                   atkingno(i, 7) = 136
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             戰鬥系統類.傷害執行_技能直傷_電腦 pageqlead(2), 1
             atkingck(136, 2) = 0
   End Select
End If
End Sub
Sub 蕾格烈芙_LAR()
If FormMainMode.personatk(3).Caption = "L.A.R" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(137, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾格烈芙" Then
   Select Case atkingck(137, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 2) >= 2 And atkingck(137, 2) = 0 Then
                   atkingck(137, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 2) < 2 And atkingck(137, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(137, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5400
                   atkingno(i, 6) = 9015
                   atkingno(i, 7) = 137
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             戰鬥系統類.回復執行_使用者 1, 1
        Case 4
             atkingck(137, 2) = 0
             If 擲骰後骰傷害數 > 0 Then
                 戰鬥系統類.回復執行_使用者 1, 1
             End If
   End Select
End If
End Sub
Sub 蕾格烈芙_SSS()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "S.S.S" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(138, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "蕾格烈芙" Then
   Select Case atkingck(138, 1)
        Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 2) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 2) = 2 Then
                       rrr(2) = rrr(2) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 2) = 3 Then
                       rrr(3) = rrr(3) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingck(138, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingck(138, 2) = 0 Then
                戰鬥系統類.人物技能欄燈開關 True, 4
                atkingck(138, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingck(138, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingck(138, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 4
                atkingck(138, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
              End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6135
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 138
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
                atkingck(138, 2) = 0
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 31 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 31, app_path & "gif\異常狀態\混沌.gif", 0, 3
                          異常狀態檢查數(31, 1) = 1
                          異常狀態檢查數(31, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
   End Select
End If
End Sub
Sub 多妮妲_殘虐傾向()
If FormMainMode.personatk(1).Caption = "殘虐傾向" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(140, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "多妮妲" Then
   Select Case atkingck(140, 1)
      Case 1
            If atkingpagetot(1, 4) >= 2 And atkingck(140, 2) = 0 Then
               atkingck(140, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 2 And atkingck(140, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(140, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\多妮妲\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8160
                   atkingno(i, 6) = 9120
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(140, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                       Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 21 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 2
                                  人物異常狀態資料庫(2, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 21, app_path & "gif\異常狀態\damage.gif", 0, 2
                                  異常狀態檢查數(21, 1) = 1
                                  異常狀態檢查數(21, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 17 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 2
                                  人物異常狀態資料庫(2, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 17, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                                  異常狀態檢查數(17, 1) = 1
                                  異常狀態檢查數(17, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, i, 3) = 23 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                  FormMainMode.personcomspe(i).person_turn = 2
                                  人物異常狀態資料庫(2, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                               If 人物異常狀態資料庫(2, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, i, 23, app_path & "gif\異常狀態\atkingerr.gif", 0, 2
                                  異常狀態檢查數(23, 1) = 1
                                  異常狀態檢查數(23, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
            End If
   End Select
End If
End Sub
Sub 多妮妲_異質者()
Dim i, j, rrr As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "異質者" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(141, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "多妮妲" Then
    Select Case atkingck(141, 1)
         Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
'                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) >= 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                   rrr = rrr + 1
                End If
             Next
             If rrr >= 1 And atkingck(141, 2) = 0 Then
                atkingck(141, 2) = 1
                戰鬥系統類.人物技能欄燈開關 True, 2
                atkingtrn(1) = Val(atkingtrn(1)) + 1
             End If
             If rrr < 1 And atkingck(141, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 2
                atkingck(141, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
         Case 2
           For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\多妮妲\atking2_1.jpg"
                atkingno(i, 2) = 1
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6465
                atkingno(i, 6) = 9765
                atkingno(i, 7) = 141
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
           Next
           戰鬥系統類.人物技能欄燈開關 False, 2
           戰鬥系統類.自動捲軸捲動
     Case 3
          atkingck(141, 2) = 0
          If Val(擲骰表單溝通暫時變數(3)) - Val(擲骰表單溝通暫時變數(2)) >= liveus(角色人物對戰人數(1, 2)) And 異常狀態檢查數(14, 2) = 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 6
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 1) = 6
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 6, 3
                          異常狀態檢查數(7, 1) = 1
                          異常狀態檢查數(7, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '==================================
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 14 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 0
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 1) = 0
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 14, app_path & "gif\異常狀態\不死.gif", 0, 3
                          異常狀態檢查數(14, 1) = 1
                          異常狀態檢查數(14, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '===============================
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 15 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 0
                              FormMainMode.personusspe(j).person_turn = 3
                              人物異常狀態資料庫(1, j, 1) = 0
                              人物異常狀態資料庫(1, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 15, app_path & "gif\異常狀態\自壞.gif", 0, 3
                          異常狀態檢查數(15, 1) = 1
                          異常狀態檢查數(15, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
         End If
    End Select
End If
End Sub
Sub 多妮妲_超級女主角()
If FormMainMode.personatk(3).Caption = "超級女主角" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(142, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "多妮妲" Then
   Select Case atkingck(142, 1)
      Case 1
            If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 5) >= 3 And atkingpagetot(1, 4) >= 2 And atkingck(142, 2) = 0 Then
               atkingck(142, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 5) < 3 Or atkingpagetot(1, 4) < 2) And atkingck(142, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(142, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\多妮妲\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 5970
                   atkingno(i, 6) = 10365
                   atkingno(i, 7) = 142
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(142, 2) = 0
            '==================
            Do
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 6
                          FormMainMode.personusspe(j).person_turn = 5
                          人物異常狀態資料庫(1, j, 1) = 6
                          人物異常狀態資料庫(1, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                   If 人物異常狀態資料庫(1, j, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 6, 5
                      異常狀態檢查數(7, 1) = 1
                      異常狀態檢查數(7, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
            Do
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 4
                          FormMainMode.personusspe(j).person_turn = 5
                          人物異常狀態資料庫(1, j, 1) = 4
                          人物異常狀態資料庫(1, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                   If 人物異常狀態資料庫(1, j, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 1, j, 8, app_path & "gif\異常狀態\defup.gif", 4, 5
                      異常狀態檢查數(8, 1) = 1
                      異常狀態檢查數(8, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
            Do
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, j, 3) = 9 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 1
                          FormMainMode.personusspe(j).person_turn = 5
                          人物異常狀態資料庫(1, j, 1) = 1
                          人物異常狀態資料庫(1, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                   If 人物異常狀態資料庫(1, j, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 1, j, 9, app_path & "gif\異常狀態\movup.gif", 1, 5
                      異常狀態檢查數(9, 1) = 1
                      異常狀態檢查數(9, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
   End Select
End If
End Sub
Sub 多妮妲_律死擊()
If FormMainMode.personatk(4).Caption = "律死擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(143, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "多妮妲" Then
   Select Case atkingck(143, 1)
      Case 1
            If movecp = 1 Then
                    If atkingpagetot(1, 1) >= 4 And atkingpagetot(1, 4) >= 2 And atkingck(143, 2) = 0 Then
                       atkingck(143, 2) = 1
                       戰鬥系統類.人物技能欄燈開關 True, 4
                       atkingtrn(1) = Val(atkingtrn(1)) + 1
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 8
                    End If
                    If (atkingpagetot(1, 1) < 4 Or atkingpagetot(1, 4) < 2) And atkingck(143, 2) = 1 Then
                       戰鬥系統類.人物技能欄燈開關 False, 4
                       atkingck(143, 2) = 0
                       atkingtrn(1) = Val(atkingtrn(1)) - 1
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 8
                     End If
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\多妮妲\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6450
                   atkingno(i, 6) = 10200
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(143, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 19 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 5
                              人物異常狀態資料庫(2, j, 1) = 0
                              人物異常狀態資料庫(2, j, 2) = 5
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 19, app_path & "gif\異常狀態\自壞.gif", 0, 5
                          異常狀態檢查數(19, 1) = 1
                          異常狀態檢查數(19, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub 傑多_因果之線()
Dim m As Integer
If FormMainMode.personatk(1).Caption = "因果之線" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(144, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "傑多" Then
   Select Case atkingck(144, 1)
      Case 1
            If atkingpagetot(1, 4) >= 1 And atkingck(144, 2) = 0 Then
               atkingck(144, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 1 And atkingck(144, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(144, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9255
                   atkingno(i, 7) = 144
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             Do
                Randomize
                m = Int(Rnd() * 106) + 1
                If Val(pagecardnum(m, 6)) = 1 And Val(pagecardnum(m, 5)) = 2 Then
                    目前數(16) = m
                     FormMainMode.tr電腦牌_翻牌.Enabled = True
                     Exit Do
                End If
            Loop
        Case 4
             目前數(17) = 2
             FormMainMode.tr電腦牌_偷牌.Enabled = True
             atkingck(144, 2) = 0
   End Select
End If
End Sub
Sub 傑多_因果之輪()
Dim m, n As Integer
Dim aw(1 To 2) As Integer
If FormMainMode.personatk(2).Caption = "因果之輪" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(145, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "傑多" Then
   Select Case atkingck(145, 1)
      Case 1
            If atkingpagetot(1, 4) >= 2 And atkingck(145, 2) = 0 Then
               atkingck(145, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 2 And atkingck(145, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(145, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7260
                   atkingno(i, 6) = 8925
                   atkingno(i, 7) = 145
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             技能動畫顯示階段數 = 10
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.防禦階段_階段初始.Enabled = False
        Case 3
                階段狀態數 = 3
                For m = 1 To 106
                    If Val(pagecardnum(m, 6)) = 2 And Val(pagecardnum(m, 5)) = 2 Then
                        Randomize
                        n = Int(Rnd() * 6) + 1
                        If n Mod 2 = 0 Then
                            戰鬥系統類.電腦牌_模擬轉牌_外 m
                        End If
                    End If
                Next
              atkingck(145, 1) = 4
              FormMainMode.trgoi1_Timer
              FormMainMode.trgoi2_Timer
        Case 4
             atkingtrn(1) = Val(atkingtrn(1)) - 1
             atkingck(145, 2) = 0
   End Select
End If
End Sub
Sub 傑多_因果之刻()
Dim m As Integer
If FormMainMode.personatk(3).Caption = "因果之刻" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(146, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "傑多" Then
   Select Case atkingck(146, 1)
      Case 1
            If atkingpagetot(1, 4) >= 4 And atkingck(146, 2) = 0 Then
               atkingck(146, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 3
            End If
            If atkingpagetot(1, 4) < 4 And atkingck(146, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(146, 2) = 0
             End If
      Case 2
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_傑多_因果之刻記錄數(i) = 1
                   atking_傑多_因果之刻記錄數(107) = atking_傑多_因果之刻記錄數(107) + 1
               End If
            Next
            atking_傑多_因果之刻記錄數(108) = 1
      Case 3
            atkingtrn(1) = Val(atkingtrn(1)) + 1
      Case 4
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7425
                   atkingno(i, 6) = 9570
                   atkingno(i, 7) = 146
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
            Do Until atking_傑多_因果之刻記錄數(108) > 106
                If atking_傑多_因果之刻記錄數(atking_傑多_因果之刻記錄數(108)) = 1 Then
                    目前數(16) = atking_傑多_因果之刻記錄數(108)
                    目前數(15) = 26
                    FormMainMode.tr牌組_回牌_使用者.Enabled = True
                    atking_傑多_因果之刻記錄數(目前數(16)) = 0
                    Exit Do
                End If
                atking_傑多_因果之刻記錄數(108) = atking_傑多_因果之刻記錄數(108) + 1
            Loop
            If atking_傑多_因果之刻記錄數(108) >= 106 Then
                If atking_傑多_因果之刻記錄數(107) < 2 Then
                    atking_傑多_因果之刻記錄數(107) = atking_傑多_因果之刻記錄數(107) + 1
                    目前數(22) = 24
                    FormMainMode.等待時間.Enabled = True
                ElseIf atking_傑多_因果之刻記錄數(107) >= 2 Then
                    atkingck(146, 2) = 0
                    Erase atking_傑多_因果之刻記錄數
                    戰鬥系統類.執行動作_技能手動結束
                End If
            End If
   End Select
End If
End Sub
Sub 傑多_因果之幻()
Dim p, i, j As Integer
If FormMainMode.personatk(4).Caption = "因果之幻" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(147, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "傑多" Then
   Select Case atkingck(147, 1)
      Case 1
            If atkingpagetot(1, 3) >= 1 And atkingck(147, 2) = 0 Then
               atkingck(147, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
            End If
            If atkingpagetot(1, 3) < 1 And atkingck(147, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(147, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7590
                   atkingno(i, 6) = 9420
                   atkingno(i, 7) = 147
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingck(147, 1) = 3
        Case 3
             atking_傑多_因果之幻骰量紀錄數(1) = 擲骰後骰傷害數
             擲骰表單溝通暫時變數(2) = 0
             擲骰表單溝通暫時變數(3) = 0
             '========================================
                For p = 1 To Val(FormMainMode.顯示列1.goi1)
                   Randomize Timer
                   i = Int(Rnd() * 6) + 1
                   If i = 1 Or i = 6 Then 擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) + 1
                Next
                For p = 1 To Val(FormMainMode.顯示列1.goi2)
                   Randomize Timer
                   j = Int(Rnd() * 6) + 1
                   If j = 1 Or j = 6 Then 擲骰表單溝通暫時變數(3) = Val(擲骰表單溝通暫時變數(3)) + 1
                Next
                '=============================
                技能動畫顯示階段數 = 1
                atkingck(147, 1) = 4
                FormMainMode.骰子執行完啟動.Enabled = False
                目前數(22) = 12
                FormMainMode.等待時間.Enabled = True
          Case 4
                atking_傑多_因果之幻骰量紀錄數(2) = 擲骰後骰傷害數
                '==========================
                If atking_傑多_因果之幻骰量紀錄數(1) < atking_傑多_因果之幻骰量紀錄數(2) Then
                    擲骰表單溝通暫時變數(2) = atking_傑多_因果之幻骰量紀錄數(2)
                Else
                    擲骰表單溝通暫時變數(2) = atking_傑多_因果之幻骰量紀錄數(1)
                End If
                擲骰後骰傷害數 = Val(擲骰表單溝通暫時變數(2))
                atkingck(147, 2) = 0
                Erase atking_傑多_因果之幻骰量紀錄數
   End Select
End If
End Sub
Sub 阿奇波爾多_大地崩壞()
If FormMainMode.personatk(1).Caption = "大地崩壞" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(149, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿奇波爾多" Then
   Select Case atkingck(149, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 4) >= 3 And atkingck(149, 2) = 0 Then
                   atkingck(149, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 3 And atkingck(149, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(149, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking1-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8055
                   atkingno(i, 6) = 10620
                   atkingno(i, 7) = 149
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\阿奇波爾多\atking1-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(149, 2) = 0
             '=================
             戰鬥系統類.傷害執行_技能直傷_電腦 2, 1
   End Select
End If
End Sub
Sub 阿奇波爾多_致命槍擊()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "致命槍擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(150, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿奇波爾多" Then
   Select Case atkingck(150, 1)
      Case 1
            If movecp > 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                        If pagecardnum(i, 1) = a5a And pagecardnum(i, 2) = 1 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And pagecardnum(i, 2) = 2 Then
                           rrr(2) = rrr(2) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And pagecardnum(i, 2) = 3 Then
                           rrr(3) = rrr(3) + 1
                        End If
                    End If
                 Next
            End If
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingck(150, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingck(150, 2) = 0 Then
                戰鬥系統類.人物技能欄燈開關 True, 2
                atkingck(150, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 9
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingck(150, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingck(150, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 2
                atkingck(150, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 9
              End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             atkingck(150, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking2-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8520
                   atkingno(i, 6) = 8280
                   atkingno(i, 7) = 150
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\阿奇波爾多\atking2-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 阿奇波爾多_劫影攻擊()
If FormMainMode.personatk(3).Caption = "劫影攻擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(151, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿奇波爾多" Then
   Select Case atkingck(151, 1)
      Case 1
            If atkingpagetot(1, 4) >= 1 And atkingck(151, 2) = 0 Then
               atkingck(151, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 1 And atkingck(151, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(151, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 10395
                   atkingno(i, 7) = 151
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(151, 2) = 0
             '======================
             If 擲骰後骰傷害數 > 0 Then
               Do
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 3) = 17 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                          FormMainMode.personcomspe(i).person_turn = 2
                          人物異常狀態資料庫(2, i, 2) = 2
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, i, 17, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                          異常狀態檢查數(17, 1) = 1
                          異常狀態檢查數(17, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub 阿奇波爾多_防護射擊()
If FormMainMode.personatk(4).Caption = "防護射擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(152, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "阿奇波爾多" Then
   Select Case atkingck(152, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (atkingpagetot(1, 5) - atking_阿奇波爾多_防護射擊_槍數值紀錄數)
                   atking_阿奇波爾多_防護射擊_槍數值紀錄數 = atkingpagetot(1, 5)
                   If atkingck(152, 2) = 0 Then
                        atkingck(152, 2) = 1
                        戰鬥系統類.人物技能欄燈開關 True, 4
                        atkingtrn(1) = Val(atkingtrn(1)) + 1
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 2
                   End If
                End If
                If atkingpagetot(1, 5) < 1 And atkingck(152, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(152, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - atking_阿奇波爾多_防護射擊_槍數值紀錄數 - 2
                   atking_阿奇波爾多_防護射擊_槍數值紀錄數 = 0
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             atkingck(152, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7725
                   atkingno(i, 6) = 9345
                   atkingno(i, 7) = 152
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atking_阿奇波爾多_防護射擊_槍數值紀錄數 = 0
   End Select
End If
End Sub
Sub 洛洛妮_逆轉戰局的槍響()
Dim bloodnum As Integer
If FormMainMode.personatk(1).Caption = "逆轉戰局的槍響" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(153, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "洛洛妮" Then
   Select Case atkingck(153, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 4) >= 3 And atkingck(153, 2) = 0 Then
                   atkingck(153, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                End If
                If atkingpagetot(1, 4) < 3 And atkingck(153, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(153, 2) = 0
                 End If
          End If
      Case 2
             atkingtrn(1) = Val(atkingtrn(1)) + 1
      Case 3
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -120
                   atkingno(i, 5) = 7035
                   atkingno(i, 6) = 9540
                   atkingno(i, 7) = 153
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數
                   atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) = liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2))
                   If atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) > 2 Then
                       atkingno(i, 11) = 1
                   Else
                       atkingno(i, 11) = 0
                   End If
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(1) = Val(atkingtrn(1)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) And atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) = atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) > atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2)
                    目前數(15) = 27
                    FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) > atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2)) <= 2 Then
                   atkingck(153, 2) = 0
               Else
                   目前數(24) = 27
                   FormMainMode.等待時間_2.Enabled = True
               End If
            End If
        Case 5
            atkingck(153, 2) = 0
            戰鬥系統類.執行動作_技能手動結束
   End Select
End If
End Sub
Sub 洛洛妮_風暴感知()
If FormMainMode.personatk(2).Caption = "風暴感知" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(154, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "洛洛妮" Then
   Select Case atkingck(154, 1)
        Case 1
             If atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 And atkingck(154, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(154, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + pageqlead(2) * 2
            ElseIf (atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1) And atkingck(154, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(154, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(2) * 2
            End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             atkingck(154, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 840
                   atkingno(i, 5) = 8325
                   atkingno(i, 6) = 9285
                   atkingno(i, 7) = 63
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 洛洛妮_砲擊壓制()
If FormMainMode.personatk(3).Caption = "砲擊壓制" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(155, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "洛洛妮" Then
   Select Case atkingck(155, 1)
        Case 1
             If movecp = 3 Then
                     If atkingpagetot(1, 5) >= 4 And atkingpagetot(1, 4) >= 2 And atkingck(155, 2) = 0 Then
                       戰鬥系統類.人物技能欄燈開關 True, 3
                       atkingck(155, 2) = 1
                       atkingtrn(1) = Val(atkingtrn(1)) + 1
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 8
                    ElseIf (atkingpagetot(1, 5) < 4 Or atkingpagetot(1, 4) < 2) And atkingck(155, 2) = 1 Then
                       戰鬥系統類.人物技能欄燈開關 False, 3
                       atkingck(155, 2) = 0
                       atkingtrn(1) = Val(atkingtrn(1)) - 1
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 8
                    End If
             End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8430
                   atkingno(i, 6) = 8985
                   atkingno(i, 7) = 63
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingck(155, 2) = 0
             If 擲骰後骰傷害數 > 0 Then
                    Do
                         For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                           If 人物異常狀態資料庫(2, i, 3) = 4 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                                FormMainMode.personcomspe(j).person_num = 10
                                FormMainMode.personcomspe(j).person_turn = 1
                                人物異常狀態資料庫(2, j, 1) = 10
                                人物異常狀態資料庫(2, j, 2) = 1
                               Exit Do
                           End If
                         Next
                         For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                            If 人物異常狀態資料庫(2, i, 2) = 0 Then
                               戰鬥系統類.人物異常狀態表設定_初設 2, i, 4, app_path & "gif\異常狀態\atkdown.gif", 10, 1
                               異常狀態檢查數(4, 1) = 1
                               異常狀態檢查數(4, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
              End If
   End Select
End If
End Sub
Sub 洛洛妮_貪婪之刃與嗜血之槍()
If FormMainMode.personatk(4).Caption = "貪婪之刃與嗜血之槍" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(156, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "洛洛妮" Then
   Select Case atkingck(156, 1)
        Case 1
             If movecp = 1 Then
                     If atkingpagetot(1, 1) >= 5 And atkingpagetot(1, 5) >= 5 And atkingck(156, 2) = 0 Then
'                     If pageqlead(1) >= 1 And atkingck(156, 2) = 0 Then
                       戰鬥系統類.人物技能欄燈開關 True, 4
                       atkingck(156, 2) = 1
                       atkingtrn(1) = Val(atkingtrn(1)) + 1
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                    ElseIf (atkingpagetot(1, 1) < 5 Or atkingpagetot(1, 5) < 5) And atkingck(156, 2) = 1 Then
'                    ElseIf pageqlead(1) < 1 And atkingck(156, 2) = 1 Then
                       戰鬥系統類.人物技能欄燈開關 False, 4
                       atkingck(156, 2) = 0
                       atkingtrn(1) = Val(atkingtrn(1)) - 1
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                    End If
             End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9495
                   atkingno(i, 6) = 9360
                   atkingno(i, 7) = 156
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
             atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0
        Case 3
             For i = 1 To 106
                 If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                     目前數(16) = i
                     atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 + 1
                     FormMainMode.tr電腦牌_翻牌.Enabled = True
                     Exit Sub
                 End If
             Next
             If i = 107 Then
                 If atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0 Then
                     For k = 1 To 3
                         戰鬥系統類.傷害執行_技能直傷_電腦 3, k
                     Next
                     目前數(22) = 27
                     FormMainMode.等待時間.Enabled = True
                 ElseIf atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 1 Or atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 2 Then
                     目前數(24) = 28
                     FormMainMode.等待時間_2.Enabled = True
                 ElseIf atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 > 2 Then
                     atkingck(156, 2) = 0
                     戰鬥系統類.執行動作_技能手動結束
                 End If
             End If
        Case 4
             目前數(17) = 10
             FormMainMode.tr電腦牌_偷牌.Enabled = True
        Case 5
             If atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0 Then
                atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 99
                目前數(22) = 27
                FormMainMode.等待時間.Enabled = True
             ElseIf atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 1 Then
                For k = 1 To 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, k
                Next
                atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 99
                目前數(24) = 28
                 FormMainMode.等待時間_2.Enabled = True
             ElseIf atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 2 Then
                For k = 1 To 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, k
                Next
                atkingck(156, 2) = 0
                戰鬥系統類.執行動作_技能手動結束
             Else
                 atkingck(156, 2) = 0
                戰鬥系統類.執行動作_技能手動結束
             End If
   End Select
End If
End Sub
Sub 克頓_竊取資料()
If FormMainMode.personatk(1).Caption = "竊取資料" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(157, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "克頓" Then
   Select Case atkingck(157, 1)
      Case 1
            If atkingpagetot(1, 4) >= 2 And atkingck(157, 2) = 0 Then
'            If pageqlead(1) >= 1 And atkingck(157, 2) = 0 Then
               atkingck(157, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 2 And atkingck(157, 2) = 1 Then
'            If pageqlead(1) < 1 And atkingck(157, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(157, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5655
                   atkingno(i, 6) = 9855
                   atkingno(i, 7) = 157
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             技能動畫顯示階段數 = 10
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.防禦階段_階段初始.Enabled = False
        Case 3
                階段狀態數 = 3
                目前數(17) = 2
                If pageqlead(2) > 0 Then
                      Do
                          Randomize
                          m = Int(Rnd() * 106) + 1
                          If Val(pagecardnum(m, 6)) = 2 And Val(pagecardnum(m, 5)) = 2 Then
                              atking_克頓_竊取資料_奪牌紀錄數(1) = m
                              atkingck(157, 1) = 5
                              戰鬥系統類.電腦牌_模擬按牌_外 m
                              Exit Do
                          End If
                      Loop
                    FormMainMode.trgoi1_Timer
                    FormMainMode.trgoi2_Timer
                Else
                    atkingtrn(1) = Val(atkingtrn(1)) - 1
                    atkingck(157, 2) = 0
                    Erase atking_克頓_竊取資料_奪牌紀錄數
                End If
        Case 4
             atkingtrn(1) = Val(atkingtrn(1)) - 1
             atkingck(157, 2) = 0
             Erase atking_克頓_竊取資料_奪牌紀錄數
        Case 5
             目前數(17) = 2
             atkingck(157, 1) = 4
             atking_克頓_竊取資料_奪牌紀錄數(2) = 目前數(9)
             '=========將座標指定至使用者手牌
             戰鬥系統類.座標計算_使用者手牌
             戰鬥系統類.執行動作_電腦牌_偷牌_使用者 atking_克頓_竊取資料_奪牌紀錄數(1)
             FormMainMode.card(atking_克頓_竊取資料_奪牌紀錄數(1)).Width = 810
             FormMainMode.card(atking_克頓_竊取資料_奪牌紀錄數(1)).Height = 1260
             FormMainMode.card(atking_克頓_竊取資料_奪牌紀錄數(1)).Picture = LoadPicture(app_path & "card\" & pagecardnum(atking_克頓_竊取資料_奪牌紀錄數(1), 8) & "-" & pageonin(atking_克頓_竊取資料_奪牌紀錄數(1)) & ".bmp")
             目前數(9) = atking_克頓_竊取資料_奪牌紀錄數(2)
             目前數(15) = 0
   End Select
End If
End Sub
Sub 克頓_逃亡計畫()
Dim rrr(1 To 2) As Integer '牌判斷暫時變數
Dim au As Integer
If FormMainMode.personatk(2).Caption = "逃亡計畫" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(158, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "克頓" Then
   Select Case atkingck(158, 1)
      Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                    If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(2) = rrr(2) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1) And atkingck(158, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingck(158, 2) = 0 Then
                戰鬥系統類.人物技能欄燈開關 True, 2
                atkingck(158, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1) And atkingck(158, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingck(158, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 2
                atkingck(158, 2) = 0
                atkingtrn(1) = Val(atkingtrn(1)) - 1
              End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7200
                   atkingno(i, 6) = 9990
                   atkingno(i, 7) = 158
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
               Randomize
               m = Int(Rnd() * 2) + 2
               au = 1
               Do
                    If liveus(角色待機人物紀錄數(1, m)) > 0 Then
                        戰鬥系統類.傷害執行_技能直傷_使用者 3, m
                        Exit Do
                    End If
                    If au < 2 Then
                        au = au + 1
                        If m = 2 Then
                            m = 3
                        Else
                            m = 2
                        End If
                    Else
                        戰鬥系統類.傷害執行_技能直傷_使用者 3, 1
                        Exit Do
                    End If
               Loop
               擲骰後骰傷害數 = 0
               擲骰表單溝通暫時變數(2) = 0
               atkingck(158, 2) = 0
   End Select
End If
End Sub
Sub 克頓_隱蔽射擊()
Dim p, i, j As Integer
If FormMainMode.personatk(3).Caption = "隱蔽射擊" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(159, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "克頓" Then
   Select Case atkingck(159, 1)
      Case 1
         If movecp > 1 Then
            If atkingpagetot(1, 5) >= 2 And atkingpagetot(1, 3) >= 1 And atkingck(159, 2) = 0 Then
               atkingck(159, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
            End If
            If (atkingpagetot(1, 5) < 2 Or atkingpagetot(1, 3) < 1) And atkingck(159, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(159, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
             End If
         End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6555
                   atkingno(i, 6) = 10110
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingck(159, 1) = 3
        Case 3
             If liveus(角色人物對戰人數(1, 2)) = liveusmax(角色人物對戰人數(1, 2)) Then
                    atking_克頓_隱蔽射擊骰量紀錄數(1) = 擲骰後骰傷害數
                    擲骰表單溝通暫時變數(2) = 0
                    擲骰表單溝通暫時變數(3) = 0
                    '========================================
                       For p = 1 To Val(FormMainMode.顯示列1.goi1)
                          Randomize Timer
                          i = Int(Rnd() * 6) + 1
                          If i = 1 Or i = 6 Then 擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) + 1
                       Next
                       For p = 1 To Val(FormMainMode.顯示列1.goi2)
                          Randomize Timer
                          j = Int(Rnd() * 6) + 1
                          If j = 1 Or j = 6 Then 擲骰表單溝通暫時變數(3) = Val(擲骰表單溝通暫時變數(3)) + 1
                       Next
                       '=============================
                       技能動畫顯示階段數 = 1
                       atkingck(159, 1) = 4
                       FormMainMode.骰子執行完啟動.Enabled = False
                       目前數(22) = 12
                       FormMainMode.等待時間.Enabled = True
                Else
                       atkingck(159, 2) = 0
                       FormMainMode.骰子執行完啟動.Enabled = True
                       Erase atking_克頓_隱蔽射擊骰量紀錄數
                End If
          Case 4
                atking_克頓_隱蔽射擊骰量紀錄數(2) = 擲骰後骰傷害數
                '==========================
                擲骰表單溝通暫時變數(2) = atking_克頓_隱蔽射擊骰量紀錄數(1) + atking_克頓_隱蔽射擊骰量紀錄數(2)
                擲骰後骰傷害數 = Val(擲骰表單溝通暫時變數(2))
                atkingck(159, 2) = 0
                Erase atking_克頓_隱蔽射擊骰量紀錄數
   End Select
End If
End Sub
Sub 克頓_惡意情報()
Dim rrr(1 To 2) As Integer '牌判斷暫時變數
Dim au As Integer
If FormMainMode.personatk(4).Caption = "惡意情報" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(160, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "克頓" Then
   Select Case atkingck(160, 1)
      Case 1
            If movecp > 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 3 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 3 Then
                           rrr(2) = rrr(2) + 1
                        End If
                    End If
                 Next
                 '========================
                 If (rrr(1) >= 1 And rrr(2) >= 1) And atkingpagetot(1, 4) >= 2 And atkingck(160, 2) = 0 Then
'                 If pageqlead(1) >= 1 And atkingck(160, 2) = 0 Then
                    戰鬥系統類.人物技能欄燈開關 True, 4
                    atkingck(160, 2) = 1
                 ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or atkingpagetot(1, 4) < 2) And atkingck(160, 2) = 1 Then
'                 ElseIf pageqlead(1) < 1 And atkingck(160, 2) = 1 Then
                    戰鬥系統類.人物技能欄燈開關 False, 4
                    atkingck(160, 2) = 0
                  End If
            End If
      Case 2
             atkingtrn(1) = Val(atkingtrn(1)) + 1
      Case 3
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7050
                   atkingno(i, 6) = 10005
                   atkingno(i, 7) = 160
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
             atkingtrn(1) = Val(atkingtrn(1)) - 1
             '=====================
              For i = 1 To 106
                   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                      atking_克頓_惡意情報紀錄數(i) = 1
                      atking_克頓_惡意情報紀錄數(0) = Val(atking_克頓_惡意情報紀錄數(0)) + 1
                   End If
               Next
        Case 4
               For i = 1 To 106
                   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                      階段狀態數 = 3
                      戰鬥系統類.電腦牌_模擬按牌_外 i
                      目前數(15) = 28
                      Exit Sub
                    End If
               Next
               If i = 107 And atking_克頓_惡意情報紀錄數(0) > 0 Then
                    For k = 1 To 106
                         If atking_克頓_惡意情報紀錄數(k) = 1 Then
                             atking_克頓_惡意情報紀錄數(k) = 0
                             階段狀態數 = 3
                             戰鬥系統類.電腦牌_模擬按牌 k
                             目前數(17) = 11
                             Exit Sub
                         End If
                    Next
                End If
         Case 5
               For k = 1 To 106
                     If atking_克頓_惡意情報紀錄數(k) = 1 Then
                         atking_克頓_惡意情報紀錄數(k) = 0
                         階段狀態數 = 3
                         戰鬥系統類.電腦牌_模擬按牌 k
                         目前數(17) = 11
                         Exit Sub
                     End If
                Next
                If k = 107 Then
                    atkingck(160, 2) = 0
                    Erase atking_克頓_惡意情報紀錄數
                    戰鬥系統類.執行動作_技能手動結束
                End If
   End Select
End If
End Sub
Sub 露緹亞_腐朽之靈()
If FormMainMode.personatk(1).Caption = "腐朽之靈" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(98, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "露緹亞" Then
   Select Case atkingck(98, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 1) >= 3 And atkingck(98, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(98, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 1) < 3 And atkingck(98, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(98, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             Do
                For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 3) = 34 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                      FormMainMode.personcomspe(i).person_turn = 3
                      人物異常狀態資料庫(2, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                   If 人物異常狀態資料庫(2, i, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 2, i, 34, app_path & "gif\異常狀態\咒縛.gif", 0, 3
                      異常狀態檢查數(34, 1) = 1
                      異常狀態檢查數(34, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            atkingck(98, 2) = 0
   End Select
End If
End Sub
Sub 露緹亞_朦朧之暗()
If FormMainMode.personatk(2).Caption = "朦朧之暗" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(99, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "露緹亞" Then
   Select Case atkingck(99, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 And atkingck(99, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(99, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 5) < 1 Or atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1) And atkingck(99, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(99, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 99
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(99, 2) = 0
            '=================
            戰鬥系統類.執行動作_距離變更 1
   End Select
End If
End Sub
Sub 露緹亞_暗影之翼()
If FormMainMode.personatk(3).Caption = "暗影之翼" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(100, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "露緹亞" Then
   Select Case atkingck(100, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 2) >= 1 And atkingpagetot(1, 3) >= 1 And atkingck(100, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(100, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                End If
                If (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 2) < 1 Or atkingpagetot(1, 3) < 1) And atkingck(100, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(100, 2) = 0
                 End If
          End If
      Case 2
           atkingtrn(1) = Val(atkingtrn(1)) + 1
      Case 3
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 100
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            atkingck(100, 2) = 0
            '=================
            戰鬥系統類.執行動作_距離變更 3
            If Val(擲骰後骰傷害數) < 0 Then
                回復執行_使用者 1, 1
            End If
   End Select
End If
End Sub
Sub 露緹亞_渦騎劍閃(ByVal Index As Integer)
Dim aw As Integer
If FormMainMode.personatk(4).Caption = "渦騎劍閃" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(101, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "露緹亞" Then
   Select Case atkingck(101, 1)
      Case 1
             If movecp = 3 Then
                 If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 5) >= 4 And atkingpagetot(1, 4) >= 1 And atkingck(101, 2) = 0 Then
'                     aw = Int(atkingpagetot(1, 1) / 2 + 0.5)
                     For i = 1 To 106
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                           aw = aw + 1
                        End If
                     Next
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_露緹亞_渦騎劍閃計算張數紀錄數) * 5 + 8
                     atking_露緹亞_渦騎劍閃計算張數紀錄數 = aw
                     戰鬥系統類.人物技能欄燈開關 True, 4
                     atkingck(101, 2) = 1
                     atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
            End If
      Case 2
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 And atkingck(101, 2) = 1 Then
                   For i = 1 To 106
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                           aw = aw + 1
                        End If
                   Next
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_露緹亞_渦騎劍閃計算張數紀錄數) * 5
                   atking_露緹亞_渦騎劍閃計算張數紀錄數 = aw
            End If
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 And atkingck(101, 2) = 1 Then
                   If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 5) >= 4 And atkingpagetot(1, 4) >= 1 And atkingck(101, 2) = 1 Then
                        For i = 1 To 106
                             If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                                aw = aw + 1
                             End If
                        Next
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_露緹亞_渦騎劍閃計算張數紀錄數) * 5
                        atking_露緹亞_渦騎劍閃計算張數紀錄數 = aw
                   ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 5) < 4 Or atkingpagetot(1, 4) < 1) And atkingck(101, 2) = 1 Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - (atking_露緹亞_渦騎劍閃計算張數紀錄數 * 5) - 8
                        戰鬥系統類.人物技能欄燈開關 False, 4
                        atkingck(101, 2) = 0
                        atkingck(101, 1) = 1
                        atkingtrn(1) = Val(atkingtrn(1)) - 1
                        atking_露緹亞_渦騎劍閃計算張數紀錄數 = 0
                    End If
            End If
            FormMainMode.trgoi1.Enabled = True
    Case 3
        If Val(pagecardnum(Index, 5)) = 1 And atkingck(101, 2) = 1 Then
               If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 5) >= 4 And atkingpagetot(1, 4) >= 1 Then
                    For i = 1 To 106
                         If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                            aw = aw + 1
                         End If
                    Next
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (aw - atking_露緹亞_渦騎劍閃計算張數紀錄數) * 5
                    atking_露緹亞_渦騎劍閃計算張數紀錄數 = aw
               ElseIf (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 5) < 4 Or atkingpagetot(1, 4) < 1) Then
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - (atking_露緹亞_渦騎劍閃計算張數紀錄數 * 5) - 8
                    戰鬥系統類.人物技能欄燈開關 False, 4
                    atkingck(101, 2) = 0
                    atkingck(101, 1) = 1
                    atkingtrn(1) = Val(atkingtrn(1)) - 1
                    atking_露緹亞_渦騎劍閃計算張數紀錄數 = 0
                End If
        End If
        FormMainMode.trgoi1.Enabled = True
      Case 4
             戰鬥系統類.人物技能欄燈開關 False, 4
             atkingck(101, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8760
                   atkingno(i, 6) = 10530
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atking_露緹亞_渦騎劍閃計算張數紀錄數 = 0
   End Select
End If
End Sub
Sub 艾蕾可_王座之炎()
Dim dge As Integer
If FormMainMode.personatk(1).Caption = "王座之炎" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(102, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾蕾可" Then
   Select Case atkingck(102, 1)
        Case 1
             If atkingpagetot(1, 1) >= 5 And atkingck(102, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 1
               atkingck(102, 2) = 1
               atkingck(102, 1) = 2
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + pageqlead(1) * 3
               atking_艾蕾可_王座之炎計算出牌張數紀錄數 = pageqlead(1)
            ElseIf atkingpagetot(1, 4) < 2 And atkingck(102, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 1
               atkingck(102, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 3
               atking_艾蕾可_王座之炎計算出牌張數紀錄數 = 0
            End If
        Case 2
                 If atkingpagetot(1, 1) < 5 Then
                     戰鬥系統類.人物技能欄燈開關 False, 1
                     atkingck(102, 2) = 0
                     atkingtrn(1) = Val(atkingtrn(1)) - 1
                     If pageqlead(1) = atking_艾蕾可_王座之炎計算出牌張數紀錄數 Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 3
                     Else
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - pageqlead(1) * 3 - 3
                     End If
                     atking_艾蕾可_王座之炎計算出牌張數紀錄數 = 0
                     atkingck(102, 1) = 1
                  End If
                  If atkingck(102, 2) = 1 Then
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + (pageqlead(1) - Val(atking_艾蕾可_王座之炎計算出牌張數紀錄數)) * 3
                     atking_艾蕾可_王座之炎計算出牌張數紀錄數 = pageqlead(1)
                  End If
                  FormMainMode.trgoi1.Enabled = True
        Case 3
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking1-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10725
                   atkingno(i, 7) = 102
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾蕾可\atking1-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            dge = Val(FormMainMode.pageusglead.Caption)
            If dge > 4 Then dge = 4
            擲骰後骰傷害數 = Val(擲骰後骰傷害數) - dge
            擲骰表單溝通暫時變數(2) = 擲骰後骰傷害數
            atking_艾蕾可_王座之炎計算出牌張數紀錄數 = 0
            atkingck(102, 2) = 0
   End Select
End If
End Sub
Sub 艾蕾可_白百合()
Dim dge As Integer
If FormMainMode.personatk(2).Caption = "白百合" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(103, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾蕾可" Then
   Select Case atkingck(103, 1)
        Case 1
             If movecp < 3 Then
                 If pageqlead(1) >= 2 And atkingck(103, 2) = 0 Then
                   atkingck(103, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If pageqlead(1) < 2 And atkingck(103, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(103, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
             End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking2-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 103
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾蕾可\atking2-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(103, 2) = 0
            '===================
            If 擲骰後骰傷害數 > 0 Then
               戰鬥系統類.執行動作_清除所有異常狀態_電腦
               '==================
               Do
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 3) = 23 And 人物異常狀態資料庫(2, i, 2) > 0 Then
                          FormMainMode.personcomspe(i).person_turn = 1
                          人物異常狀態資料庫(2, i, 2) = 1
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, i, 23, app_path & "gif\異常狀態\atkingerr.gif", 0, 1
                          異常狀態檢查數(23, 1) = 1
                          異常狀態檢查數(23, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub 艾蕾可_聖王威光()
Dim dge As Integer
If FormMainMode.personatk(3).Caption = "聖王威光" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(104, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾蕾可" Then
   Select Case atkingck(104, 1)
        Case 1
             If atkingpagetot(1, 4) >= 3 And atkingck(104, 2) = 0 Then
               atkingck(104, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 3
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 3 And atkingck(104, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 3
               atkingck(104, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking3-1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 104
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾蕾可\atking3-2_1.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atking_艾蕾可_聖王威光紀錄數(1) = Val(FormMainMode.顯示列1.goi2)
             atking_艾蕾可_聖王威光紀錄數(2) = pageqlead(2)
        Case 4
            atkingck(104, 2) = 0
            '===================
            If 擲骰後骰傷害數 <= 0 Then
               dge = Int(atking_艾蕾可_聖王威光紀錄數(1) / 4 + 0.9)
               戰鬥系統類.傷害執行_技能直傷_電腦 dge, 1
            End If
            '===================
            If atking_艾蕾可_聖王威光紀錄數(2) = 0 Then
                戰鬥系統類.傷害執行_技能直傷_電腦 2, 1
            End If
            '===================
            Erase atking_艾蕾可_聖王威光紀錄數
   End Select
End If
End Sub
Sub 艾蕾可_救濟天使()
Dim dge As Integer
If FormMainMode.personatk(4).Caption = "救濟天使" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(105, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "艾蕾可" Then
   Select Case atkingck(105, 1)
        Case 1
             If atkingpagetot(1, 4) >= 5 And atkingck(105, 2) = 0 Then
               atkingck(105, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 4
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If atkingpagetot(1, 4) < 5 And atkingck(105, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 4
               atkingck(105, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
             End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10590
                   atkingno(i, 7) = 105
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingck(105, 2) = 0
            '===================
            If liveus(角色待機人物紀錄數(1, 2)) = 0 And liveus(角色待機人物紀錄數(1, 3)) = 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 7 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 7
                              FormMainMode.personusspe(j).person_turn = 4
                              人物異常狀態資料庫(1, j, 1) = 7
                              人物異常狀態資料庫(1, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 7, app_path & "gif\異常狀態\atkup.gif", 7, 4
                          異常狀態檢查數(7, 1) = 1
                          異常狀態檢查數(7, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '=================================
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 8 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 7
                              FormMainMode.personusspe(j).person_turn = 4
                              人物異常狀態資料庫(1, j, 1) = 7
                              人物異常狀態資料庫(1, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 8, app_path & "gif\異常狀態\defup.gif", 7, 4
                          異常狀態檢查數(8, 1) = 1
                          異常狀態檢查數(8, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '================
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 37 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_turn = 4
                              人物異常狀態資料庫(1, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 37, app_path & "gif\異常狀態\再生.gif", 0, 4
                          異常狀態檢查數(37, 1) = 1
                          異常狀態檢查數(37, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '================
            Else
                '================
                For i = 2 To 3
                     If FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption > 0 Then
                        Do
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                                  If 人物異常狀態資料庫(1, j, 3) = 35 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_turn = 1
                                      人物異常狀態資料庫(1, j, 2) = 1
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(1, i) - 1) + 1 To 14 * 角色待機人物紀錄數(1, i)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 35, app_path & "gif\異常狀態\庇護.png", 0, 1
                                  異常狀態檢查數(35, 1) = 1
                                  異常狀態檢查數(35, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        戰鬥系統類.回復執行_使用者 1, i
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub 梅莉_夢幻魔杖()
Dim m As Integer, n As Integer, bd As Integer
If FormMainMode.personatk(1).Caption = "夢幻魔杖" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(106, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅莉" Then
   Select Case atkingck(106, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 3 And atkingck(106, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
                   atkingck(106, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 5) < 3 And atkingck(106, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   atkingck(106, 2) = 0
                 End If
          End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9405
                   atkingno(i, 6) = 10245
                   atkingno(i, 7) = 106
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             技能動畫顯示階段數 = 11
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.攻擊階段_階段2.Enabled = False
        Case 3
            Randomize
            m = Int(Rnd() * 100) + 1
            If liveus(角色人物對戰人數(1, 2)) <= liveus41(角色人物對戰人數(1, 2)) Then
                Randomize
                bd = Int(Rnd() * 2) + 1
            End If
            If m Mod (2 - bd) = 0 Then '===相當於50~100%機率
                 Randomize
                 n = Int(Rnd() * 100) + 1
                 If liveus(角色人物對戰人數(1, 2)) <= liveusmax(角色人物對戰人數(1, 2)) Then
                     bd = liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2))
                     If bd > 8 Then bd = 8
                 End If
                 If n Mod (10 - bd) = 0 Then '===相當於10~50%機率
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) * 4
                     FormMainMode.messageus.AddItem "夢幻魔杖效果發動!  攻擊力變為4倍"
                     戰鬥系統類.自動捲軸捲動
                Else
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) * 2
                     FormMainMode.messageus.AddItem "夢幻魔杖效果發動!  攻擊力變為2倍"
                     戰鬥系統類.自動捲軸捲動
                End If
                FormMainMode.trgoi1_Timer
            End If
            atkingck(106, 1) = 4
        Case 4
             atkingtrn(1) = Val(atkingtrn(1)) - 1
             atkingck(106, 1) = 5
        Case 5
             atkingck(106, 2) = 0
             If Val(擲骰後骰傷害數) <= 0 Then
                 Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                        If 人物異常狀態資料庫(2, j, 3) = 5 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                         FormMainMode.personcomspe(j).person_num = 3
                         FormMainMode.personcomspe(j).person_turn = 3
                         人物異常狀態資料庫(2, j, 1) = 3
                         人物異常狀態資料庫(2, j, 2) = 3
                         Exit Do
                        End If
                    Next
                   For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, j, 2) = 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 2, j, 5, app_path & "gif\異常狀態\defdown.gif", 3, 3
                         異常狀態檢查數(5, 1) = 1
                         異常狀態檢查數(5, 2) = 1
                         Exit Do
                     End If
                   Next
                Loop
            End If
   End Select
End If
End Sub
Sub 梅莉_徬徨夢羽()
Dim m As Integer, n As Integer, bd As Integer
If FormMainMode.personatk(2).Caption = "徬徨夢羽" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(107, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅莉" Then
   Select Case atkingck(107, 1)
      Case 1
            If atkingpagetot(1, 5) >= 1 And atkingpagetot(1, 2) >= 3 And atkingck(107, 2) = 0 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
               atkingck(107, 2) = 1
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            End If
            If (atkingpagetot(1, 5) < 1 Or atkingpagetot(1, 2) < 3) And atkingck(107, 2) = 1 Then
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               atkingck(107, 2) = 0
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8865
                   atkingno(i, 6) = 9210
                   atkingno(i, 7) = 107
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             技能動畫顯示階段數 = 10
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.防禦階段_階段初始.Enabled = False
        Case 3
            bd = 0
            Randomize
            m = Int(Rnd() * 100) + 1
            If liveus(角色人物對戰人數(1, 2)) <= liveus41(角色人物對戰人數(1, 2)) Then
                bd = 1
            End If
            If m Mod (3 - bd) = 0 Then '===相當於33~50%機率
                 Randomize
                 n = Int(Rnd() * 100) + 1
                 If liveus(角色人物對戰人數(1, 2)) <= liveusmax(角色人物對戰人數(1, 2)) Then
                     bd = liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2))
                     If bd > 8 Then bd = 8
                 End If
                 If n Mod (10 - bd) = 0 Then '===相當於10~50%機率
                     攻擊防禦骰子總數(2) = 0
                     FormMainMode.messageus.AddItem "徬徨夢羽效果發動!  對手攻擊力變為0"
                     戰鬥系統類.自動捲軸捲動
                Else
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) \ 2
                     FormMainMode.messageus.AddItem "徬徨夢羽效果發動!  對手攻擊力變為1/2"
                     戰鬥系統類.自動捲軸捲動
                End If
            Else
                攻擊防禦骰子總數(2) = Int((攻擊防禦骰子總數(2) * 2) / 3)
                FormMainMode.messageus.AddItem "徬徨夢羽效果發動!  對手攻擊力變為2/3"
                戰鬥系統類.自動捲軸捲動
            End If
            FormMainMode.trgoi2_Timer
            '=====================
            戰鬥系統類.回復執行_使用者 1, 1
            '=====================
            atkingck(107, 1) = 4
        Case 4
             atkingtrn(1) = Val(atkingtrn(1)) - 1
             atkingck(107, 2) = 0
   End Select
End If
End Sub
Sub 梅莉_綿羊幻夢()
Dim bloodnum As Integer
If FormMainMode.personatk(3).Caption = "綿羊幻夢" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(108, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅莉" Then
   Select Case atkingck(108, 1)
      Case 1
           If movecp < 3 Then
                If pageqlead(1) >= 2 And atkingck(108, 2) = 0 Then
                    atkingck(108, 2) = 1
                    戰鬥系統類.人物技能欄燈開關 True, 3
                 End If
                 If pageqlead(1) < 2 And atkingck(108, 2) = 1 Then
                    戰鬥系統類.人物技能欄燈開關 False, 3
                    atkingck(108, 2) = 0
                  End If
            End If
      Case 2
             atkingtrn(1) = Val(atkingtrn(1)) + 1
      Case 3
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6930
                   atkingno(i, 6) = 9540
                   atkingno(i, 7) = 108
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_梅莉_綿羊幻夢_抽牌紀錄數
                   Select Case liveus(角色人物對戰人數(1, 2))
                       Case Is >= 5
                           atking_梅莉_綿羊幻夢_抽牌紀錄數(2) = 4
                           atkingno(i, 11) = 1
                       Case Else
                           atking_梅莉_綿羊幻夢_抽牌紀錄數(2) = 2
                           atkingno(i, 11) = 0
                    End Select
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(1) = Val(atkingtrn(1)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_梅莉_綿羊幻夢_抽牌紀錄數(2) And atking_梅莉_綿羊幻夢_抽牌紀錄數(1) = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_梅莉_綿羊幻夢_抽牌紀錄數(1) = atking_梅莉_綿羊幻夢_抽牌紀錄數(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_梅莉_綿羊幻夢_抽牌紀錄數(1) > atking_梅莉_綿羊幻夢_抽牌紀錄數(2)
                    目前數(15) = 29
                    FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_梅莉_綿羊幻夢_抽牌紀錄數(1) > atking_梅莉_綿羊幻夢_抽牌紀錄數(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_梅莉_綿羊幻夢_抽牌紀錄數(2)) <= 2 Then
                   戰鬥系統類.傷害執行_技能直傷_使用者 1, 1
                   atkingck(108, 2) = 0
               Else
                   目前數(24) = 31
                   FormMainMode.等待時間_2.Enabled = True
               End If
            End If
        Case 5
            atkingck(108, 2) = 0
            戰鬥系統類.傷害執行_技能直傷_使用者 1, 1
            戰鬥系統類.執行動作_技能手動結束
   End Select
End If
End Sub
Sub 梅莉_夢境搖籃()
If FormMainMode.personatk(4).Caption = "夢境搖籃" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(109, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "梅莉" Then
   Select Case atkingck(109, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 1) >= 4 And atkingpagetot(1, 4) >= 2 And atkingck(109, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
                   atkingck(109, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 1) < 4 Or atkingpagetot(1, 4) < 2) And atkingck(109, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(109, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7710
                   atkingno(i, 6) = 9030
                   atkingno(i, 7) = 109
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If liveus(角色人物對戰人數(1, 2)) > 2 Then
                 For i = 1 To 3
                     戰鬥系統類.傷害執行_技能直傷_電腦 1, i
                 Next
            Else
                 For i = 1 To 3
                     戰鬥系統類.傷害執行_技能直傷_電腦 4, i
                 Next
             End If
             atkingck(109, 2) = 0
   End Select
End If
End Sub
Sub 貝琳達_雪光()
Dim bloodnum As Integer
If FormMainMode.personatk(1).Caption = "雪光" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(110, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "貝琳達" Then
   Select Case atkingck(110, 1)
      Case 1
            If atkingpagetot(1, 4) >= 2 And atkingck(110, 2) = 0 Then
                atkingck(110, 2) = 1
                戰鬥系統類.人物技能欄燈開關 True, 1
             End If
             If atkingpagetot(1, 4) < 2 And atkingck(110, 2) = 1 Then
                戰鬥系統類.人物技能欄燈開關 False, 1
                atkingck(110, 2) = 0
              End If
      Case 2
             atkingtrn(1) = Val(atkingtrn(1)) + 1
      Case 3
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7035
                   atkingno(i, 6) = 9510
                   atkingno(i, 7) = 110
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_貝琳達_雪光_抽牌紀錄數
                   Select Case livecom(角色人物對戰人數(2, 2))
                       Case Is = livecommax(角色人物對戰人數(2, 2))
                           atking_貝琳達_雪光_抽牌紀錄數(2) = 4
                           atkingno(i, 11) = 1
                       Case Else
                           atking_貝琳達_雪光_抽牌紀錄數(2) = 2
                           atkingno(i, 11) = 0
                    End Select
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(1) = Val(atkingtrn(1)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_貝琳達_雪光_抽牌紀錄數(2) And atking_貝琳達_雪光_抽牌紀錄數(1) = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_貝琳達_雪光_抽牌紀錄數(1) = atking_貝琳達_雪光_抽牌紀錄數(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_貝琳達_雪光_抽牌紀錄數(1) > atking_貝琳達_雪光_抽牌紀錄數(2)
                    目前數(15) = 31
                    FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_貝琳達_雪光_抽牌紀錄數(1) > atking_貝琳達_雪光_抽牌紀錄數(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_貝琳達_雪光_抽牌紀錄數(2)) <= 2 Then
                   atkingck(110, 2) = 0
               Else
                   目前數(24) = 33
                   FormMainMode.等待時間_2.Enabled = True
               End If
            End If
        Case 5
            atkingck(110, 2) = 0
            戰鬥系統類.執行動作_技能手動結束
   End Select
End If
End Sub
Sub 貝琳達_水晶幻鏡()
If FormMainMode.personatk(2).Caption = "水晶幻鏡" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(111, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "貝琳達" Then
   Select Case atkingck(111, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(1, 2) >= 2 And atkingpagetot(1, 4) >= 2 And atkingck(111, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(111, 2) = 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
            ElseIf (atkingpagetot(1, 2) < 2 Or atkingpagetot(1, 4) < 2) And atkingck(111, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(111, 2) = 0
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
            End If
          End If
        Case 2
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_貝琳達_水晶幻鏡紀錄狀態數(i) = True
               End If
            Next
            目前數(30) = 1
        Case 3
            atkingtrn(1) = Val(atkingtrn(1)) + 1
        Case 4
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7290
                   atkingno(i, 6) = 9120
                   atkingno(i, 7) = 111
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
            Do
                If atking_貝琳達_水晶幻鏡紀錄狀態數(目前數(30)) = True Then
                    目前數(16) = 目前數(30)
                    目前數(15) = 32
                    FormMainMode.tr牌組_回牌_使用者.Enabled = True
                    atking_貝琳達_水晶幻鏡紀錄狀態數(目前數(16)) = False
                    Exit Do
                End If
                目前數(30) = 目前數(30) + 1
            Loop Until 目前數(30) >= 106
            If 目前數(30) >= 106 Then
                If 目前數(30) < 2 Then
                    目前數(30) = 目前數(30) + 1
                    目前數(22) = 29
                    FormMainMode.等待時間.Enabled = True
                ElseIf 目前數(30) >= 2 Then
                    atkingck(111, 2) = 0
                    Erase atking_貝琳達_水晶幻鏡紀錄狀態數
                    戰鬥系統類.執行動作_技能手動結束
                End If
            End If
   End Select
End If
End Sub
Sub 貝琳達_裂地冰牙()
Dim wtr As Integer, wert(1 To 3) As Boolean, wery As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "裂地冰牙" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(112, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "貝琳達" Then
   Select Case atkingck(112, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(1, 5) >= 4 And atkingpagetot(1, 4) >= 1 And atkingck(112, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 4
                   atkingck(112, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 5) < 4 Or atkingpagetot(1, 4) < 1) And atkingck(112, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 4
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(112, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7320
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(擲骰後骰傷害數) > 0 Then
                 Do
                        wtr = Int(Rnd() * 3) + 1
                        If wert(wtr) = False Then
                            wert(wtr) = True
                            wery = wery + 1
                            If livecom(角色待機人物紀錄數(2, wtr)) > 0 Then
                                戰鬥系統類.傷害執行_技能直傷_電腦 2, wtr
                                 Exit Do
                            End If
                        End If
                 Loop Until wery > 3
             End If
             atkingck(112, 2) = 0
   End Select
End If
End Sub
Sub 貝琳達_溶魂之雨()
If FormMainMode.personatk(4).Caption = "溶魂之雨" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(113, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "貝琳達" Then
   Select Case atkingck(113, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(1, 1) >= 1 And atkingpagetot(1, 5) >= 1 _
                   And atkingpagetot(1, 4) >= 1 And atkingpagetot(1, 3) >= 1 And atkingck(113, 2) = 0 Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 10
                        atkingck(113, 2) = 1
                        戰鬥系統類.人物技能欄燈開關 True, 4
                        atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 1) < 1 Or atkingpagetot(1, 5) < 1 _
                    Or atkingpagetot(1, 4) < 1 Or atkingpagetot(1, 3) < 1) And atkingck(113, 2) = 1 Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 10
                        戰鬥系統類.人物技能欄燈開關 False, 4
                        atkingck(113, 2) = 0
                        atkingtrn(1) = Val(atkingtrn(1)) - 1
                        If atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = True Then
                              攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 15
                              atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = False
                        End If
                        If atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = True Then
                             攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 10
                             atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = False
                        End If
                End If
                 '=====================
                 If atkingck(113, 2) = 1 Then
                     If pageqlead(1) >= 10 And atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = False Then
                         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 10
                         atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = True
                     ElseIf pageqlead(1) < 10 And atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = True Then
                         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 10
                         atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = False
                     End If
                     If pageqlead(1) >= 15 And atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = False Then
                         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 15
                         atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = True
                     ElseIf pageqlead(1) < 15 And atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = True Then
                         攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 15
                         atking_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = False
                     End If
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             atkingck(113, 2) = 0
             Erase atking_貝琳達_溶魂之雨_攻擊力加成紀錄數
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8205
                   atkingno(i, 6) = 10080
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 尤莉卡_奸佞的鐵鎚()
Dim wert As Integer '暫時變數
If FormMainMode.personatk(1).Caption = "奸佞的鐵鎚" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(46, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "尤莉卡" Then
   Select Case atkingck(46, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(1, 1) >= 2 And atkingpagetot(1, 4) >= 1 And atkingck(46, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
                   atkingck(46, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 1
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                   '==========
                   If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
                   End If
                   '==========
                End If
                If (atkingpagetot(1, 1) < 2 Or atkingpagetot(1, 4) < 1) And atkingck(46, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
                   戰鬥系統類.人物技能欄燈開關 False, 1
                   atkingck(46, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                   '==========
                   If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                       攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
                   End If
                   '==========
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 1
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking1_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9420
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 46
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then wert = 2 Else wert = 1
             '====================
             Do
                  For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 3) > 0 Then
                             人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
                             If 人物異常狀態資料庫(2, i, 2) = 0 Then
                               '===繼承下一狀態資料
                                戰鬥系統類.異常狀態繼承_電腦
                                If 人物異常狀態資料庫(2, i, 3) = 19 Then
                                    戰鬥系統類.傷害執行_立即死亡_電腦 1 '自壞回合數歸0時執行死亡動作
                                End If
                             Else
                                FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
                             End If
                     End If
                  Next
                  '=====================
                  wert = Val(wert) - 1
             Loop Until wert <= 0
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) > 0 And 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                Randomize
                wert = Int(Rnd() * 3) + 1
                Do
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                     If 人物異常狀態資料庫(1, i, 2) >= 3 And 人物異常狀態資料庫(1, i, 3) = 39 Then
                        Exit Do
                     End If
                     If 人物異常狀態資料庫(1, i, 3) = 39 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 3 Then
                         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + (Val(wert) - 1)
                         If 人物異常狀態資料庫(1, i, 2) > 3 Then 人物異常狀態資料庫(1, i, 2) = 3
                         FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 2) = 0 And (Val(wert) - 1) > 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 1, i, 39, app_path & "gif\異常狀態\臨界.gif", 0, (Val(wert) - 1)
                         異常狀態檢查數(39, 1) = 1
                         異常狀態檢查數(39, 2) = 1
                         Exit Do
                     End If
                   Next
                   If i = 14 * 角色人物對戰人數(1, 2) + 1 And (Val(wert) - 1) = 0 Then Exit Do
                Loop
            End If
            atkingck(46, 2) = 0
            '===============超載技能使用結束
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                atkingck(49, 1) = 6
                技能.尤莉卡_超載 '(階段6)
            End If
            '===============
   End Select
End If
End Sub
Sub 尤莉卡_不善的信仰()
Dim wert As Integer '暫時變數
If FormMainMode.personatk(2).Caption = "不善的信仰" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(47, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "尤莉卡" Then
   Select Case atkingck(47, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(1, 2) >= 3 And atkingck(47, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
                   atkingck(47, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 2
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 2) < 3 And atkingck(47, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
                   戰鬥系統類.人物技能欄燈開關 False, 2
                   atkingck(47, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7170
                   atkingno(i, 6) = 10440
                   atkingno(i, 7) = 47
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then wert = 2 Else wert = 3
                '=============
                If Val(擲骰表單溝通暫時變數(2)) Mod Val(wert) = 0 Then
                    擲骰表單溝通暫時變數(2) = 0
                    擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
                End If
            End If
            '======================================
            If Val(擲骰表單溝通暫時變數(2)) <= 0 And 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                Randomize
                wert = Int(Rnd() * 3) + 1
                Do
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                     If 人物異常狀態資料庫(1, i, 2) >= 3 And 人物異常狀態資料庫(1, i, 3) = 39 Then
                        Exit Do
                     End If
                     If 人物異常狀態資料庫(1, i, 3) = 39 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 3 Then
                         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + (Val(wert) - 1)
                         If 人物異常狀態資料庫(1, i, 2) > 3 Then 人物異常狀態資料庫(1, i, 2) = 3
                         FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 2) = 0 And (Val(wert) - 1) > 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 1, i, 39, app_path & "gif\異常狀態\臨界.gif", 0, (Val(wert) - 1)
                         異常狀態檢查數(39, 1) = 1
                         異常狀態檢查數(39, 2) = 1
                         Exit Do
                     End If
                   Next
                   If i = 14 * 角色人物對戰人數(1, 2) + 1 And (Val(wert) - 1) = 0 Then Exit Do
                Loop
            End If
            atkingck(47, 2) = 0
            '===============超載技能使用結束
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                atkingck(49, 1) = 6
                技能.尤莉卡_超載 '(階段6)
            End If
            '===============
   End Select
End If
End Sub
Sub 尤莉卡_曲惡的安寧()
Dim wert As Integer '暫時變數
If FormMainMode.personatk(3).Caption = "曲惡的安寧" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(48, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "尤莉卡" Then
   Select Case atkingck(48, 1)
      Case 1
           If movecp = 3 Then
                If atkingpagetot(1, 2) >= 3 And atkingpagetot(1, 4) >= 1 And atkingck(48, 2) = 0 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
                   atkingck(48, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 3
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If (atkingpagetot(1, 2) < 3 Or atkingpagetot(1, 4) < 1) And atkingck(48, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
                   戰鬥系統類.人物技能欄燈開關 False, 3
                   atkingck(48, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
          End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 3
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking3_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6450
                   atkingno(i, 6) = 10215
                   atkingno(i, 7) = 48
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                For k = 1 To 3
                    戰鬥系統類.回復執行_使用者 2, k
                Next
            Else
                戰鬥系統類.回復執行_使用者 2, 1
            End If
            '======================================
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                    If 人物異常狀態資料庫(1, i, 3) = 39 Then
                      人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
                      If 人物異常狀態資料庫(1, i, 2) = 0 Then
                        '===繼承下一狀態資料
                         戰鬥系統類.異常狀態繼承_使用者
                         異常狀態檢查數(39, 2) = 0
                     Else
                         FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
                         異常狀態檢查數(39, 1) = 1
                     End If
                   End If
                Next
            End If
            atkingck(48, 2) = 0
            '===============超載技能使用結束
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_使用者 = True Then
                atkingck(49, 1) = 6
                技能.尤莉卡_超載 '(階段6)
            End If
            '===============
   End Select
End If
End Sub
Sub 尤莉卡_超載()
Dim wert As Integer '暫時變數
If FormMainMode.personatk(4).Caption = "超載" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(49, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "尤莉卡" Then
   Select Case atkingck(49, 1)
      Case 1
                If atkingpagetot(1, 4) >= 1 And atkingck(49, 2) = 0 Then
                   atkingck(49, 2) = 1
                   戰鬥系統類.人物技能欄燈開關 True, 4
                   atkingtrn(1) = Val(atkingtrn(1)) + 1
                End If
                If atkingpagetot(1, 4) < 1 And atkingck(49, 2) = 1 Then
                   戰鬥系統類.人物技能欄燈開關 False, 4
                   atkingck(49, 2) = 0
                   atkingtrn(1) = Val(atkingtrn(1)) - 1
                 End If
      Case 2
             戰鬥系統類.人物技能欄燈開關 False, 4
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking4_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7920
                   atkingno(i, 6) = 10005
                   atkingno(i, 7) = 49
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             戰鬥系統類.人物技能欄燈開關 True, 4
             '==================
             Do
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                 If 人物異常狀態資料庫(1, i, 2) >= 3 And 人物異常狀態資料庫(1, i, 3) = 39 Then
                    atking_尤莉卡_超載目前階段紀錄數(3) = 2
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(1, i, 3) = 39 And 人物異常狀態資料庫(1, i, 2) > 0 And 人物異常狀態資料庫(1, i, 2) < 3 Then
                     人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) + 1
                     FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
                     atking_尤莉卡_超載目前階段紀錄數(3) = 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 1, i, 39, app_path & "gif\異常狀態\臨界.gif", 0, 1
                     異常狀態檢查數(39, 1) = 1
                     異常狀態檢查數(39, 2) = 1
                     atking_尤莉卡_超載目前階段紀錄數(3) = 1
                     Exit Do
                 End If
               Next
            Loop
            '========================超載3時執行封印
            If atking_尤莉卡_超載目前階段紀錄數(3) = 2 Then
                Do
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 3) = 22 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 1
                          人物異常狀態資料庫(1, i, 2) = 1
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, i, 22, app_path & "gif\異常狀態\atkingerr.gif", 0, 1
                          異常狀態檢查數(22, 1) = 1
                          異常狀態檢查數(22, 2) = 1
                          Exit Do
                       End If
                    Next
               Loop
            End If
        Case 4
            '========================超載3時攻防2倍階段-執行
            If atking_尤莉卡_超載目前階段紀錄數(3) = 2 Then
                If Val(atking_尤莉卡_超載目前階段紀錄數(4)) = 0 Then
                    atking_尤莉卡_超載目前階段紀錄數(1) = 攻擊防禦骰子總數(1)
                    atking_尤莉卡_超載目前階段紀錄數(2) = 攻擊防禦骰子總數(1) * 2
                    atking_尤莉卡_超載目前階段紀錄數(4) = 1
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) * 2
                ElseIf Val(atking_尤莉卡_超載目前階段紀錄數(4)) = 1 Then
                    atking_尤莉卡_超載目前階段紀錄數(1) = atking_尤莉卡_超載目前階段紀錄數(1) + (攻擊防禦骰子總數(1) - atking_尤莉卡_超載目前階段紀錄數(2))
                    攻擊防禦骰子總數(1) = atking_尤莉卡_超載目前階段紀錄數(1) * 2
                    atking_尤莉卡_超載目前階段紀錄數(2) = atking_尤莉卡_超載目前階段紀錄數(1) * 2
                End If
            End If
        Case 5
            '========================超載3時攻防2倍階段-開始階段時清除資料
            atking_尤莉卡_超載目前階段紀錄數(1) = 0
            atking_尤莉卡_超載目前階段紀錄數(2) = 0
            atking_尤莉卡_超載目前階段紀錄數(4) = 0
        Case 6
            '========================超載技能結束(普通)
            戰鬥系統類.人物技能欄燈開關 False, 4
            atkingck(49, 2) = 0
            Erase atking_尤莉卡_超載目前階段紀錄數
        Case 7
            '========================更換角色時重新載入技能
            If atking_尤莉卡_超載目前階段紀錄數(3) > 0 Then
                戰鬥系統類.人物技能欄燈開關 True, 4
                atking_尤莉卡_超載目前階段紀錄數(1) = 0
                atking_尤莉卡_超載目前階段紀錄數(2) = 0
                atking_尤莉卡_超載目前階段紀錄數(4) = 0
            End If
        Case 8
            '========================超載技能結束(回合結束階段)
            戰鬥系統類.人物技能欄燈開關 False, 4
            atkingck(49, 2) = 0
            If atking_尤莉卡_超載目前階段紀錄數(3) = 2 Then
                戰鬥系統類.執行動作_清除所有異常狀態_使用者
            End If
            Erase atking_尤莉卡_超載目前階段紀錄數
   End Select
End If
End Sub
Sub 羅莎琳_EX_染血之刃()
If FormMainMode.personatk(2).Caption = "Ex染血之刃" And (執行動作_檢查是否有指定異常狀態(1, 22) = False Or atkingck(50, 2) = 1) _
   And FormMainMode.uspi1(角色人物對戰人數(1, 2)) = "羅莎琳" Then
   Select Case atkingck(50, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(1, 1) >= 3 And atkingpagetot(1, 3) >= 2 And atkingck(50, 2) = 0 Then
               戰鬥系統類.人物技能欄燈開關 True, 2
               atkingck(50, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 9
            ElseIf (atkingpagetot(1, 1) < 3 Or atkingpagetot(1, 3) < 2) And atkingck(50, 2) = 1 Then
               戰鬥系統類.人物技能欄燈開關 False, 2
               atkingck(50, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 9
            End If
          End If
        Case 2
             戰鬥系統類.人物技能欄燈開關 False, 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\atkingEX2_1.jpg"
                   atkingno(i, 2) = 1
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 50
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            回復執行_使用者 1, 1
        Case 4
            atkingck(50, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                回復執行_使用者 1, 1
            End If
   End Select
End If
End Sub
