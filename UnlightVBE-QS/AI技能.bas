Attribute VB_Name = "AI技能"
Public atking_AI_艾依查庫_神速之劍計算數值紀錄數(1 To 2) As Integer  '技能-AI-艾依查庫-神速之劍計算劍數值紀錄暫時數(1.目前計算數值/2.(廢除))
Public atking_AI_音音夢_成長模式狀態數(1 To 2) As Integer 'AI-音音夢成長模式狀態檢查數(1.狀態執行階段/2.狀態啟動檢查值)
Public atking_AI_梅倫_Jackpot紀錄數(1 To 2) As Integer '技能-AI-梅倫-Jackpot抽牌紀錄數(1.總共數/2.目前數)
Public atking_AI_帕茉_慈悲的藍眼_tot(1 To 2) As Integer  '技能-帕茉-慈悲的藍眼骰子量紀錄暫時變數(1.數值/2.是否啟動)
Public atking_AI_艾茵_十三隻眼_tot(1 To 2) As Integer '技能-AI-艾茵-十三隻眼骰子量紀錄暫時變數(1.數值/2.是否啟動)
Public atking_AI_傑多_因果之幻骰量紀錄數(1 To 3) As Integer '技能-AI-傑多-因果之幻擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後結果)
Public atking_AI_傑多_因果之刻記錄數(1 To 108) As Integer '技能-AI-傑多-因果之刻紀錄對手出牌編號數(1~106.記錄牌編號/107.總共回張數/108.目前數)
Public atking_AI_阿奇波爾多_防護射擊_槍數值紀錄數 As Integer '技能-AI-阿奇波爾多-防護射擊目前累計加槍數值紀錄數
Public atking_AI_蕾_守護模式狀態啟動值 As Boolean '技能-蕾-AI-Ex-協奏曲-加百烈的守護免除直傷模式啟動值
Public atking_AI_艾伯李斯特_雷擊紀錄數(1 To 2) As Integer '技能-AI-艾伯李斯特-雷擊丟棄對手牌紀錄數(1.總共數/2.目前數)
Public atking_AI_艾伯李斯特_智略紀錄數 As Integer '技能-AI_艾伯李斯特-智略抽牌目前數
Public atking_AI_利恩_反擊的狼煙紀錄數(1 To 2) As Integer '技能-AI-利恩-反擊的狼煙抽牌目前數(1.總共數/2.目前數)
Public atking_AI_瑪格莉特_月光紀錄數(0 To 107) As Integer '技能-AI-瑪格莉特-月光紀錄對手牌編號暫時數(0.目前丟棄張數值/1~106牌編號選擇值/107.總共能丟棄張數值)
Public atking_AI_庫勒尼西_瘋狂眼窩紀錄數 As Integer '技能-AI-庫勒尼西-瘋狂眼窩丟棄對手牌紀錄目前數
Public atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 As Integer '技能-AI-洛洛妮-貪婪之刃與嗜血之槍搶牌目前數
Public atking_AI_史塔夏_殺戮模式狀態數(1 To 5) As Integer 'AI-史塔夏殺戮模式狀態檢查數(1.狀態執行階段/2.狀態啟動檢查值/3.紀錄數值(原始)/4.紀錄數值(變更後)/5.數值紀錄是否啟動)
Public atking_AI_夏洛特_大聖堂骰量紀錄數(1 To 3) As Integer '技能-AI-夏洛特-大聖堂擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後結果)
Public atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 As Integer  '技能-AI-艾蕾可-王座之炎計算出牌張數值紀錄暫時數
Public atking_AI_艾蕾可_聖王威光紀錄數(1 To 2) As Integer  '技能-AI-艾蕾可-聖王威光紀錄暫時數(1.對手當回合防禦力/2.對手當回合出牌數/3.使用者當回合攻擊力)
Public atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 As Integer  '技能-AI-露緹亞-渦騎劍閃計算劍卡張數值紀錄暫時數
Public atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(1 To 2) As Integer '技能-AI-梅莉-綿羊幻夢抽牌目前數(1.總共數/2.目前數)
Public atking_AI_古魯瓦爾多_精神力吸收紀錄數(0 To 106) As Integer '技能-AI-古魯瓦爾多-精神力吸收紀錄對手牌編號暫時數(0.總共張數值/1~106牌編號選擇值)
Public atking_AI_伊芙琳_怠惰的墓表紀錄數(0 To 2) As Integer '技能-AI-伊芙琳-怠惰的墓表紀錄對手牌編號暫時數(0.總共張數值/1~2牌編號)
Public atking_AI_伊芙琳_赤紅石榴階段紀錄數(0 To 106, 1 To 4) As Integer '技能-AI-伊芙琳-赤紅石榴紀錄效果及階段暫時數(0.(1).當前效果/(2).當前效果階段/(3)總共抽牌數量/(4)目前抽/棄牌數量,1~106.(1)牌號選定紀錄值)
Public atking_AI_布勞_發條機構紀錄數 As Integer '技能-AI-布勞-發條機構抽牌目前數
Public atking_AI_貝琳達_雪光_抽牌紀錄數(1 To 2) As Integer '技能-AI-貝琳達-雪光抽牌目前數(1.總共數/2.目前數)
Public atking_AI_貝琳達_水晶幻鏡紀錄狀態數(1 To 106) As Boolean '技能-AI-貝琳達-水晶幻鏡紀錄對手出牌編號數
Public atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1 To 2) As Boolean '技能-AI-貝琳達-溶魂之雨攻擊力加成暫時紀錄數(1.是否10張已+10/2.是否15張已+15)
Public atking_AI_蕾_終曲_無盡輪迴的終結紀錄數 As Integer  '技能-AI-蕾-Ex-終曲-無盡輪迴的終結紀錄對手之防禦牌值暫時數
Public atking_AI_羅莎琳_黑霧幻影紀錄狀態數(1 To 106) As Boolean '技能-AI-羅莎琳-黑霧幻影(普、EX)紀錄對手出牌編號數
Public atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1 To 2) As Integer '技能-AI-洛洛妮-逆轉戰局的槍響抽牌目前數(1.總共數/2.目前數)
Public atking_AI_克頓_竊取資料_奪牌紀錄數(1 To 2) As Integer  '技能-AI-克頓-竊取資料奪取對手出牌牌號紀錄數(1.奪牌編號/2.奪牌原方出牌順序)
Public atking_AI_克頓_隱蔽射擊骰量紀錄數(1 To 3) As Integer '技能-AI-克頓-隱蔽射擊擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後總結果)
Public atking_AI_克頓_惡意情報紀錄數(0 To 106) As Integer '技能-AI-克頓-惡意情報紀錄對手牌編號暫時數(0.目前階段/1~106牌編號選擇值)
Public atking_AI_尤莉卡_超載目前階段紀錄數(1 To 4)  As Integer  '技能-AI-尤莉卡-超載執行目前階段數值紀錄暫時數(1.紀錄數值(原始)/2.紀錄數值(變更後)/3.目前執行階段(總)/4.超載3時攻防骰量加倍是否啟動)

Sub 古魯瓦爾多_猛擊()
Dim rrr As Integer '暫時變數
If FormMainMode.comaiatk(1).Caption = "猛擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(3, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "古魯瓦爾多" Then
   Select Case atkingckai(3, 1)
      Case 1
          If movecp = 1 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          End If
          If rrr >= 2 And atkingckai(3, 2) = 0 Then
             攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
             atkingckai(3, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 2 And atkingckai(3, 2) = 1 Then
             攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
             atkingckai(3, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
      Case 2
             atkingckai(3, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\古魯瓦爾多_猛擊_2.jpeg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 10305
                   atkingno(i, 6) = 8925
                   atkingno(i, 7) = 8
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
Sub 古魯瓦爾多_血之恩賜()
Dim bloodtot As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "血之恩賜" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(62, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "古魯瓦爾多" Then
   Select Case atkingckai(62, 1)
        Case 1
             If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 2 And atkingckai(62, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(62, 2) = 0 Then
               atkingckai(62, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
            ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 2) And atkingckai(62, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(62, 2) = 1 Then
               atkingckai(62, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\Grunwaldatking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6915
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 62
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(FormMainMode.顯示列1.goi1) <= 0 Then
                atkingckai(62, 2) = 0
            End If
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) <= 0 Then
                bloodtot = Abs(Val(擲骰表單溝通暫時變數(2)))
                戰鬥系統類.回復執行_電腦 bloodtot, 1
            End If
            '=============
            atkingckai(62, 2) = 0
   End Select
End If
End Sub
Sub 羊角獸2012_致命格擋()
If FormMainMode.comaiatk(2).Caption = "致命格擋" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(15, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羊角獸2012" Then
   Select Case atkingckai(15, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 1 And atkingckai(15, 2) = 0 Then
               atkingckai(15, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(1)
          End If
      Case 2
             atkingckai(15, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\羊角獸2012\羊角獸2012_致命格擋_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 47
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
Sub 羊角獸2012_致命衝撞()
If FormMainMode.comaiatk(1).Caption = "致命衝撞" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(14, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羊角獸2012" Then
   Select Case atkingckai(14, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 1 And atkingckai(14, 2) = 0 Then
               atkingckai(14, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(1) + 10
          End If
      Case 2
             atkingckai(14, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\羊角獸2012\羊角獸2012_致命衝撞_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 46
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(1) + 10
             戰鬥系統類.直接寫入顯示列數值 2, 攻擊防禦骰子總數(2)
   End Select
End If
End Sub
Sub 吸血姬蕾米雅_吸血()
If FormMainMode.comaiatk(1).Caption = "吸血" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(16, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "吸血姬蕾米雅" Then
   Select Case atkingckai(16, 1)
      Case 1
         If movecp = 1 Then
            If atkingpagetot(2, 1) >= 6 And atkingckai(16, 2) = 0 Then
               atkingckai(16, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 7
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(16, 2) = 1 Then
               atkingckai(16, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 7
            End If
        End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\吸血姬蕾米雅\VampireLAMIAatking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 48
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(16, 1) = 3
       Case 3
            atkingckai(16, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                回復執行_電腦 1, 1
            End If
   End Select
End If
End Sub
Sub 吸血姬蕾米雅_高貴的晚餐()
If FormMainMode.comaiatk(2).Caption = "高貴的晚餐" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(17, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "吸血姬蕾米雅" Then
   Select Case atkingckai(17, 1)
      Case 1
         If movecp > 1 Then
            If atkingpagetot(2, 5) >= 4 And atkingckai(17, 2) = 0 Then
               atkingckai(17, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 5) < 4 And atkingckai(17, 2) = 1 Then
               atkingckai(17, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\吸血姬蕾米雅\VampireLAMIAatking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 17
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(17, 1) = 3
       Case 3
            atkingckai(17, 2) = 0
            戰鬥系統類.傷害執行_技能直傷_使用者 1, 1
   End Select
End If
End Sub
Sub 吸血姬蕾米雅_消失()
If FormMainMode.comaiatk(3).Caption = "消失" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(18, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "吸血姬蕾米雅" Then
   Select Case atkingckai(18, 1)
      Case 1
            If atkingpagetot(2, 2) >= 3 And atkingckai(18, 2) = 0 Then
               atkingckai(18, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 2) < 3 And atkingckai(18, 2) = 1 Then
               atkingckai(18, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\吸血姬蕾米雅\VampireLAMIAatking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 50
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(18, 1) = 3
       Case 3
            atkingckai(18, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) - 1
                擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
            End If
   End Select
End If
End Sub
Sub 妖精王妃_冰結之翼()
If FormMainMode.comaiatk(1).Caption = "冰結之翼" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(8, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "妖精王妃" Then
   Select Case atkingckai(8, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 2 And atkingckai(8, 2) = 0 Then
               atkingckai(8, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\妖精王妃\妖精王妃_冰結之翼_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -840
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 8
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(8, 1) = 3
       Case 3
            atkingckai(8, 2) = 0
                Do
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                     If 人物異常狀態資料庫(1, i, 3) = 10 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         人物異常狀態資料庫(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 2) = 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 1, i, 10, app_path & "gif\異常狀態\atkdown.gif", 5, 3
                         異常狀態檢查數(10, 1) = 1
                         異常狀態檢查數(10, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
   End Select
End If
End Sub
Sub 妖精王妃_煉獄之翼()
If FormMainMode.comaiatk(2).Caption = "煉獄之翼" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(9, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "妖精王妃" Then
   Select Case atkingckai(9, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 2 And atkingckai(9, 2) = 0 Then
               atkingckai(9, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\妖精王妃\妖精王妃_煉獄之翼_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -840
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 9
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(9, 1) = 3
       Case 3
            atkingckai(9, 2) = 0
                Do
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                     If 人物異常狀態資料庫(1, i, 3) = 11 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         人物異常狀態資料庫(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 2) = 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 1, i, 11, app_path & "gif\異常狀態\defdown.gif", 5, 3
                         異常狀態檢查數(11, 1) = 1
                         異常狀態檢查數(11, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
   End Select
End If
End Sub
Sub 妖精王妃_混沌之翼()
If FormMainMode.comaiatk(3).Caption = "混沌之翼" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(10, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "妖精王妃" Then
   Select Case atkingckai(10, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 2 And atkingckai(10, 2) = 0 Then
               atkingckai(10, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\妖精王妃\妖精王妃_混沌之翼_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6180
                   atkingno(i, 6) = 9630
                   atkingno(i, 7) = 10
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(10, 1) = 3
       Case 3
            atkingckai(10, 2) = 0
                Do
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                     If 人物異常狀態資料庫(1, i, 3) = 12 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         人物異常狀態資料庫(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 2) = 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 1, i, 12, app_path & "gif\異常狀態\movdown.gif", 1, 3
                         異常狀態檢查數(12, 1) = 1
                         異常狀態檢查數(12, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
   End Select
End If
End Sub
Sub 南瓜王_重壓()
If FormMainMode.comaiatk(1).Caption = "重壓" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(7, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "南瓜王" Then
   Select Case atkingckai(7, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 3 And atkingckai(7, 2) = 0 Then
               atkingckai(7, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\南瓜王\南瓜王_重壓_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1080
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 28
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(7, 1) = 3
       Case 3
            atkingckai(7, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
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
            End If
   End Select
End If
End Sub
Sub 南瓜王_超再生()
If FormMainMode.comaiatk(2).Caption = "超再生" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(6, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "南瓜王" Then
   Select Case atkingckai(6, 1)
      Case 1
          If atkingpagetot(2, 3) >= 1 And atkingckai(6, 2) = 0 Then
               atkingckai(6, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
          ElseIf atkingpagetot(2, 3) < 1 And atkingckai(6, 2) = 1 Then
               atkingckai(6, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\怪物卡\南瓜王\南瓜王_超再生_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = -2040
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 6
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
            回復執行_電腦 3, 1
            atkingckai(6, 2) = 0
   End Select
End If
End Sub
Sub 雪莉_自殺傾向(ByVal Index As Integer)
Dim atkingtotai As Integer '特數量暫時統計變數
Dim a As Integer, i As Integer, j As Integer '暫時變數
If FormMainMode.comaiatk(1).Caption = "自殺傾向" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(1, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "雪莉" Then
 Select Case atkingckai(1, 1)
   Case 1
'      If livecom(角色人物對戰人數(2, 2)) <= 4 Then j = 54 Else j = 57 '視情況排除特3移2卡
'      If livecom(角色人物對戰人數(2, 2)) <= 4 Then
'          For i = 1 To 106
'                If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) <> 1 Then
'                   If pagecardnum(i, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
'                   If pagecardnum(i, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
'                ElseIf Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) <> 1 Then
'                   If pagecardnum(i, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
'                   If pagecardnum(i, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
'                End If
'           Next
'      Else
           For i = 1 To 106
               If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) <> 1 Then
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 3) = a4a Then
                        atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                    ElseIf pagecardnum(i, 1) = a4a Then
                        atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                    ElseIf pagecardnum(i, 3) = a4a Then
                        atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
                    End If
               End If
           Next
'      End If
      
      
      If atkingtotai < livecom(角色人物對戰人數(2, 2)) And atkingtotai > 1 Then
         For a = 1 To 106
             戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 a
         Next
      ElseIf atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then
         atkingtotai = 0
            If livecom(角色人物對戰人數(2, 2)) >= (livecom(角色人物對戰人數(2, 2)) \ 4) * 3 Then '如果血量大於4分之3的話，特3以上卡優先
                For i = 106 To 55 Step -1
                    If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) <> 1 Then
                       If Val(pagecardnum(i, 2)) >= 3 And pagecardnum(i, 1) = a4a Then
                            atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                            If atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then Exit For
                            戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 i
                       ElseIf Val(pagecardnum(i, 4)) >= 3 And pagecardnum(i, 3) = a4a Then
                            atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
                            If atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then Exit For
                            戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 i
                       End If
                    End If
                    
'                    If atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then Exit For
'
'                    戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 a
                Next
            End If
            If atkingtotai < livecom(角色人物對戰人數(2, 2)) Then
               a = 1
               Do While a <= 106
                  If livecom(角色人物對戰人數(2, 2)) <= 4 Then
                      If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 And (Val(pagecardnum(i, 2)) <> 3 And pagecardnum(i, 1) = a4a) Then
                            If pagecardnum(a, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                            If pagecardnum(a, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                            If atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then Exit Do
                            '===========================
                            戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 a
                      End If
                  Else
                      If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 Then
                            If pagecardnum(a, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                            If pagecardnum(a, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                            If atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then Exit Do
                            '===========================
                            戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 a
                      End If
                  End If
                    
'                    If atkingtotai >= livecom(角色人物對戰人數(2, 2)) Then Exit Do
'
'                    戰鬥系統類.comatk_AI_雪莉_自殺傾向_特 a
'                  End If
                  a = a + 1
               Loop
            End If
      End If
   Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(1, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(1, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 2)) * 5
               If atkingckai(1, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(1, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
    Case 3
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(1, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(1, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 4)) * 5
               If atkingckai(1, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(1, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
   Case 4
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_自殺傾向_2.jpg"
                atkingno(i, 2) = 2
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
       '-------------
    Case 5
        戰鬥系統類.傷害執行_技能直傷_電腦 Val(atkingpagetot(2, 4)), 1
        atkingckai(1, 2) = 0
   End Select
End If
End Sub
Sub 雪莉_異質者()
Dim atkingtotai As Integer '特數量暫時統計變數
Dim a As Integer, i As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "異質者" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(12, 2) = 1) _
    And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "雪莉" Then
 Select Case atkingckai(12, 1)
   Case 1
      atkingckai(12, 1) = 2
      For i = 55 To 106
         If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And ((Val(pagecardnum(i, 2)) = 3 And pagecardnum(i, 1) = a4a) Or (Val(pagecardnum(i, 4)) = 3 And pagecardnum(i, 3) = a4a)) Then
            atkingtotai = Val(atkingtotai) + 1
         End If
      Next
      If atkingtotai >= 1 Then
         Select Case livecom(角色人物對戰人數(2, 2))
            Case Is < 3
                If Val(FormMainMode.顯示列1.goi1) - Val(FormMainMode.顯示列1.goi2) >= livecom(角色人物對戰人數(2, 2)) Then
                    GoTo AI技能_雪莉_異質者_出牌階段二
                End If
            Case 3
                If Val(FormMainMode.顯示列1.goi1) - Val(FormMainMode.顯示列1.goi2) >= 9 Then
                    GoTo AI技能_雪莉_異質者_出牌階段二
                End If
            Case Is > 3
                If Int(Val(FormMainMode.顯示列1.goi1) / 3 + 0.9) - Int(Val(FormMainMode.顯示列1.goi2) / 3 + 0.9) >= livecom(角色人物對戰人數(2, 2)) Then
                    GoTo AI技能_雪莉_異質者_出牌階段二
                End If
         End Select
      End If
      '==========如果不符合任何條件時
      Exit Sub
    '================================
AI技能_雪莉_異質者_出牌階段二:
      For a = 55 To 106
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) = 3 Then
                    戰鬥系統類.comatk_AI_雪莉_多妮妲_異質者_特 a
                    Exit For
                ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) = 3 Then
                    戰鬥系統類.comatk_AI_雪莉_多妮妲_異質者_特 a
                    Exit For
                End If
             End If
      Next
    Case 2
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
'                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) >= 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
             If rrr >= 1 And atkingckai(12, 2) = 0 Then
                atkingckai(12, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
             If rrr < 1 And atkingckai(12, 2) = 1 Then
                atkingckai(12, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
   Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_異質者_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 10110
                atkingno(i, 7) = 12
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
          atkingckai(12, 2) = 0
          If Val(擲骰表單溝通暫時變數(2)) - Val(擲骰表單溝通暫時變數(3)) >= livecom(角色人物對戰人數(2, 2)) And 異常狀態檢查數(18, 2) = 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 6
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 1) = 6
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 6, 3
                          異常狀態檢查數(1, 1) = 1
                          異常狀態檢查數(1, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '==================================
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 18 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 1) = 0
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 18, app_path & "gif\異常狀態\不死.gif", 0, 3
                          異常狀態檢查數(18, 1) = 1
                          異常狀態檢查數(18, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '===============================
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 19 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 1) = 0
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 19, app_path & "gif\異常狀態\自壞.gif", 0, 3
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
Sub 雪莉_巨大黑犬()
Dim a As Integer, i As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "巨大黑犬" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(2, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "雪莉" Then
  Select Case atkingckai(2, 1)
   Case 1
       For i = 1 To 106
          If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) >= 3 And Val(pagecardnum(i, 5)) = 2 And movecp < 3 Then
             戰鬥系統類.comatk_AI_雪莉_巨大黑犬_劍 i
             Exit For
          ElseIf pagecardnum(i, 3) = a1a And Val(pagecardnum(i, 4)) >= 3 And Val(pagecardnum(i, 5)) = 2 And movecp < 3 Then
             戰鬥系統類.comatk_AI_雪莉_巨大黑犬_劍 i
             Exit For
          End If
       Next
       atkingckai(2, 1) = 2
    Case 2
          If movecp < 3 Then
            If atkingpagetot(2, 1) >= 3 And atkingckai(2, 2) = 0 Then
               atkingckai(2, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 1) < 3 And atkingckai(2, 2) = 1 Then
               atkingckai(2, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
   Case 3
       For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_巨大黑犬_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 9810
                atkingno(i, 6) = 8940
                atkingno(i, 7) = 2
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
        Do
           atkingckai(2, 2) = 0
           For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
             If 人物異常狀態資料庫(1, i, 3) = 11 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                 FormMainMode.personusspe(i).person_num = 4
                 FormMainMode.personusspe(i).person_turn = 3
                 人物異常狀態資料庫(1, i, 1) = 4
                 人物異常狀態資料庫(1, i, 2) = 3
                 Exit Do
             End If
           Next
           For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
              If 人物異常狀態資料庫(1, i, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 1, i, 11, app_path & "gif\異常狀態\defdown.gif", 4, 3
                 異常狀態檢查數(11, 1) = 1
                 異常狀態檢查數(11, 2) = 1
                 Exit Do
             End If
           Next
        Loop
  End Select
End If
End Sub
Sub 雪莉_飛刃雨()
Dim atkingtotai As Integer '特數量暫時統計變數
Dim ak As Integer, j As Integer, ui As Integer '暫時變數
If FormMainMode.comaiatk(4).Caption = "飛刃雨" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(5, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "雪莉" Then
 Select Case atkingckai(5, 1)
   Case 1
      If movecp = 3 Then
          For j = 49 To 54   '防1移1卡優先
              If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                   If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                      戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                      ak = 1
                      Exit For
                   ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                      戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                      ak = 1
                      Exit For
                   End If
              End If
          Next
          If ak = 0 Then
             For j = 39 To 44   '槍1移1卡其次優先
                If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                   If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                      戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                      ak = 1
                      Exit For
                   ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                      戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                      ak = 1
                      Exit For
                   End If
                End If
             Next
          End If
          If ak = 0 Then
             For j = 1 To 106
                If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                   If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                      戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                      ak = 1
                      Exit For
                   ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                      戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                      ak = 1
                      Exit For
                   End If
                End If
             Next
          End If
          If ak = 1 Then
             atkingckai(5, 2) = 1
    '         atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
       End If
   Case 2
      atkingckai(5, 1) = 3
      atkingckai(5, 2) = 0 '基於AI出牌公平判斷原則
      If moveturn = 2 Then
'        If livecom(角色人物對戰人數(2, 2)) <= 5 Then  '視情況撇除特3移2卡
'            ui = 54
'        Else
'            ui = 57
'        End If
        For j = 1 To 106
          If (livecom(角色人物對戰人數(2, 2)) <= 5 And Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 And ((Val(pagecardnum(j, 2)) <> 3 And pagecardnum(j, 1) = a4a) Or (Val(pagecardnum(j, 4)) <> 3 And pagecardnum(j, 3) = a4a))) Or _
              livecom(角色人物對戰人數(2, 2)) > 5 And Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
            pagecardnum(j, 11) = 1
            If pagecardnum(j, 1) = a4a And pagecardnum(j, 3) = a4a Then
                pagecardnum(j, 11) = 0
            ElseIf pagecardnum(j, 1) = a4a Then
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
            End If
         End If
       Next
    End If
   Case 3
          If atkingckai(5, 2) = 0 And movecp = 3 Then
             For i = 1 To 106
                  If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                     atkingckai(5, 2) = 1
                     atkingckai(5, 1) = 4
                     atkingtrn(2) = Val(atkingtrn(2)) + 1
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(FormMainMode.pagecomqlead) * 2
                     atking_sheri_4_tot_ai = Val(FormMainMode.pagecomqlead)
                     Exit For
                  End If
             Next
          End If
    Case 4
            If atkingpagetot(2, 3) = 0 Then
               atkingckai(5, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               atkingckai(5, 1) = 3
               If Val(FormMainMode.pagecomqlead) = atking_sheri_4_tot_ai Then
                  攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(FormMainMode.pagecomqlead) * 2
               Else
                  攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(FormMainMode.pagecomqlead) * 2 - 2
               End If
               atking_sheri_4_tot_ai = 0
            ElseIf atkingpagetot(2, 3) > 1 Then
               For i = 1 To 106
                 If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                    ttt = ttt + 1
                 End If
               Next
               If ttt = 0 Then
                 atkingckai(5, 2) = 0
                 atkingtrn(2) = Val(atkingtrn(2)) - 1
                 atkingckai(5, 1) = 3
                 If Val(FormMainMode.pagecomqlead) = atking_sheri_4_tot_ai Then
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(FormMainMode.pagecomqlead) * 2
                 Else
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(FormMainMode.pagecomqlead) * 2 - 2
                 End If
                 atking_sheri_4_tot_ai = 0
               End If
            End If
            If atkingckai(5, 2) = 1 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (Val(FormMainMode.pagecomqlead) - Val(atking_sheri_4_tot_ai)) * 2
               atking_sheri_4_tot_ai = Val(FormMainMode.pagecomqlead)
            End If
   Case 5
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\雪莉\雪莉_飛刃雨_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 9690
                atkingno(i, 7) = 22
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
        atkingckai(5, 2) = 0
   End Select
End If
End Sub
Sub 蕾_輪旋曲_琉璃色的微風()
If FormMainMode.comaiatk(1).Caption = "輪旋曲-琉璃色的微風" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(4, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(4, 1)
      Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 4 And atkingckai(4, 2) = 0 Then
               atkingckai(4, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 3
            ElseIf atkingpagetot(2, 1) < 4 And atkingckai(4, 2) = 1 Then
               atkingckai(4, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 3
            End If
          End If
      Case 2
             atkingckai(4, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-輪旋曲-琉璃色的微風_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7485
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 20
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '================
             戰鬥系統類.直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) - 3
   End Select
End If
End Sub
Sub 蕾_EX_輪旋曲_琉璃色的微風()
If FormMainMode.comaiatk(1).Caption = "Ex輪旋曲-琉璃色的微風" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(13, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(13, 1)
      Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 5 And atkingckai(13, 2) = 0 Then
               atkingckai(13, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 8
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 6
             ElseIf atkingpagetot(2, 1) < 5 And atkingckai(13, 2) = 1 Then
               atkingckai(13, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 8
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 6
             End If
          End If
      Case 2
             atkingckai(13, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-輪旋曲-琉璃色的微風2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7320
                   atkingno(i, 6) = 9000
                   atkingno(i, 7) = 41
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '================
             戰鬥系統類.直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) - 6
   End Select
End If
End Sub
Sub 蕾_EX_協奏曲_加百烈的守護()
If FormMainMode.comaiatk(2).Caption = "Ex協奏曲-加百烈的守護" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(58, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(58, 1)
        Case 1
            If atkingpagetot(2, 4) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(58, 2) = 0 Then
'            If atkingpagetot(2, 3) >= 1 And atkingckai(58, 2) = 0 Then
               atkingckai(58, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
            ElseIf (atkingpagetot(2, 4) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(58, 2) = 1 Then
'            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(58, 2) = 1 Then
               atkingckai(58, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
'               攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-協奏曲-加百烈的守護_2.jpg"
                   atkingno(i, 2) = 2
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
          atking_AI_蕾_守護模式狀態啟動值 = True
    Case 3
          atking_AI_蕾_守護模式狀態啟動值 = False
          atkingckai(58, 2) = 0
    Case 4
          擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) - 5
          擲骰後骰傷害數 = 擲骰後骰傷害數 - 5
   End Select
End If
End Sub
Sub 蕾_EX_安魂曲_死神的鎮魂歌()
Dim rrr As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "Ex安魂曲-死神的鎮魂歌" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(63, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(63, 1)
        Case 1
            For i = 1 To 106
               If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                  rrr = rrr + 1
               End If
            Next
          If rrr >= 1 And atkingckai(63, 2) = 0 Then
             atkingckai(63, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 1 And atkingckai(63, 2) = 1 Then
             atkingckai(63, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-安魂曲-死神的鎮魂歌_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(63, 2) = 0
             If livecom(角色人物對戰人數(2, 2)) <= 0 Then
                 For i = 2 To 3
                     If livecom(角色待機人物紀錄數(2, i)) > 0 Then
                        Do
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                                  If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 9
                                      FormMainMode.personcomspe(j).person_turn = 2
                                      人物異常狀態資料庫(2, j, 1) = 9
                                      人物異常狀態資料庫(2, j, 2) = 2
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                               If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 9, 2
                                  異常狀態檢查數(1, 1) = 1
                                  異常狀態檢查數(1, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        Do
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                                  If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 9
                                      FormMainMode.personcomspe(j).person_turn = 2
                                      人物異常狀態資料庫(2, j, 1) = 9
                                      人物異常狀態資料庫(2, j, 2) = 2
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                               If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, j, 2, app_path & "gif\異常狀態\defup.gif", 9, 2
                                  異常狀態檢查數(2, 1) = 1
                                  異常狀態檢查數(2, 2) = 1
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
Dim atkingtotai As Integer '特數量暫時統計變數
Dim pagene(1 To 106) As Integer '選擇牌暫時變數
Dim a As Integer, i As Integer '暫時變數
Dim k As String '暫時變數
Dim num(1 To 2) As Integer
If FormMainMode.comaiatk(4).Caption = "終曲-無盡輪迴的終結" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(11, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
 Select Case atkingckai(11, 1)
   Case 1
      atkingckai(11, 1) = 2
       If movecp < 3 Then
            For i = 1 To 106
               If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                 If pagecardnum(i, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                 If pagecardnum(i, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
               End If
            Next
            
            If atkingtotai >= 4 Then
               atkingtotai = 0
               Select Case movecp
                      Case 1
                         k = a1a
                      Case Is > 1
                         k = a5a
                End Select
               '====================1階段-先選擇第一張牌
               Do
                    '===========(非與當下階段對應)
                    For a = 106 To 1 Step -1
                       If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                           If pagecardnum(a, 1) = a4a And pagecardnum(a, 3) <> k And pagene(a) = 0 Then
                               atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                               pagene(a) = 1
                               Exit Do
                           ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 1) <> k And pagene(a) = 0 Then
                               atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                               pagene(a) = 1
                               Exit Do
                           End If
                        End If
                    Next
                    If atkingtotai = 0 Then
                        '===========(選擇所有)
                        For a = 106 To 1 Step -1
                           If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                               If pagecardnum(a, 1) = a4a And pagene(a) = 0 Then
                                   atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                                   pagene(a) = 1
                                   Exit Do
                               ElseIf pagecardnum(a, 3) = a4a And pagene(a) = 0 Then
                                   atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                                   pagene(a) = 1
                                   Exit Do
                               End If
                            End If
                        Next
                    End If
               Loop
               If atkingtotai < 4 Then
                   '==============2階段-依剩下特數值選非與當下階段對應槍/劍的第2張牌(不限特數值)
                  For a = 106 To 1 Step -1
                     If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                         If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) >= 4 - atkingtotai And pagecardnum(a, 3) <> k And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                             pagene(a) = 1
                             Exit For
                         ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) >= 4 - atkingtotai And pagecardnum(a, 1) <> k And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                             pagene(a) = 1
                             Exit For
                         End If
                      End If
                      If atkingtotai >= 4 Then Exit For
                  Next
              End If
              If atkingtotai < 4 Then
                   '====================3階段-依剩下特數值選廣域的第2張牌
                  For a = 106 To 1 Step -1
                     If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                         If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) = 4 - atkingtotai And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                             pagene(a) = 1
                             Exit For
                         ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) = 4 - atkingtotai And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                             pagene(a) = 1
                             Exit For
                         End If
                      End If
                      If atkingtotai >= 4 Then Exit For
                  Next
              End If
              If atkingtotai < 4 Then
                 '====================4階段-選所有剩下的牌
                 For a = 106 To 1 Step -1
                     If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                         If pagecardnum(a, 1) = a4a And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                             pagene(a) = 1
                         ElseIf pagecardnum(a, 3) = a4a And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                             pagene(a) = 1
                         End If
                      End If
                      If atkingtotai >= 4 Then Exit For
                  Next
              End If
           End If
           If atkingtotai >= 4 Then
               '===========進行實際出牌程序
               For a = 1 To 106
                   If pagene(a) = 1 Then
                       戰鬥系統類.comatk_AI_蕾_終曲_無盡輪迴的終結_特 a
                   End If
               Next
           End If
       End If
   Case 2
          If movecp < 3 Then
            If atkingpagetot(2, 4) >= 4 And atkingckai(11, 2) = 0 Then
               atkingckai(11, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 16
            ElseIf atkingpagetot(2, 4) < 4 And atkingckai(11, 2) = 1 Then
               atkingckai(11, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 16
            End If
          End If
   Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\蕾\蕾-終曲-無盡輪迴的終結_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 8655
                atkingno(i, 6) = 0
                atkingno(i, 7) = 39
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
             atkingckai(11, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                 num(1) = 1
                 num(2) = FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption
                 For i = 2 To 3
                    If FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption > 0 And FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption < num(2) Then
                        num(1) = i
                        num(2) = FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption
                    End If
                Next
                傷害執行_技能直傷_使用者 Val(擲骰表單溝通暫時變數(2)), num(1)
            End If
            擲骰表單溝通暫時變數(2) = 0
            擲骰後骰傷害數 = 0
   End Select
End If
End Sub
Sub 艾伯李斯特_精密射擊()
If FormMainMode.comaiatk(1).Caption = "精密射擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(19, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾伯李斯特" Then
   Select Case atkingckai(19, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 2 And atkingckai(19, 2) = 0 Then
                   atkingckai(19, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
                End If
                If atkingpagetot(2, 5) < 2 And atkingckai(19, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
                   atkingckai(19, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(19, 2) = 0
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
If FormMainMode.comaiatk(2).Caption = "雷擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(66, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾伯李斯特" Then
   Select Case atkingckai(66, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 2 And atkingckai(66, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
                   atkingckai(66, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 2 And atkingckai(66, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
                   atkingckai(66, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking2_2.jpg"
                   atkingno(i, 2) = 2
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
             If Val(擲骰表單溝通暫時變數(2)) > 0 And Val(FormMainMode.pageusglead.Caption) > 0 Then
                 atking_AI_艾伯李斯特_雷擊紀錄數(1) = Val(擲骰表單溝通暫時變數(2))
                 atking_AI_艾伯李斯特_雷擊紀錄數(2) = 1
                 '==========================
                  Do Until atking_AI_艾伯李斯特_雷擊紀錄數(2) > atking_AI_艾伯李斯特_雷擊紀錄數(1) Or Val(FormMainMode.pageusglead.Caption) <= 0
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                            目前數(21) = 4
                            目前數(20) = m
                            atking_AI_艾伯李斯特_雷擊紀錄數(2) = atking_AI_艾伯李斯特_雷擊紀錄數(2) + 1
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                   Loop
             Else
                 atkingckai(66, 1) = 5
                 FormMainMode.骰子執行完啟動.Enabled = True
             End If
        Case 4
             Do Until atking_AI_艾伯李斯特_雷擊紀錄數(2) > atking_AI_艾伯李斯特_雷擊紀錄數(1) Or Val(FormMainMode.pageusglead.Caption) <= 0
                 Randomize
                 m = Int(Rnd() * 106) + 1
                 If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                     目前數(21) = 4
                     目前數(20) = m
                     atking_AI_艾伯李斯特_雷擊紀錄數(2) = atking_AI_艾伯李斯特_雷擊紀錄數(2) + 1
                     FormMainMode.tr使用者_棄牌.Enabled = True
                     Exit Sub
                 End If
            Loop
            If atking_AI_艾伯李斯特_雷擊紀錄數(2) > atking_AI_艾伯李斯特_雷擊紀錄數(1) Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
                atkingckai(66, 1) = 5
                目前數(24) = 22
                FormMainMode.等待時間_2.Enabled = True
            End If
        Case 5
            atkingckai(66, 2) = 0
            Erase atking_AI_艾伯李斯特_雷擊紀錄數
   End Select
End If
End Sub
Sub 艾伯李斯特_茨林()
If FormMainMode.comaiatk(3).Caption = "茨林" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(67, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾伯李斯特" Then
   Select Case atkingckai(67, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 2 And atkingpagetot(2, 2) >= 2 And atkingckai(67, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 7
                   atkingckai(67, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 4) < 2 Or atkingpagetot(2, 2) < 2) And atkingckai(67, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 7
                   atkingckai(67, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1200
                   atkingno(i, 5) = 6705
                   atkingno(i, 6) = 10245
                   atkingno(i, 7) = 67
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾伯李斯特\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(FormMainMode.顯示列1.goi1) <= 0 Then
                atkingckai(67, 2) = 0
            End If
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) < 0 Then
                戰鬥系統類.傷害執行_技能直傷_使用者 Abs(擲骰表單溝通暫時變數(2)), 1
            End If
            atkingckai(67, 2) = 0
   End Select
End If
End Sub
Sub 艾伯李斯特_智略()
If FormMainMode.comaiatk(4).Caption = "智略" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(68, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾伯李斯特" Then
   Select Case atkingckai(68, 1)
      Case 1
            If pageqlead(2) >= 3 And atkingckai(68, 2) = 0 Then
               atkingckai(68, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If pageqlead(2) < 3 And atkingckai(68, 2) = 1 Then
               atkingckai(68, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾伯李斯特\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6060
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 68
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(FormMainMode.pageul.Caption) < 2 And atking_AI_艾伯李斯特_智略紀錄數 = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_AI_艾伯李斯特_智略紀錄數 = atking_AI_艾伯李斯特_智略紀錄數 + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_艾伯李斯特_智略紀錄數 > 2
                    目前數(15) = 23
                    FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_艾伯李斯特_智略紀錄數 > 2 Or Val(FormMainMode.pageul.Caption) <= 0 Then
               atking_AI_艾伯李斯特_智略紀錄數 = 0
               atkingckai(68, 2) = 0
            End If
   End Select
End If
End Sub
Sub 史塔夏_殺戮器官()
If FormMainMode.comaiatk(1).Caption = "殺戮器官" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(88, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "史塔夏" Then
   Select Case atkingckai(88, 1)
        Case 1
            If pageqlead(2) >= 3 And atkingckai(88, 2) = 0 Then
               atkingckai(88, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf pageqlead(2) < 3 And atkingckai(88, 2) = 1 Then
               atkingckai(88, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_殺戮器官_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -360
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 88
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciamini2.png"
            FormMainMode.personcomminijpg.小人物影子圖片 = app_path & "gif\史塔夏\殺戮\Staciaminidown2.png"
            FormMainMode.personcomminijpg.小人物影子Left = 90
            FormMainMode.personcomminijpg.小人物影子top差 = -60
            Form6.jpgcom.大人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaperson2.png"
            FormMainMode.顯示列1.電腦方小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaf2.png"
            atking_AI_史塔夏_殺戮模式狀態數(2) = 1
            atkingckai(88, 2) = 0
'            formsettingpersonus.smalldownleft = -90
'            formsettingpersonus.smalldowntop = -60
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personcomminijpg.Visible = True
   End Select
End If
End Sub

Sub 史塔夏_愚者之手()
Dim apn As Integer
If FormMainMode.comaiatk(2).Caption = "愚者之手" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(20, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "史塔夏" Then
   Select Case atkingckai(20, 1)
        Case 1
            If movecp < 3 Then
             For i = 1 To 3
                 If liveus(i) > 0 Then
                     apn = apn + 1
                 End If
             Next
             If atkingpagetot(2, 1) >= 6 And atkingckai(20, 2) = 0 Then
               atkingckai(20, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + apn * 4
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(20, 2) = 1 Then
               atkingckai(20, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - apn * 4
            End If
          End If
        Case 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_愚者之手_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 76
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(20, 1) = 3
        Case 3
            For i = 1 To 3
                 If liveus(i) > 0 Then
                     apn = apn + 1
                 End If
            Next
            If atking_AI_史塔夏_殺戮模式狀態數(2) = 1 Then
                戰鬥系統類.傷害執行_技能直傷_電腦 apn, 1
            End If
            atkingckai(20, 2) = 0
   End Select
End If
End Sub
Sub 史塔夏_時間種子()
Dim bloodtot As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "時間種子" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(55, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "史塔夏" Then
   Select Case atkingckai(55, 1)
        Case 1
            If movecp < 3 Then
             If atkingpagetot(2, 2) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(55, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(55, 2) = 0 Then
               atkingckai(55, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 2) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(55, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(55, 2) = 1 Then
               atkingckai(55, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_時間種子_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 55
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
            If Val(livecom(角色人物對戰人數(2, 2))) < Val(livecommax(角色人物對戰人數(2, 2))) Then
               Select Case atking_AI_史塔夏_殺戮模式狀態數(2)
                   Case 0
                        回復執行_電腦 bloodtot, 1
                   Case 1
                        回復執行_電腦 bloodtot \ 2, 1
                  End Select
            End If
            atkingckai(55, 2) = 0
   End Select
End If
End Sub

Sub 史塔夏_命運的鐵門()
Dim num(1 To 2, 1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.comaiatk(4).Caption = "命運的鐵門" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(21, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "史塔夏" Then
   Select Case atkingckai(21, 1)
        Case 1
         If movecp = 3 Then
             If atkingpagetot(2, 1) >= 9 And atkingckai(21, 2) = 0 Then
               atkingckai(21, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(2, 1) < 9 And atkingckai(21, 2) = 1 Then
               atkingckai(21, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
         End If
        Case 2
             戰鬥系統類.自動捲軸捲動
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\史塔夏\史塔夏_命運的鐵門_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6720
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 21
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(livecom(角色人物對戰人數(2, 2))) < Val(livecommax(角色人物對戰人數(2, 2))) Then
               If atking_AI_史塔夏_殺戮模式狀態數(2) = 1 Then
                  回復執行_電腦 3, 1
               End If
            End If
            atkingckai(21, 1) = 4
        Case 4
           num(1, 2) = 999 '目的取最低HP數
           num(2, 2) = 999
           For i = 2 To 3
               If livecom(角色待機人物紀錄數(2, i)) < num(2, 2) And livecom(角色待機人物紀錄數(2, i)) > 0 Then
                   num(2, 1) = i
                   num(2, 2) = livecom(角色待機人物紀錄數(2, i))
               End If
            Next
            For i = 1 To 3
               If Val(FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption) < num(1, 2) And Val(FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption) > 0 Then
                   num(1, 1) = i
                   num(1, 2) = FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption
               End If
           Next
           If num(2, 2) < num(1, 2) Or num(1, 2) = num(2, 2) Then
               戰鬥系統類.傷害執行_立即死亡_電腦 num(2, 1)
           Else
               戰鬥系統類.傷害執行_立即死亡_使用者 num(1, 1)
           End If
           atkingckai(21, 2) = 0
   End Select
End If
End Sub
Sub 阿貝爾_霸王閃擊()
If FormMainMode.comaiatk(1).Caption = "霸王閃擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(22, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿貝爾" Then
   Select Case atkingckai(22, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingckai(22, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
                   atkingckai(22, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 1) < 3 And atkingckai(22, 2) = 1 Then
                   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
                   atkingckai(22, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(22, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
If FormMainMode.comaiatk(2).Caption = "閃電旋風刺" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(71, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿貝爾" Then
   Select Case atkingckai(71, 1)
      Case 1
           If movecp = 2 Then
                If atkingpagetot(2, 3) >= 1 And atkingckai(71, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
                   atkingckai(71, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 3) < 1 And atkingckai(71, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
                   atkingckai(71, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6555
                   atkingno(i, 6) = 8625
                   atkingno(i, 7) = 71
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(71, 2) = 0
             If movecp > 1 Then
                 戰鬥系統類.執行動作_距離變更 movecp - 1
             End If
   End Select
End If
End Sub
Sub 阿貝爾_幻影劍舞()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "幻影劍舞" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(23, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿貝爾" Then
   Select Case atkingckai(23, 1)
      Case 1
            If movecp = 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
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
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(23, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(23, 2) = 0 Then
                atkingckai(23, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 9
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(23, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(23, 2) = 1 Then
                atkingckai(23, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 9
              End If
      Case 2
             atkingckai(23, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8520
                   atkingno(i, 6) = 8280
                   atkingno(i, 7) = 99
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\阿貝爾\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 布勞_時間爆彈()
Dim tn As Integer
If FormMainMode.comaiatk(3).Caption = "時間爆彈" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(24, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "布勞" Then
   Select Case atkingckai(24, 1)
        Case 1
             If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(24, 2) = 0 Then
               atkingckai(24, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 7
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(24, 2) = 1 Then
               atkingckai(24, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 7
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7125
                   atkingno(i, 6) = 9330
                   atkingno(i, 7) = 24
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\布勞\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(24, 2) = 0
            tn = Val(FormMainMode.turni)
            If tn = 2 Or tn = 3 Or tn = 5 Or tn = 7 Or tn = 11 Or tn = 13 Or tn = 17 Then
               戰鬥系統類.傷害執行_技能直傷_使用者 3, 1
            End If
   End Select
End If
End Sub
Sub 布勞_時間追獵()
If FormMainMode.comaiatk(2).Caption = "時間追獵" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(70, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "布勞" Then
   Select Case atkingckai(70, 1)
        Case 1
             If movecp < 3 Then
                 If atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(70, 2) = 0 Then
                   atkingckai(70, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(70, 2) = 1 Then
                   atkingckai(70, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking2_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(70, 2) = 0
            戰鬥系統類.直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) - Val(FormMainMode.turni)
'            攻擊防禦骰子總數(1) = FormMainMode.顯示列1.goi1
   End Select
End If
End Sub

Sub 艾依查庫_連射()
If FormMainMode.comaiatk(1).Caption = "連射" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(25, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾依查庫" Then
   Select Case atkingckai(25, 1)
      Case 1
           If movecp > 1 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          End If
          If rrr >= 2 And atkingckai(25, 2) = 0 Then
             攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
             atkingckai(25, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 2 And atkingckai(25, 2) = 1 Then
             攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
             atkingckai(25, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
      Case 2
             atkingckai(25, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6780
                   atkingno(i, 6) = 10185
                   atkingno(i, 7) = 25
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾依查庫\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 艾依查庫_憤怒一擊()
Dim ape As Integer
If FormMainMode.comaiatk(3).Caption = "憤怒一擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(69, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾依查庫" Then
   Select Case atkingckai(69, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(69, 2) = 0 Then
                   ape = (livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2))) * 2
                   If ape > 16 Then ape = 16
                   atkingckai(69, 2) = 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + ape
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(69, 2) = 1 Then
                   ape = (livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2))) * 2
                   If ape > 16 Then ape = 16
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - ape
                   atkingckai(69, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(69, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6615
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾依查庫\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub

Sub 艾依查庫_神速之劍(ByVal Index As Integer)
Dim aw As Integer
If FormMainMode.comaiatk(2).Caption = "神速之劍" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(26, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾依查庫" Then
   Select Case atkingckai(26, 1)
      Case 1
             If movecp > 1 Then
                 If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 1) >= 2 And atkingckai(26, 2) = 0 Then
                     aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_艾依查庫_神速之劍計算數值紀錄數(1))
                     atking_AI_艾依查庫_神速之劍計算數值紀錄數(1) = aw
                     atkingckai(26, 2) = 1
                     atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
            End If
      Case 2
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 2 And atkingckai(26, 2) = 1 Then
                   aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_艾依查庫_神速之劍計算數值紀錄數(1))
                   atking_AI_艾依查庫_神速之劍計算數值紀錄數(1) = aw
            End If
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 2 And atkingckai(26, 2) = 1 Then
                   If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 1) >= 2 And atkingckai(26, 2) = 1 Then
                        aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - (atking_AI_艾依查庫_神速之劍計算數值紀錄數(1) - aw)
                        atking_AI_艾依查庫_神速之劍計算數值紀錄數(1) = aw
                   ElseIf (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 1) < 2) And atkingckai(26, 2) = 1 Then
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - atking_AI_艾依查庫_神速之劍計算數值紀錄數(1)
                        atkingckai(26, 2) = 0
                        atkingckai(26, 1) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) - 1
                        Erase atking_AI_艾依查庫_神速之劍計算數值紀錄數
                    End If
            End If
'            formmainmode.trgoi2.Enabled = True
    Case 3
        If Val(pagecardnum(Index, 5)) = 2 And atkingckai(26, 2) = 1 Then
               If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 1) >= 2 Then
                    aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_艾依查庫_神速之劍計算數值紀錄數(1))
                    atking_AI_艾依查庫_神速之劍計算數值紀錄數(1) = aw
               ElseIf (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 1) < 2) Then
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - atking_AI_艾依查庫_神速之劍計算數值紀錄數(1)
                    atkingckai(26, 2) = 0
                    atkingckai(26, 1) = 1
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
                    Erase atking_AI_艾依查庫_神速之劍計算數值紀錄數
                End If
        End If
'        formmainmode.trgoi2.Enabled = True
      Case 4
             atkingckai(26, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking2-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 9165
                   atkingno(i, 6) = 10350
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾依查庫\atking2-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             Erase atking_AI_艾依查庫_神速之劍計算數值紀錄數
   End Select
End If
End Sub
Sub 艾依查庫_不屈之心()
If FormMainMode.comaiatk(4).Caption = "不屈之心" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(27, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾依查庫" Then
   Select Case atkingckai(27, 1)
      Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          If rrr >= 2 And atkingckai(27, 2) = 0 Then
'          If pageqlead(2) >= 1 And atkingckai(27, 2) = 0 Then
             atkingckai(27, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 2 And atkingckai(27, 2) = 1 Then
'          If pageqlead(2) < 1 And atkingckai(27, 2) = 1 Then
             atkingckai(27, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾依查庫\atking4_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(27, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) >= livecom(角色人物對戰人數(2, 2)) Then
                 擲骰表單溝通暫時變數(2) = livecom(角色人物對戰人數(2, 2)) - 1
                 擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
             End If
   End Select
End If
End Sub
Sub 音音夢_愉快抽血(ByVal Index As Integer)
Dim n(1 To 2) As Integer
If FormMainMode.comaiatk(3).Caption = "愉快抽血" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(111, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "音音夢" Then
 Select Case atkingckai(111, 1)
    Case 1
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(111, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(111, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 2)) * 5
               If atkingckai(111, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(111, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
    Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(111, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(111, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 4)) * 5
               If atkingckai(111, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(111, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
    Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\音音夢\atking3_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6645
                atkingno(i, 6) = 9555
                atkingno(i, 7) = 111
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
       '-------------
    Case 4
       atkingckai(111, 2) = 0
        n(1) = 999 '取最小HP值
        n(2) = 0
        For i = 2 To 3
            If livecom(角色待機人物紀錄數(2, i)) > 0 And livecom(角色待機人物紀錄數(2, i)) < n(1) Then
                n(1) = livecom(角色待機人物紀錄數(2, i))
                n(2) = i
            End If
        Next
        If n(2) > 0 Then
            戰鬥系統類.傷害執行_技能直傷_電腦 Val(atkingpagetot(2, 4)), n(2)
        Else
            戰鬥系統類.傷害執行_技能直傷_電腦 Val(atkingpagetot(2, 4)), 1
        End If
  End Select
End If
End Sub
Sub 音音夢_溫柔注射()
Dim n(1 To 2) As Integer
If FormMainMode.comaiatk(2).Caption = "溫柔注射" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(28, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "音音夢" Then
   Select Case atkingckai(28, 1)
        Case 1
            If movecp < 3 Then
             If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 2) >= 2 And atkingckai(28, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(28, 2) = 0 Then
               atkingckai(28, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
            ElseIf (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 2) < 2) And atkingckai(28, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(28, 2) = 1 Then
               atkingckai(28, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\音音夢\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6165
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 28
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(28, 2) = 0
            '=======================
            n(1) = 999 '取最小HP值
            n(2) = 0
            For i = 2 To 3
                If livecom(角色待機人物紀錄數(2, i)) > 0 And livecom(角色待機人物紀錄數(2, i)) < n(1) Then
                    n(1) = livecom(角色待機人物紀錄數(2, i))
                    n(2) = i
                End If
            Next
            If n(2) > 0 Then
                If livecom(角色人物對戰人數(2, 2)) >= n(1) Then
                    戰鬥系統類.回復執行_電腦 livecom(角色人物對戰人數(2, 2)) - n(1), n(2)
                Else
                    戰鬥系統類.傷害執行_技能直傷_電腦 n(1) - livecom(角色人物對戰人數(2, 2)), n(2)
                End If
            End If
   End Select
End If
End Sub
Sub 音音夢_美味牛奶()
If FormMainMode.comaiatk(1).Caption = "美味牛奶" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(29, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "音音夢" Then
   Select Case atkingckai(29, 1)
        Case 1
            If pageqlead(2) >= 2 And atkingckai(29, 2) = 0 Then
               atkingckai(29, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf pageqlead(2) < 2 And atkingckai(29, 2) = 1 Then
               atkingckai(29, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\音音夢\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6360
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 29
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\音音夢\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(29, 2) = 0
            '==============================
            For k = 2 To 3
                傷害執行_技能直傷_電腦 1, k
            Next
            '==============================
            atking_AI_音音夢_成長模式狀態數(2) = 1
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = app_path & "gif\音音夢\成長\Nenemmini2.png"
            FormMainMode.personcomminijpg.小人物影子圖片 = app_path & "gif\音音夢\成長\Nenemminidown2.png"
            FormMainMode.personcomminijpg.小人物影子Left = 20
            FormMainMode.personcomminijpg.小人物影子top差 = -90
            Form6.jpgcom.大人物圖片 = app_path & "gif\音音夢\成長\Nenemperson2.png"
            FormMainMode.顯示列1.電腦方小人物圖片 = app_path & "gif\音音夢\成長\Nenemf2.png"
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personcomminijpg.Visible = True
   End Select
End If
End Sub
Sub 音音夢_秘密苦藥()
If FormMainMode.comaiatk(4).Caption = "秘密苦藥" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(112, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "音音夢" Then
   Select Case atkingckai(112, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(112, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(112, 2) = 0 Then
               atkingckai(112, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(112, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(112, 2) = 1 Then
               atkingckai(112, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\音音夢\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6750
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 112
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(112, 2) = 0
            '=======================
            For i = 2 To 3
                戰鬥系統類.回復執行_電腦 10, i
            Next
            戰鬥系統類.傷害執行_立即死亡_電腦 1
            '=======================
            If atking_AI_音音夢_成長模式狀態數(2) = 1 Then
                牌總階段數(2) = 牌總階段數(2) + 1
            End If
   End Select
End If
End Sub
Sub 梅倫_High_hand()
If FormMainMode.comaiatk(1).Caption = "High hand" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(64, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅倫" Then
   Select Case atkingckai(64, 1)
        Case 1
             If atkingpagetot(2, 4) >= 2 And atkingckai(64, 2) = 0 Then
               atkingckai(64, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + pageqlead(1) * 2
            ElseIf atkingpagetot(2, 4) < 2 And atkingckai(64, 2) = 1 Then
               atkingckai(64, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - pageqlead(1) * 2
            End If
        Case 2
             atkingckai(64, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\High hand_2.jpg"
                   atkingno(i, 2) = 2
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
Sub 梅倫_Lowball()
If FormMainMode.comaiatk(3).Caption = "Lowball" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(65, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅倫" Then
   Select Case atkingckai(65, 1)
        Case 1
             If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 _
                And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 5) >= 1 And atkingckai(65, 2) = 0 Then
'            If atkingpagetot(2, 1) >= 1 And atkingckai(65, 2) = 0 Then
                    atkingckai(65, 2) = 1
                    atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1 _
               Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 5) < 1) And atkingckai(65, 2) = 1 Then
'            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(65, 2) = 1 Then
                    atkingckai(65, 2) = 0
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\Lowball_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(65, 2) = 0
   End Select
End If
End Sub
Sub 梅倫_Gamble()
If FormMainMode.comaiatk(4).Caption = "Gamble" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(30, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅倫" Then
   Select Case atkingckai(30, 1)
        Case 1
            If movecp = 1 Then
                 If pageqlead(2) >= 3 And atkingckai(30, 2) = 0 Then
                   atkingckai(30, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf pageqlead(2) < 3 And atkingckai(30, 2) = 1 Then
                   atkingckai(30, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\Gamble_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6960
                   atkingno(i, 6) = 9780
                   atkingno(i, 7) = 106
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(30, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) = 1 Then
                 戰鬥系統類.傷害執行_立即死亡_使用者 1
             End If
   End Select
End If
End Sub
Sub 梅倫_Jackpot()
Dim m As Integer
If FormMainMode.comaiatk(2).Caption = "Jackpot" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(31, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅倫" Then
   Select Case atkingckai(31, 1)
        Case 1
            If movecp = 2 Then
                If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(31, 2) = 0 Then
'                If atkingpagetot(2, 2) >= 1 And atkingckai(31, 2) = 0 Then
                   atkingckai(31, 2) = 1
                ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(31, 2) = 1 Then
'                ElseIf atkingpagetot(2, 2) < 1 And atkingckai(31, 2) = 1 Then
                   atkingckai(31, 2) = 0
                End If
            End If
        Case 2
             atking_AI_梅倫_Jackpot紀錄數(1) = pageqlead(2) * 2
             atking_AI_梅倫_Jackpot紀錄數(2) = 1
        Case 3
             atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 4
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅倫\Jackpot_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6480
                   atkingno(i, 6) = 10020
                   atkingno(i, 7) = 31
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
             If Val(FormMainMode.pageul.Caption) < atking_AI_梅倫_Jackpot紀錄數(1) And atking_AI_梅倫_Jackpot紀錄數(2) = 1 Then
               戰鬥系統類.執行動作_洗牌
             End If
             If atking_AI_梅倫_Jackpot紀錄數(2) > atking_AI_梅倫_Jackpot紀錄數(1) Or Val(FormMainMode.pageul.Caption) <= 0 Then
                 atkingckai(31, 2) = 0
                 戰鬥系統類.執行動作_技能手動結束
            Else
                目前數(15) = 22
                FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                atking_AI_梅倫_Jackpot紀錄數(2) = atking_AI_梅倫_Jackpot紀錄數(2) + 1
            End If
   End Select
End If
End Sub
Sub 羅莎琳_染血之刃()
If FormMainMode.comaiatk(2).Caption = "染血之刃" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(32, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羅莎琳" Then
   Select Case atkingckai(32, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(32, 2) = 0 Then
               atkingckai(32, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(32, 2) = 1 Then
               atkingckai(32, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_染血之刃_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 32
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            回復執行_電腦 1, 1
        Case 4
            atkingckai(32, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                回復執行_電腦 1, 1
            End If
   End Select
End If
End Sub
Sub 羅莎琳_黑霧的纏繞()
Dim m As Integer '暫時變數
If FormMainMode.comaiatk(4).Caption = "黑霧的纏繞" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(59, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羅莎琳" Then
   Select Case atkingckai(59, 1)
        Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(59, 2) = 0 Then
               atkingckai(59, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
            ElseIf atkingpagetot(2, 4) < 2 And atkingckai(59, 2) = 1 Then
               atkingckai(59, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_黑霧的纏繞_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(59, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                       Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 20 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  人物異常狀態資料庫(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 20, app_path & "gif\異常狀態\damage.gif", 0, 2
                                  異常狀態檢查數(20, 1) = 1
                                  異常狀態檢查數(20, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 16 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  人物異常狀態資料庫(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 16, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                                  異常狀態檢查數(16, 1) = 1
                                  異常狀態檢查數(16, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 22 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  人物異常狀態資料庫(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 22, app_path & "gif\異常狀態\atkingerr.gif", 0, 2
                                  異常狀態檢查數(22, 1) = 1
                                  異常狀態檢查數(22, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
            End If
   End Select
End If
End Sub
Sub 羅莎琳_咀咒的刻印()
If FormMainMode.comaiatk(3).Caption = "咀咒的刻印" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(60, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羅莎琳" Then
   Select Case atkingckai(60, 1)
        Case 1
            If movecp > 1 Then
                If atkingpagetot(2, 2) >= 5 And atkingpagetot(2, 4) >= 1 And atkingckai(60, 2) = 0 Then
    '             If atkingpagetot(2, 2) >= 1 And atkingck(24, 2) = 0 Then
                   atkingckai(60, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 2) < 5 Or atkingpagetot(2, 4) < 1) And atkingckai(60, 2) = 1 Then
    '            ElseIf atkingpagetot(2, 2) < 1 And atkingck(24, 2) = 1 Then
                   atkingckai(60, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             atkingckai(60, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_咀咒的刻印_2.jpg"
                   atkingno(i, 2) = 2
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
             If atkingpagetot(1, 4) >= 1 Then
                 戰鬥系統類.直接寫入顯示列數值 1, Int(Val(FormMainMode.顯示列1.goi1) / 3 + 0.9)
             Else
                 戰鬥系統類.直接寫入顯示列數值 1, Int(Val(FormMainMode.顯示列1.goi1) / 2 + 0.9)
             End If
'             攻擊防禦骰子總數(1) = FormMainMode.顯示列1.goi1
   End Select
End If
End Sub
Sub CC_滅菌空間()
If FormMainMode.comaiatk(1).Caption = "滅菌空間" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(103, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "C.C." Then
   Select Case atkingckai(103, 1)
        Case 1
             If atkingpagetot(2, 4) >= 1 And atkingckai(103, 2) = 0 Then
               atkingckai(103, 2) = 1
            ElseIf atkingpagetot(2, 4) < 1 And atkingckai(103, 2) = 1 Then
               atkingckai(103, 2) = 0
            End If
        Case 2
            atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_滅菌空間_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7275
                   atkingno(i, 6) = 9480
                   atkingno(i, 7) = 103
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            For i = 1 To 3
                回復執行_電腦 1, i
            Next
            atkingckai(103, 2) = 0
            '======================
               戰鬥系統類.執行動作_清除所有異常狀態_電腦
           '======================
   End Select
End If
End Sub
Sub CC_白銀戰機()
Dim bloodntot As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "白銀戰機" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(33, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "C.C." Then
   Select Case atkingckai(33, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 5) >= 2 And atkingckai(33, 2) = 0 Then
'             If atkingpagetot(2, 1) >= 1 And atkingckai(33, 2) = 0 Then
               atkingckai(33, 2) = 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
            ElseIf (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 5) < 2) And atkingckai(33, 2) = 1 Then
'            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(33, 2) = 1 Then
               atkingckai(33, 2) = 0
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
            End If
          End If
        Case 2
            atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_白銀戰機_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -720
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 33
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            atkingckai(33, 2) = 0
            For i = 1 To 3
                If i = 1 Then
                    Randomize
                    bloodntot = Int(Rnd() * 3) + 0
                    If Val(FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption) > 1 And bloodntot < Val(FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption) Then
                       戰鬥系統類.傷害執行_技能直傷_使用者 bloodntot, 1
                    ElseIf Val(FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption) = 2 And bloodntot = 2 Then
                       bloodntot = 1
                       戰鬥系統類.傷害執行_技能直傷_使用者 bloodntot, 1
                    End If
                Else
                    Randomize
                    bloodntot = Int(Rnd() * 3) + 0
                    戰鬥系統類.傷害執行_技能直傷_使用者 bloodntot, i
                End If
            Next
   End Select
End If
End Sub
Sub CC_原子之心()
If FormMainMode.comaiatk(3).Caption = "原子之心" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(57, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "C.C." Then
   Select Case atkingckai(57, 1)
        Case 1
             If atkingpagetot(2, 2) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(57, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(57, 2) = 0 Then
               atkingckai(57, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 2
            ElseIf (atkingpagetot(2, 2) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(57, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(57, 2) = 1 Then
               atkingckai(57, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 2
            End If
        Case 2
            '===========將所有技能無效化-使用者方(階段1)
            atkingtrn(1) = 0
            For i = 1 To UBound(atkingck)
                 atkingck(i, 2) = 0
            Next
        Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_原子之心_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6945
                   atkingno(i, 6) = 10050
                   atkingno(i, 7) = 57
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
              '=================更改數值為原骰數值
              FormMainMode.顯示列1.goi1 = 攻擊防禦骰子總數(3)
              FormMainMode.顯示列1.goi2 = 攻擊防禦骰子總數(4) + 2
              '===================
                For i = 1 To 4
                    戰鬥系統類.人物技能欄燈開關 False, i
                Next
                '==================
                atking_蕾_守護模式狀態啟動值 = False
                Erase atking_史塔夏_殺戮模式狀態數
        Case 4
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 1)
            FormMainMode.personusminijpg.小人物影子圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 2)
            FormMainMode.顯示列1.使用者方小人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 4)
            FormMainMode.顯示列1.使用者方小人物圖片left = -FormMainMode.顯示列1.使用者方小人物圖片width
            FormMainMode.personusminijpg.小人物影子Left = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 1, 5))
            FormMainMode.personusminijpg.小人物影子top差 = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 1, 6))
            FormMainMode.personusminijpg.Visible = True
            Form6.jpgus.大人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 3)
            戰鬥系統類.執行動作_距離變更 movecp
            atkingckai(57, 2) = 0
   End Select
End If
End Sub

Sub CC_高頻電磁手術刀()
If FormMainMode.comaiatk(4).Caption = "高頻電磁手術刀" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(50, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "C.C." Then
   Select Case atkingckai(50, 1)
        Case 1
            If movecp = 1 Then
                If atkingpagetot(2, 4) >= 6 And atkingckai(50, 2) = 0 Then
                   atkingckai(50, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 24
                ElseIf atkingpagetot(2, 4) < 6 And atkingckai(50, 2) = 1 Then
                   atkingckai(50, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 24
                End If
            End If
        Case 2
             atkingckai(50, 1) = 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_高頻電磁手術刀_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(50, 2) = 0
   End Select
End If
End Sub

Sub 帕茉_戰慄的狼牙()
Dim rrr As Integer '暫時變數
If FormMainMode.comaiatk(4).Caption = "戰慄的狼牙" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(34, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "帕茉" Then
   Select Case atkingckai(34, 1)
      Case 1
         If movecp = 1 Then
            If atkingpagetot(2, 1) >= 6 And atkingckai(34, 2) = 0 Then
               atkingckai(34, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(34, 2) = 1 Then
               atkingckai(34, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_戰慄的狼牙_1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6645
                   atkingno(i, 6) = 9330
                   atkingno(i, 7) = 34
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
     Case 3
           For rrr = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
             If 人物異常狀態資料庫(2, rrr, 3) = 26 Then
                回復執行_電腦 人物異常狀態資料庫(2, rrr, 2), 1
                戰鬥系統類.傷害執行_技能直傷_使用者 人物異常狀態資料庫(2, rrr, 2), 1
                Exit For
             End If
           Next
            '=====================
               執行動作_清除所有異常狀態_使用者
               執行動作_清除所有異常狀態_電腦
           '======================
           atkingckai(34, 2) = 0
   End Select
End If
End Sub
Sub 帕茉_慈悲的藍眼()
If FormMainMode.comaiatk(3).Caption = "慈悲的藍眼" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(35, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "帕茉" Then
   Select Case atkingckai(35, 1)
      Case 1
          If movecp > 1 Then
             If atkingpagetot(2, 1) >= 6 And atkingckai(35, 2) = 0 Then
               atkingckai(35, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(35, 2) = 1 Then
               atkingckai(35, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
      Case 2
          atking_AI_帕茉_慈悲的藍眼_tot(1) = atking_AI_帕茉_慈悲的藍眼_tot(1) + 攻擊防禦骰子總數(2)
          攻擊防禦骰子總數(2) = 0
          atking_AI_帕茉_慈悲的藍眼_tot(2) = 1
          atkingckai(35, 1) = 1
      Case 3
          atking_AI_帕茉_慈悲的藍眼_tot(1) = atking_AI_帕茉_慈悲的藍眼_tot(1) + 攻擊防禦骰子總數(2)
          攻擊防禦骰子總數(2) = atking_AI_帕茉_慈悲的藍眼_tot(1)
          atking_AI_帕茉_慈悲的藍眼_tot(1) = 0
          atking_AI_帕茉_慈悲的藍眼_tot(2) = 0
          atkingckai(35, 1) = 1
      Case 4
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_慈悲的藍眼_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6945
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 35
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 5
            Do
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, i, 2) >= 9 And 人物異常狀態資料庫(2, i, 3) = 26 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(2, i, 3) = 26 And 人物異常狀態資料庫(2, i, 2) = 8 Then
                     FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) + 1
                     人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 1
                     Exit Do
                 ElseIf 人物異常狀態資料庫(2, i, 3) = 26 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) <= 7 Then
'                 If 人物異常狀態資料庫(2, i, 3) = 26 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) <= 97 Then
                     FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) + 2
                     人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 2
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 2, i, 26, app_path & "gif\異常狀態\聖痕.gif", 0, 2
                     異常狀態檢查數(26, 1) = 1
                     異常狀態檢查數(26, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
            '================
            回復執行_電腦 2, 1
            '================
            atkingckai(35, 2) = 0
            atkingckai(35, 1) = 0
            Erase atking_AI_帕茉_慈悲的藍眼_tot
   End Select
End If
End Sub
Sub 帕茉_靜謐之背()
If FormMainMode.comaiatk(2).Caption = "靜謐之背" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(36, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "帕茉" Then
   Select Case atkingckai(36, 1)
      Case 1
         If movecp < 3 Then
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 2 And atkingckai(36, 2) = 0 Then
'            If atkingpagetot(2, 1) >= 1 And atkingckai(36, 2) = 0 Then
               atkingckai(36, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 2) And atkingckai(36, 2) = 1 Then
'            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(36, 2) = 1 Then
               atkingckai(36, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        End If
      Case 2
             atkingckai(36, 1) = 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_靜謐之背_2.jpg"
                   atkingno(i, 2) = 2
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
            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, i, 3) = 26 And 人物異常狀態資料庫(2, i, 2) >= 1 Then
                     擲骰表單溝通暫時變數(2) = Val(擲骰表單溝通暫時變數(2)) + 人物異常狀態資料庫(2, i, 2)
                     擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
                     Exit For
                 End If
            Next
           atkingckai(36, 2) = 0
           For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
             If 人物異常狀態資料庫(2, i, 3) = 26 And 人物異常狀態資料庫(2, i, 2) >= 1 Then
                 FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) - 1
                 人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
                 If 人物異常狀態資料庫(2, i, 2) = 0 Then
                     '===繼承下一狀態資料
                     戰鬥系統類.異常狀態繼承_電腦
                 End If
                 Exit For
             End If
           Next
   End Select
End If
End Sub
Sub 艾茵_十三隻眼()
Dim rrr(1 To 2) As Integer '暫時變數
If FormMainMode.comaiatk(4).Caption = "十三隻眼" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(37, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾茵" Then
   Select Case atkingckai(37, 1)
        Case 1
           If movecp < 3 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr(1) = rrr(1) + 1
                End If
                If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr(2) = rrr(2) + 1
                End If
             Next
           End If
          If rrr(1) >= 1 And rrr(2) >= 1 And atkingckai(37, 2) = 0 Then
'          If rrr(1) >= 1 And atkingckai(37, 2) = 0 Then
             atkingckai(37, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If (rrr(1) < 1 Or rrr(2) < 1) And atkingckai(37, 2) = 1 Then
'          If rrr(1) < 1 And atkingckai(37, 2) = 1 Then
             atkingckai(37, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
        Case 2
            If atking_AI_艾茵_十三隻眼_tot(2) = 0 Then
                atking_AI_艾茵_十三隻眼_tot(1) = 攻擊防禦骰子總數(2)
                atking_AI_艾茵_十三隻眼_tot(2) = 1
                攻擊防禦骰子總數(2) = 13
                攻擊防禦骰子總數(1) = 0
                atkingckai(37, 1) = 1
            ElseIf atking_AI_艾茵_十三隻眼_tot(2) = 1 Then
                atking_AI_艾茵_十三隻眼_tot(1) = atking_AI_艾茵_十三隻眼_tot(1) + (攻擊防禦骰子總數(2) - 13)
                攻擊防禦骰子總數(2) = 13
                攻擊防禦骰子總數(1) = 0
                atkingckai(37, 1) = 1
            End If
        Case 3
           atking_AI_艾茵_十三隻眼_tot(1) = atking_AI_艾茵_十三隻眼_tot(1) + (攻擊防禦骰子總數(2) - 13)
           攻擊防禦骰子總數(2) = atking_AI_艾茵_十三隻眼_tot(1)
           atking_AI_艾茵_十三隻眼_tot(1) = 0
           atking_AI_艾茵_十三隻眼_tot(2) = 0
           atkingckai(37, 1) = 1
        Case 4
             戰鬥系統類.自動捲軸捲動
             atkingckai(37, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_十三隻眼_2.jpg"
                   atkingno(i, 2) = 2
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
             Erase atking_AI_艾茵_十三隻眼_tot
             '==============
            戰鬥系統類.直接寫入顯示列數值 2, 13
'            攻擊防禦骰子總數(2) = FormMainMode.顯示列1.goi2
            戰鬥系統類.直接寫入顯示列數值 1, 0
'            攻擊防禦骰子總數(1) = FormMainMode.顯示列1.goi1
        Case 5
            攻擊防禦骰子總數(1) = 0
   End Select
End If
End Sub
Sub 艾茵_兩個身體()
Dim bloodtot As Single  '暫時變數
Dim num As Integer
If FormMainMode.comaiatk(2).Caption = "兩個身體" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(38, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾茵" Then
   Select Case atkingckai(38, 1)
        Case 1
             If atkingpagetot(2, 3) >= 1 And atkingckai(38, 2) = 0 Then
               atkingckai(38, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(38, 2) = 1 Then
               atkingckai(38, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_兩個身體_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6330
                   atkingno(i, 6) = 9285
                   atkingno(i, 7) = 114
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(38, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                bloodtot = Val(擲骰表單溝通暫時變數(2)) \ Val(2)
                Do
                    Randomize
                    num = Int(Rnd() * 3) + 1
                    If liveus(角色待機人物紀錄數(1, num)) > 0 Then
                        戰鬥系統類.傷害執行_技能直傷_使用者 bloodtot, num
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
If FormMainMode.comaiatk(3).Caption = "九個靈魂" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(56, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾茵" Then
   Select Case atkingckai(56, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(2, 2) >= 5 And atkingpagetot(2, 4) >= 1 And atkingckai(56, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(56, 2) = 0 Then
               atkingckai(56, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 9
            ElseIf (atkingpagetot(2, 2) < 5 Or atkingpagetot(2, 4) < 1) And atkingckai(56, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(56, 2) = 1 Then
               atkingckai(56, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 9
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_九個靈魂_2\艾茵_九個靈魂main.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6330
                   atkingno(i, 6) = 9510
                   atkingno(i, 7) = 56
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 11) = 0
                   '=================
                   Randomize
                   pic = Int(Rnd() * 8) + 1
                   atkingno(i, 10) = app_path & "gif\艾茵\艾茵_九個靈魂_2\艾茵_九個靈魂" & pic & ".jpg"
                   Exit For
                 End If
             Next
        Case 3
            bloodtot = Int(atkingpagetot(2, 4) / 2 + 0.5)
            '=============
            If Val(livecom(角色人物對戰人數(2, 2))) < Val(livecommax(角色人物對戰人數(2, 2))) Then
                戰鬥系統類.回復執行_電腦 bloodtot, 1
            End If
            atkingckai(56, 2) = 0
   End Select
End If
End Sub

Sub 夏洛特_冬之夢()
If FormMainMode.comaiatk(2).Caption = "冬之夢" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(39, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "夏洛特" Then
   Select Case atkingckai(39, 1)
      Case 1
            If movecp < 3 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(39, 2) = 0 Then
                   atkingckai(39, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(39, 2) = 1 Then
                   atkingckai(39, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                 End If
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9375
                   atkingno(i, 7) = 39
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(39, 2) = 0
             '========================
             For i = 18 To (turn + 3) Step -1
                  pageeventnum(2, i, 1) = pageeventnum(2, i - 2, 1)
                  pageeventnum(2, i, 2) = pageeventnum(2, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 2)
                  pageeventnum(2, i, 1) = "劍5/槍5"
                  pageeventnum(2, i, 2) = 一般系統類.事件卡資料庫("劍5/槍5", 2)
             Next
   End Select
End If
End Sub
Sub 夏洛特_幸福的理由()
If FormMainMode.comaiatk(4).Caption = "幸福的理由" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(115, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "夏洛特" Then
   Select Case atkingckai(115, 1)
      Case 1
            If atkingpagetot(2, 4) >= 3 And atkingckai(115, 2) = 0 Then
               atkingckai(115, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 3 And atkingckai(115, 2) = 1 Then
               atkingckai(115, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 600
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6915
                   atkingno(i, 6) = 9690
                   atkingno(i, 7) = 115
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(115, 2) = 0
             '========================
             If 牌總階段數(2) > 0 Then
                 牌總階段數(2) = 牌總階段數(2) - 1
             End If
             '========================
             For i = 18 To (turn + 4) Step -1
                  pageeventnum(2, i, 1) = pageeventnum(2, i - 2, 1)
                  pageeventnum(2, i, 2) = pageeventnum(2, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 3)
                  pageeventnum(2, i, 1) = "機會5"
                  pageeventnum(2, i, 2) = 一般系統類.事件卡資料庫("機會5", 2)
             Next
   End Select
End If
End Sub
Sub 夏洛特_大聖堂()
Dim p, i, j As Integer
If FormMainMode.comaiatk(1).Caption = "大聖堂" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(90, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "夏洛特" Then
   Select Case atkingckai(90, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(90, 2) = 0 Then
               atkingckai(90, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(90, 2) = 1 Then
               atkingckai(90, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(90, 1) = 3
        Case 3
             atking_AI_夏洛特_大聖堂骰量紀錄數(1) = 擲骰後骰傷害數
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
                atkingckai(90, 1) = 4
                FormMainMode.骰子執行完啟動.Enabled = False
                目前數(22) = 12
                FormMainMode.等待時間.Enabled = True
          Case 4
                atking_AI_夏洛特_大聖堂骰量紀錄數(2) = 擲骰後骰傷害數
                '==========================
                If atking_AI_夏洛特_大聖堂骰量紀錄數(1) > atking_AI_夏洛特_大聖堂骰量紀錄數(2) Then
                    擲骰表單溝通暫時變數(2) = atking_AI_夏洛特_大聖堂骰量紀錄數(2)
                Else
                    擲骰表單溝通暫時變數(2) = atking_AI_夏洛特_大聖堂骰量紀錄數(1)
                End If
                擲骰後骰傷害數 = Val(擲骰表單溝通暫時變數(2))
                atkingckai(90, 2) = 0
                Erase atking_AI_夏洛特_大聖堂骰量紀錄數
   End Select
End If
End Sub

Sub 泰瑞爾_Rud_913()
If FormMainMode.comaiatk(1).Caption = "Rud-913" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(40, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "泰瑞爾" Then
   Select Case atkingckai(40, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(40, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(40, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(40, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   atkingckai(40, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6675
                   atkingno(i, 6) = 9105
                   atkingno(i, 7) = 40
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(40, 2) = 0
            '================
            戰鬥系統類.執行動作_距離變更 3
   End Select
End If
End Sub
Sub 泰瑞爾_Von_541()
If FormMainMode.comaiatk(2).Caption = "Von-541" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(76, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "泰瑞爾" Then
   Select Case atkingckai(76, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(76, 2) = 0 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
               atkingckai(76, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If (atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(76, 2) = 1 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
               atkingckai(76, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking2_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(76, 2) = 0
            '================
            If 擲骰後骰傷害數 >= 10 Then
                戰鬥系統類.傷害執行_技能直傷_使用者 擲骰後骰傷害數, 1
                擲骰後骰傷害數 = 0
                擲骰表單溝通暫時變數(2) = 0
            End If
   End Select
End If
End Sub

Sub 泰瑞爾_Wil_846()
If FormMainMode.comaiatk(4).Caption = "Wil-846" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(41, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "泰瑞爾" Then
   Select Case atkingckai(41, 1)
      Case 1
           If movecp = 3 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(41, 2) = 0 Then
                   atkingckai(41, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 5) < 2) And atkingckai(41, 2) = 1 Then
                   atkingckai(41, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6720
                   atkingno(i, 6) = 10320
                   atkingno(i, 7) = 41
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(41, 2) = 0
            '================
            戰鬥系統類.傷害執行_技能直傷_使用者 2, 1
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
Sub 泰瑞爾_Chr_799()
Dim m As Integer
If FormMainMode.comaiatk(3).Caption = "Chr-799" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(77, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "泰瑞爾" Then
   Select Case atkingckai(77, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(77, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(77, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 5) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(77, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   atkingckai(77, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\泰瑞爾\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 120
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6750
                   atkingno(i, 6) = 9255
                   atkingno(i, 7) = 77
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(77, 2) = 0
            '================
            m = Int(Rnd() * 3) + 1
            Select Case m
                Case 1
                       Do
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                                  If 人物異常狀態資料庫(1, j, 3) = 10 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 3
                                      FormMainMode.personusspe(j).person_turn = 5
                                      人物異常狀態資料庫(1, j, 1) = 3
                                      人物異常狀態資料庫(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 10, app_path & "gif\異常狀態\atkdown.gif", 3, 5
                                  異常狀態檢查數(10, 1) = 1
                                  異常狀態檢查數(10, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                                If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 3
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 人物異常狀態資料庫(2, j, 1) = 3
                                 人物異常狀態資料庫(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 3, 5
                                 異常狀態檢查數(1, 1) = 1
                                 異常狀態檢查數(1, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
                Case 2
                        Do
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                                  If 人物異常狀態資料庫(1, j, 3) = 11 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 3
                                      FormMainMode.personusspe(j).person_turn = 5
                                      人物異常狀態資料庫(1, j, 1) = 3
                                      人物異常狀態資料庫(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 11, app_path & "gif\異常狀態\defdown.gif", 3, 5
                                  異常狀態檢查數(11, 1) = 1
                                  異常狀態檢查數(11, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                                If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 3
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 人物異常狀態資料庫(2, j, 1) = 3
                                 人物異常狀態資料庫(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 2, app_path & "gif\異常狀態\defup.gif", 3, 5
                                 異常狀態檢查數(2, 1) = 1
                                 異常狀態檢查數(2, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
                Case 3
                        Do
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                                  If 人物異常狀態資料庫(1, j, 3) = 12 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 1
                                      FormMainMode.personusspe(j).person_turn = 5
                                      人物異常狀態資料庫(1, j, 1) = 1
                                      人物異常狀態資料庫(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, j, 12, app_path & "gif\異常狀態\movdown.gif", 1, 5
                                  異常狀態檢查數(12, 1) = 1
                                  異常狀態檢查數(12, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                                If 人物異常狀態資料庫(2, j, 3) = 3 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 1
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 人物異常狀態資料庫(2, j, 1) = 1
                                 人物異常狀態資料庫(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 3, app_path & "gif\異常狀態\movup.gif", 1, 5
                                 異常狀態檢查數(3, 1) = 1
                                 異常狀態檢查數(3, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
            End Select
   End Select
End If
End Sub
Sub 瑪格莉特_月光()
Dim m As Integer
If FormMainMode.comaiatk(1).Caption = "月光" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(78, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "瑪格莉特" Then
   Select Case atkingckai(78, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 4) >= 1 And atkingckai(78, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
                   atkingckai(78, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 1 And atkingckai(78, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
                   atkingckai(78, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -240
                   atkingno(i, 5) = 6195
                   atkingno(i, 6) = 10350
                   atkingno(i, 7) = 78
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
             Erase atking_AI_瑪格莉特_月光紀錄數
             '========================
             For i = 1 To 106
                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                    If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                         atking_AI_瑪格莉特_月光紀錄數(i) = 1
                         atking_AI_瑪格莉特_月光紀錄數(107) = atking_AI_瑪格莉特_月光紀錄數(107) + 1
                     End If
                End If
            Next
            If atking_AI_瑪格莉特_月光紀錄數(107) > 2 Then
                atking_AI_瑪格莉特_月光紀錄數(107) = 2
            End If
            '=========================
            If atking_AI_瑪格莉特_月光紀錄數(107) > 0 Then
                Do
                    m = Int(Rnd() * 106) + 1
                    If atking_AI_瑪格莉特_月光紀錄數(m) = 1 Then
                        目前數(20) = m
                        目前數(21) = 5
                        atking_AI_瑪格莉特_月光紀錄數(m) = 0
                        atking_AI_瑪格莉特_月光紀錄數(0) = atking_AI_瑪格莉特_月光紀錄數(0) + 1
                        FormMainMode.tr使用者_棄牌.Enabled = True
                        Exit Sub
                    End If
                Loop
            Else
               目前數(22) = 25
               FormMainMode.等待時間.Enabled = True
            End If
        Case 4
            If atking_AI_瑪格莉特_月光紀錄數(107) > 1 And atking_AI_瑪格莉特_月光紀錄數(0) < atking_AI_瑪格莉特_月光紀錄數(107) Then
                Do
                    m = Int(Rnd() * 106) + 1
                    If atking_AI_瑪格莉特_月光紀錄數(m) = 1 Then
                        目前數(20) = m
                        目前數(21) = 5
                        atking_AI_瑪格莉特_月光紀錄數(m) = 0
                        atking_AI_瑪格莉特_月光紀錄數(0) = atking_AI_瑪格莉特_月光紀錄數(0) + 1
                        FormMainMode.tr使用者_棄牌.Enabled = True
                        Exit Sub
                    End If
                Loop
            ElseIf atking_AI_瑪格莉特_月光紀錄數(0) >= 2 Then
               目前數(24) = 26
               FormMainMode.等待時間_2.Enabled = True
            Else
               目前數(24) = 25
               FormMainMode.等待時間_2.Enabled = True
            End If
        Case 5
            If atking_AI_瑪格莉特_月光紀錄數(107) = 0 Then
                atking_AI_瑪格莉特_月光紀錄數(107) = 99
               目前數(22) = 25
               FormMainMode.等待時間.Enabled = True
            ElseIf atking_AI_瑪格莉特_月光紀錄數(107) > 0 And atking_AI_瑪格莉特_月光紀錄數(0) = 0 Then
               atkingckai(78, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            ElseIf atking_AI_瑪格莉特_月光紀錄數(107) > 0 And atking_AI_瑪格莉特_月光紀錄數(0) = 1 Then
               戰鬥系統類.傷害執行_技能直傷_使用者 atking_AI_瑪格莉特_月光紀錄數(0), 1
               目前數(24) = 26
               FormMainMode.等待時間_2.Enabled = True
            End If
        Case 6
            If atking_AI_瑪格莉特_月光紀錄數(107) > 0 And atking_AI_瑪格莉特_月光紀錄數(0) = 1 Then
               atkingckai(78, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            ElseIf atking_AI_瑪格莉特_月光紀錄數(0) >= 2 Then
               戰鬥系統類.傷害執行_技能直傷_使用者 atking_AI_瑪格莉特_月光紀錄數(0), 1
               atkingckai(78, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            End If
   End Select
End If
End Sub

Sub 瑪格莉特_恍惚()
If FormMainMode.comaiatk(2).Caption = "恍惚" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(42, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "瑪格莉特" Then
   Select Case atkingckai(42, 1)
        Case 1
            If movecp = 1 Then
             If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(42, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(42, 2) = 0 Then
               atkingckai(42, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
            ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(42, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(42, 2) = 1 Then
               atkingckai(42, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5580
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 126
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(42, 2) = 0
            '===============
            If 擲骰後骰傷害數 <= 0 Then
                Do
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 3) = 16 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 2
                          人物異常狀態資料庫(1, i, 2) = 2
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, i, 16, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                          異常狀態檢查數(16, 1) = 1
                          異常狀態檢查數(16, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub 瑪格莉特_地獄獵心獸()
Dim m As Integer
If FormMainMode.comaiatk(4).Caption = "地獄獵心獸" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(43, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "瑪格莉特" Then
   Select Case atkingckai(43, 1)
        Case 1
             If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(43, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(43, 2) = 0 Then
               atkingckai(43, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(43, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(43, 2) = 1 Then
               atkingckai(43, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6630
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 43
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(43, 2) = 0
            '===============
            m = (atkingpagetot(2, 1) + atkingpagetot(2, 5)) \ 5
            戰鬥系統類.傷害執行_技能直傷_使用者 m, 1
   End Select
End If
End Sub
Sub 庫勒尼西_沙漠中的海市蜃樓()
If FormMainMode.comaiatk(1).Caption = "沙漠中的海市蜃樓" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(44, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "庫勒尼西" Then
   Select Case atkingckai(44, 1)
      Case 1
           If movecp = 3 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 2
                atkingckai(44, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
      Case 2
             atkingckai(44, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
If FormMainMode.comaiatk(2).Caption = "瘋狂眼窩" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(79, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "庫勒尼西" Then
   Select Case atkingckai(79, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 1 And atkingckai(79, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 2
                   atkingckai(79, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 1 And atkingckai(79, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 2
                   atkingckai(79, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking2_2.jpg"
                   atkingno(i, 2) = 2
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
             If Val(FormMainMode.pageusglead.Caption) > 0 Then
                 atking_AI_庫勒尼西_瘋狂眼窩紀錄數 = 1
                 '==========================
                  Do Until atking_AI_庫勒尼西_瘋狂眼窩紀錄數 > 3 Or Val(FormMainMode.pageusglead.Caption) <= 0
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                            目前數(21) = 6
                            目前數(20) = m
                            atking_AI_庫勒尼西_瘋狂眼窩紀錄數 = atking_AI_庫勒尼西_瘋狂眼窩紀錄數 + 1
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                   Loop
             Else
                 atkingckai(79, 1) = 5
                 FormMainMode.骰子執行完啟動.Enabled = True
             End If
        Case 4
             Do Until atking_AI_庫勒尼西_瘋狂眼窩紀錄數 > 3 Or Val(FormMainMode.pageusglead.Caption) <= 0
                 Randomize
                 m = Int(Rnd() * 106) + 1
                 If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                     目前數(21) = 6
                     目前數(20) = m
                     atking_AI_庫勒尼西_瘋狂眼窩紀錄數 = atking_AI_庫勒尼西_瘋狂眼窩紀錄數 + 1
                     FormMainMode.tr使用者_棄牌.Enabled = True
                     Exit Sub
                 End If
            Loop
            If atking_AI_庫勒尼西_瘋狂眼窩紀錄數 > 3 Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
                atkingckai(79, 1) = 5
                目前數(24) = 22
                FormMainMode.等待時間_2.Enabled = True
            End If
        Case 5
            atkingckai(79, 2) = 0
   End Select
End If
End Sub

Sub 庫勒尼西_黑暗漩渦()
Dim m As Integer
If FormMainMode.comaiatk(4).Caption = "黑暗漩渦" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(46, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "庫勒尼西" Then
   Select Case atkingckai(46, 1)
        Case 1
             If atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(46, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(46, 2) = 0 Then
               atkingckai(46, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
            ElseIf (atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(46, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(46, 2) = 1 Then
               atkingckai(46, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6480
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 46
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(46, 2) = 0
            '===============
            m = movecp + 1
            If m > 3 Then m = 3
            戰鬥系統類.執行動作_距離變更 m
   End Select
End If
End Sub
Sub 庫勒尼西_深淵()
Dim m As Integer
If FormMainMode.comaiatk(3).Caption = "深淵" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(45, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "庫勒尼西" Then
   Select Case atkingckai(45, 1)
        Case 1
             If atkingpagetot(2, 4) >= 3 And atkingckai(45, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(45, 2) = 0 Then
               atkingckai(45, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 4) < 3 And atkingckai(45, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(45, 2) = 1 Then
               atkingckai(45, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\庫勒尼西\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8970
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 45
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(45, 2) = 0
            '===============
            m = Int(atkingpagetot(2, 4) / 2 + 0.9)
            戰鬥系統類.傷害執行_技能直傷_使用者 m, 1
   End Select
End If
End Sub
Sub 蕾格烈芙_CTL()
Dim i As Integer
If FormMainMode.comaiatk(1).Caption = "C.T.L" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(80, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾格烈芙" Then
   Select Case atkingckai(80, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(80, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(80, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   For i = 1 To 106
                       If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                              攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                              目前數(28) = 1
                              Exit For
                          End If
                       End If
                   Next
                End If
                If (atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) And atkingckai(80, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   If 目前數(28) = 1 Then
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                       目前數(28) = 0
                   End If
                   atkingckai(80, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(80, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6540
                   atkingno(i, 6) = 9990
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\蕾格烈芙\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             If 目前數(28) = 1 Then
                 For i = 1 To 106
                       If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                              Exit For
                          End If
                       End If
                  Next
                  If i = 107 Then
                      攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                      直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 6
                      For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, i, 3) = 32 Then
                              攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                              直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 6
                          End If
                      Next
                  End If
                  目前數(28) = 0
             End If
   End Select
End If
End Sub
Sub 蕾格烈芙_BPA()
If FormMainMode.comaiatk(2).Caption = "B.P.A" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(81, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾格烈芙" Then
   Select Case atkingckai(81, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(81, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
                   atkingckai(81, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(81, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
                   atkingckai(81, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6030
                   atkingno(i, 6) = 10530
                   atkingno(i, 7) = 81
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             戰鬥系統類.傷害執行_技能直傷_使用者 pageqlead(1), 1
             atkingckai(81, 2) = 0
   End Select
End If
End Sub

Sub 蕾格烈芙_LAR()
If FormMainMode.comaiatk(3).Caption = "L.A.R" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(47, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾格烈芙" Then
   Select Case atkingckai(47, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 2) >= 2 And atkingckai(47, 2) = 0 Then
                   atkingckai(47, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 2) < 2 And atkingckai(47, 2) = 1 Then
                   atkingckai(47, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5400
                   atkingno(i, 6) = 9015
                   atkingno(i, 7) = 47
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             戰鬥系統類.回復執行_電腦 1, 1
        Case 4
             atkingckai(47, 2) = 0
             If 擲骰後骰傷害數 > 0 Then
                 戰鬥系統類.回復執行_電腦 1, 1
             End If
   End Select
End If
End Sub
Sub 傑多_因果之幻()
Dim p, i, j As Integer
Dim ak As Integer
If FormMainMode.comaiatk(4).Caption = "因果之幻" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(48, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "傑多" Then
   Select Case atkingckai(48, 1)
      Case 1
            If atkingpagetot(2, 3) >= 1 And atkingckai(48, 2) = 0 Then
'            If pageqlead(2) >= 1 And atkingckai(48, 2) = 0 Then
               atkingckai(48, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
            End If
            If atkingpagetot(2, 3) < 1 And atkingckai(48, 2) = 1 Then
'            If pageqlead(2) < 1 And atkingckai(48, 2) = 1 Then
               atkingckai(48, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7590
                   atkingno(i, 6) = 9420
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(48, 1) = 3
        Case 3
             atking_AI_傑多_因果之幻骰量紀錄數(1) = 擲骰後骰傷害數
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
                atkingckai(48, 1) = 4
                FormMainMode.骰子執行完啟動.Enabled = False
                目前數(22) = 12
                FormMainMode.等待時間.Enabled = True
          Case 4
                atking_AI_傑多_因果之幻骰量紀錄數(2) = 擲骰後骰傷害數
                '==========================
                If atking_AI_傑多_因果之幻骰量紀錄數(1) < atking_AI_傑多_因果之幻骰量紀錄數(2) Then
                    擲骰表單溝通暫時變數(2) = atking_AI_傑多_因果之幻骰量紀錄數(2)
                Else
                    擲骰表單溝通暫時變數(2) = atking_AI_傑多_因果之幻骰量紀錄數(1)
                End If
                擲骰後骰傷害數 = Val(擲骰表單溝通暫時變數(2))
                atkingckai(48, 2) = 0
                Erase atking_AI_傑多_因果之幻骰量紀錄數
          Case 5
                 atkingckai(48, 1) = 1
                  For j = 49 To 54   '防1移1卡優先
                      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                              戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                              ak = 1
                              Exit For
                           ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                              戰鬥系統類.comatk_AI_雪莉_飛刃雨_移 j
                              ak = 1
                              Exit For
                           End If
                      End If
                  Next
                  If ak = 0 Then
                     For j = 39 To 44   '槍1移1卡其次優先
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                              戰鬥系統類.comatk_AI_傑多_因果之幻_移 j
                              ak = 1
                              Exit For
                           ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                              戰鬥系統類.comatk_AI_傑多_因果之幻_移 j
                              ak = 1
                              Exit For
                           End If
                        End If
                     Next
                  End If
                  If ak = 0 Then
                     For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) >= 1 Then
                              戰鬥系統類.comatk_AI_傑多_因果之幻_移 j
                              ak = 1
                              Exit For
                           ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) >= 1 Then
                              戰鬥系統類.comatk_AI_傑多_因果之幻_移 j
                              ak = 1
                              Exit For
                           End If
                        End If
                     Next
                  End If
   End Select
End If
End Sub
Sub 伊芙琳_紅蓮車輪()
Dim bloodtot As Integer '暫時變數
Dim num(1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.comaiatk(3).Caption = "紅蓮車輪" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(51, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "伊芙琳" Then
   Select Case atkingckai(51, 1)
        Case 1
            If movecp < 3 Then
                 If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 5) >= 2 And atkingpagetot(2, 4) >= 1 And atkingckai(51, 2) = 0 Then
    '             If atkingpagetot(2, 1) >= 1 And atkingckai(51, 2) = 0 Then
                   atkingckai(51, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 13
                ElseIf (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 5) < 2 Or atkingpagetot(2, 4) < 1) And atkingckai(51, 2) = 1 Then
    '            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(51, 2) = 1 Then
                   atkingckai(51, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 13
                End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 8250
                   atkingno(i, 6) = 10275
                   atkingno(i, 7) = 51
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\伊芙琳\Evelynatking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            Do
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, i, 2) >= 9 And 人物異常狀態資料庫(2, i, 3) = 25 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(2, i, 3) = 25 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 9 Then
                     FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) + 1
                     人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 2, i, 25, app_path & "gif\異常狀態\能力低下.gif", 0, 1
                     異常狀態檢查數(25, 1) = 1
                     異常狀態檢查數(25, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
        Case 4
            bloodtot = Val(FormMainMode.顯示列1.goi2) \ 10
            num(2) = 999
            For i = 1 To 3
               If liveus(角色待機人物紀錄數(1, i)) < num(2) And liveus(角色待機人物紀錄數(1, i)) > 0 Then
                   num(1) = i
                   num(2) = liveus(角色待機人物紀錄數(1, i))
               End If
            Next
            戰鬥系統類.傷害執行_技能直傷_使用者 bloodtot, num(1)
            atkingckai(51, 2) = 0
   End Select
End If
End Sub
Sub 多妮妲_律死擊()
If FormMainMode.comaiatk(4).Caption = "律死擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(52, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "多妮妲" Then
   Select Case atkingckai(52, 1)
      Case 1
            If movecp = 1 Then
                    If atkingpagetot(2, 1) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(52, 2) = 0 Then
                       atkingckai(52, 2) = 1
                       atkingtrn(2) = Val(atkingtrn(2)) + 1
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 8
                    End If
                    If (atkingpagetot(2, 1) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(52, 2) = 1 Then
                       atkingckai(52, 2) = 0
                       atkingtrn(2) = Val(atkingtrn(2)) - 1
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 8
                     End If
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\多妮妲\atking4_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(52, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                          If 人物異常狀態資料庫(1, j, 3) = 15 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 0
                              FormMainMode.personusspe(j).person_turn = 5
                              人物異常狀態資料庫(1, j, 1) = 0
                              人物異常狀態資料庫(1, j, 2) = 5
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, j, 15, app_path & "gif\異常狀態\自壞.gif", 0, 5
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
Sub 多妮妲_殘虐傾向()
If FormMainMode.comaiatk(1).Caption = "殘虐傾向" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(53, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "多妮妲" Then
   Select Case atkingckai(53, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(53, 2) = 0 Then
               atkingckai(53, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(53, 2) = 1 Then
               atkingckai(53, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\多妮妲\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(53, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                       Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 20 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  人物異常狀態資料庫(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 20, app_path & "gif\異常狀態\damage.gif", 0, 2
                                  異常狀態檢查數(20, 1) = 1
                                  異常狀態檢查數(20, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 16 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  人物異常狀態資料庫(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 16, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                                  異常狀態檢查數(16, 1) = 1
                                  異常狀態檢查數(16, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 22 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  人物異常狀態資料庫(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 22, app_path & "gif\異常狀態\atkingerr.gif", 0, 2
                                  異常狀態檢查數(22, 1) = 1
                                  異常狀態檢查數(22, 2) = 1
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
Dim atkingtotai As Integer '特數量暫時統計變數
Dim a As Integer, i As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "異質者" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(82, 2) = 1) _
    And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "多妮妲" Then
 Select Case atkingckai(82, 1)
   Case 1
      atkingckai(82, 1) = 2
      For i = 55 To 106
         If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And ((Val(pagecardnum(i, 2)) = 3 And pagecardnum(i, 1) = a4a) Or (Val(pagecardnum(i, 4)) = 3 And pagecardnum(i, 3) = a4a)) Then
            atkingtotai = Val(atkingtotai) + 1
         End If
      Next
      If atkingtotai >= 1 Then
         Select Case livecom(角色人物對戰人數(2, 2))
            Case Is < 3
                If Val(FormMainMode.顯示列1.goi1) - Val(FormMainMode.顯示列1.goi2) >= livecom(角色人物對戰人數(2, 2)) Then
                    GoTo AI技能_多妮妲_異質者_出牌階段二
                End If
            Case 3
                If Val(FormMainMode.顯示列1.goi1) - Val(FormMainMode.顯示列1.goi2) >= 9 Then
                    GoTo AI技能_多妮妲_異質者_出牌階段二
                End If
            Case Is > 3
                If Int(Val(FormMainMode.顯示列1.goi1) / 3 + 0.9) - Int(Val(FormMainMode.顯示列1.goi2) / 3 + 0.9) >= livecom(角色人物對戰人數(2, 2)) Then
                    GoTo AI技能_多妮妲_異質者_出牌階段二
                End If
         End Select
      End If
      '==========如果不符合任何條件時
      Exit Sub
    '================================
AI技能_多妮妲_異質者_出牌階段二:
      For a = 55 To 106
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) = 3 Then
                    戰鬥系統類.comatk_AI_雪莉_多妮妲_異質者_特 a
                    Exit For
                ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) = 3 Then
                    戰鬥系統類.comatk_AI_雪莉_多妮妲_異質者_特 a
                    Exit For
                End If
             End If
      Next
    Case 2
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
'                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) >= 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
             If rrr >= 1 And atkingckai(82, 2) = 0 Then
                atkingckai(82, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
             If rrr < 1 And atkingckai(82, 2) = 1 Then
                atkingckai(82, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
   Case 3
        For i = 人物技能數字指示 To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\多妮妲\atking2_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 10110
                atkingno(i, 7) = 82
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
          atkingckai(82, 2) = 0
          If Val(擲骰表單溝通暫時變數(2)) - Val(擲骰表單溝通暫時變數(3)) >= livecom(角色人物對戰人數(2, 2)) And 異常狀態檢查數(18, 2) = 0 Then
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 6
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 1) = 6
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 6, 3
                          異常狀態檢查數(1, 1) = 1
                          異常狀態檢查數(1, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '==================================
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 18 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 1) = 0
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 18, app_path & "gif\異常狀態\不死.gif", 0, 3
                          異常狀態檢查數(18, 1) = 1
                          異常狀態檢查數(18, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '===============================
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 19 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 1) = 0
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 19, app_path & "gif\異常狀態\自壞.gif", 0, 3
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

Sub 蕾_協奏曲_加百烈的守護()
Dim i As Integer, j As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "協奏曲-加百烈的守護" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(54, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(54, 1)
        Case 1
            If atkingpagetot(2, 4) >= 2 And atkingpagetot(2, 3) >= 1 And atkingckai(54, 2) = 0 Then
'            If atkingpagetot(2, 3) >= 1 And atkingckai(54, 2) = 0 Then
               atkingckai(54, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 4) < 2 Or atkingpagetot(2, 3) < 1) And atkingckai(54, 2) = 1 Then
'            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(54, 2) = 1 Then
               atkingckai(54, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-協奏曲-加百烈的守護_2.jpg"
                   atkingno(i, 2) = 2
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
            atkingckai(54, 2) = 0
            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, j, 1) >= 10 And 人物異常狀態資料庫(2, j, 3) = 2 Then
                     FormMainMode.personcomspe(j).person_turn = 3
                     人物異常狀態資料庫(2, j, 2) = 3
                     Exit Do
                 End If
                 If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                     FormMainMode.personcomspe(j).person_num = 人物異常狀態資料庫(2, j, 1) + 1
                     FormMainMode.personcomspe(j).person_turn = 3
                     人物異常狀態資料庫(2, j, 1) = 人物異常狀態資料庫(2, j, 1) + 1
                     人物異常狀態資料庫(2, j, 2) = 3
                     '========DEF+1立即生效
'                         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 1
                         戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) + 1
                    '===============
                     Exit Do
                 End If
            Next
           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 2, app_path & "gif\異常狀態\defup.gif", 3, 3
                 異常狀態檢查數(2, 1) = 1
                 異常狀態檢查數(2, 2) = 1
                  '========DEF+3立即生效
'                         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
                         戰鬥系統類.直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) + 3
                  '===============
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub 阿奇波爾多_大地崩壞()
If FormMainMode.comaiatk(1).Caption = "大地崩壞" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(89, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿奇波爾多" Then
   Select Case atkingckai(89, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(89, 2) = 0 Then
                   atkingckai(89, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(89, 2) = 1 Then
                   atkingckai(89, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8055
                   atkingno(i, 6) = 10620
                   atkingno(i, 7) = 89
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\阿奇波爾多\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(89, 2) = 0
             '=================
             戰鬥系統類.傷害執行_技能直傷_使用者 2, 1
   End Select
End If
End Sub

Sub 阿奇波爾多_致命槍擊()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "致命槍擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(83, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿奇波爾多" Then
   Select Case atkingckai(83, 1)
      Case 1
            If movecp > 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
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
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(83, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(83, 2) = 0 Then
                atkingckai(83, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 9
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(83, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(83, 2) = 1 Then
                atkingckai(83, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 9
              End If
      Case 2
             atkingckai(83, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking2-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8520
                   atkingno(i, 6) = 8280
                   atkingno(i, 7) = 150
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\阿奇波爾多\atking2-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub 阿奇波爾多_劫影攻擊()
If FormMainMode.comaiatk(3).Caption = "劫影攻擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(84, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿奇波爾多" Then
   Select Case atkingckai(84, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(84, 2) = 0 Then
               atkingckai(84, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 1 And atkingckai(84, 2) = 1 Then
               atkingckai(84, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking3_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(84, 2) = 0
             '======================
             If 擲骰後骰傷害數 > 0 Then
               Do
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 3) = 16 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 2
                          人物異常狀態資料庫(1, i, 2) = 2
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, i, 16, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                          異常狀態檢查數(16, 1) = 1
                          異常狀態檢查數(16, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub

Sub 阿奇波爾多_防護射擊()
If FormMainMode.comaiatk(4).Caption = "防護射擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(49, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿奇波爾多" Then
   Select Case atkingckai(49, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (atkingpagetot(2, 5) - atking_AI_阿奇波爾多_防護射擊_槍數值紀錄數)
                   atking_AI_阿奇波爾多_防護射擊_槍數值紀錄數 = atkingpagetot(2, 5)
                   If atkingckai(49, 2) = 0 Then
                        atkingckai(49, 2) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) + 1
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 2
                   End If
                End If
                If atkingpagetot(2, 5) < 1 And atkingckai(49, 2) = 1 Then
                   atkingckai(49, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - atking_AI_阿奇波爾多_防護射擊_槍數值紀錄數 - 2
                   atking_AI_阿奇波爾多_防護射擊_槍數值紀錄數 = 0
                 End If
          End If
      Case 2
             atkingckai(49, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿奇波爾多\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7725
                   atkingno(i, 6) = 9345
                   atkingno(i, 7) = 49
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atking_AI_阿奇波爾多_防護射擊_槍數值紀錄數 = 0
   End Select
End If
End Sub
Sub 伊芙琳_慟哭之歌()
If FormMainMode.comaiatk(2).Caption = "慟哭之歌" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(61, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "伊芙琳" Then
   Select Case atkingckai(61, 1)
        Case 1
            If movecp > 1 Then
                If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 1 And atkingckai(61, 2) = 0 Then
    '             If atkingpagetot(2, 2) >= 1 And atkingckai(61, 2) = 0 Then
                   atkingckai(61, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 1) And atkingckai(61, 2) = 1 Then
    '            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(61, 2) = 1 Then
                   atkingckai(61, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6195
                   atkingno(i, 6) = 8730
                   atkingno(i, 7) = 61
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
             戰鬥系統類.直接寫入顯示列數值 1, Val(FormMainMode.顯示列1.goi1) \ 2
'             攻擊防禦骰子總數(1) = FormMainMode.顯示列1.goi1
        Case 3
            atkingckai(61, 2) = 0
            Do
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, i, 2) >= 9 And 人物異常狀態資料庫(2, i, 3) = 25 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(2, i, 3) = 25 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 9 Then
                     FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) + 1
                     人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 2, i, 25, app_path & "gif\異常狀態\能力低下.gif", 0, 1
                     異常狀態檢查數(25, 1) = 1
                     異常狀態檢查數(25, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
   End Select
End If
End Sub
Sub 利恩_劫影攻擊()
If FormMainMode.comaiatk(1).Caption = "劫影攻擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(72, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "利恩" Then
   Select Case atkingckai(72, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(72, 2) = 0 Then
               atkingckai(72, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 1 And atkingckai(72, 2) = 1 Then
               atkingckai(72, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(72, 2) = 0
             '======================
             If 擲骰後骰傷害數 > 0 Then
                    Do
                         For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                           If 人物異常狀態資料庫(1, i, 3) = 16 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                               FormMainMode.personusspe(i).person_turn = 2
                               人物異常狀態資料庫(1, i, 2) = 2
                               Exit Do
                           End If
                         Next
                         For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                            If 人物異常狀態資料庫(1, i, 2) = 0 Then
                               戰鬥系統類.人物異常狀態表設定_初設 1, i, 16, app_path & "gif\異常狀態\moveerr.gif", 0, 2
                               異常狀態檢查數(16, 1) = 1
                               異常狀態檢查數(16, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
             End If
   End Select
End If
End Sub
Sub 利恩_毒牙()
If FormMainMode.comaiatk(2).Caption = "毒牙" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(73, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "利恩" Then
   Select Case atkingckai(73, 1)
      Case 1
            If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 4) >= 3 And atkingckai(73, 2) = 0 Then
                   atkingckai(73, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 4) < 3) And atkingckai(73, 2) = 1 Then
                   atkingckai(73, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
                 End If
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking2_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(73, 2) = 0
             '======================
             If 擲骰後骰傷害數 > 0 Then
               Do
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 3) = 20 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 3
                          人物異常狀態資料庫(1, i, 2) = 3
                          Exit Do
                      End If
                    Next
                    For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                       If 人物異常狀態資料庫(1, i, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 1, i, 20, app_path & "gif\異常狀態\damage.gif", 0, 3
                          異常狀態檢查數(20, 1) = 1
                          異常狀態檢查數(20, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub 利恩_反擊的狼煙()
If FormMainMode.comaiatk(3).Caption = "反擊的狼煙" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(74, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "利恩" Then
   Select Case atkingckai(74, 1)
        Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(74, 2) = 0 Then
               atkingckai(74, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 4) < 1 And atkingckai(74, 2) = 1 Then
               atkingckai(74, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking3_2.jpg"
                   atkingno(i, 2) = 2
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
            If 擲骰後骰傷害數 > 0 And livecom(角色人物對戰人數(2, 2)) > 0 Then
                atking_AI_利恩_反擊的狼煙紀錄數(1) = 擲骰後骰傷害數 + 1
                If Val(FormMainMode.pageul.Caption) < atking_AI_利恩_反擊的狼煙紀錄數(1) And atking_AI_利恩_反擊的狼煙紀錄數(2) = 0 Then
                   戰鬥系統類.執行動作_洗牌
                End If
                atking_AI_利恩_反擊的狼煙紀錄數(2) = atking_AI_利恩_反擊的狼煙紀錄數(2) + 1
                If Val(FormMainMode.pageul.Caption) > 0 Then
                    Do Until atking_AI_利恩_反擊的狼煙紀錄數(2) > atking_AI_利恩_反擊的狼煙紀錄數(1)
                        目前數(15) = 25
                        FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                        Exit Sub
                    Loop
                End If
            End If
            If atking_AI_利恩_反擊的狼煙紀錄數(2) > atking_AI_利恩_反擊的狼煙紀錄數(1) Or 擲骰後骰傷害數 <= 0 _
                Or Val(FormMainMode.pageul.Caption) <= 0 Or livecom(角色人物對戰人數(2, 2)) <= 0 Then
                目前數(24) = 22
                atkingckai(74, 1) = 4
                FormMainMode.等待時間_2.Enabled = True
            End If
        Case 4
            atkingckai(74, 2) = 0
            Erase atking_AI_利恩_反擊的狼煙紀錄數
   End Select
End If
End Sub
Sub 利恩_背刺()
If FormMainMode.comaiatk(4).Caption = "背刺" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(75, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "利恩" Then
   Select Case atkingckai(75, 1)
      Case 1
            If movecp = 3 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(75, 2) = 0 Then
                   If 執行動作_檢查是否有指定異常狀態(1, 16) = True Then
                        atkingckai(75, 2) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) + 1
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 12
                    End If
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(75, 2) = 1 Then
                   atkingckai(75, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 12
                 End If
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\利恩\atking4_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(75, 2) = 0
             If 執行動作_檢查是否有指定異常狀態(1, 16) = False Then
                 直接寫入顯示列數值 2, Val(FormMainMode.顯示列1.goi2) - 12
                 攻擊防禦骰子總數(2) = Val(FormMainMode.顯示列1.goi2)
             End If
   End Select
End If
End Sub
Sub 洛洛妮_風暴感知()
If FormMainMode.comaiatk(2).Caption = "風暴感知" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(85, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "洛洛妮" Then
   Select Case atkingckai(85, 1)
        Case 1
             If atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(85, 2) = 0 Then
               atkingckai(85, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + pageqlead(1) * 2
            ElseIf (atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(85, 2) = 1 Then
               atkingckai(85, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - pageqlead(1) * 2
            End If
        Case 2
             atkingckai(85, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking2_2.jpg"
                   atkingno(i, 2) = 2
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
If FormMainMode.comaiatk(3).Caption = "砲擊壓制" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(86, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "洛洛妮" Then
   Select Case atkingckai(86, 1)
        Case 1
             If movecp = 3 Then
                     If atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(86, 2) = 0 Then
                       atkingckai(86, 2) = 1
                       atkingtrn(2) = Val(atkingtrn(2)) + 1
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 8
                    ElseIf (atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(86, 2) = 1 Then
                       atkingckai(86, 2) = 0
                       atkingtrn(2) = Val(atkingtrn(2)) - 1
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 8
                    End If
             End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking3_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(86, 2) = 0
             If 擲骰後骰傷害數 > 0 Then
                    Do
                         For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                           If 人物異常狀態資料庫(1, i, 3) = 10 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                FormMainMode.personusspe(j).person_num = 10
                                FormMainMode.personusspe(j).person_turn = 1
                                人物異常狀態資料庫(1, j, 1) = 10
                                人物異常狀態資料庫(1, j, 2) = 1
                               Exit Do
                           End If
                         Next
                         For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                            If 人物異常狀態資料庫(1, i, 2) = 0 Then
                               戰鬥系統類.人物異常狀態表設定_初設 1, i, 10, app_path & "gif\異常狀態\atkdown.gif", 10, 1
                               異常狀態檢查數(10, 1) = 1
                               異常狀態檢查數(10, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
              End If
   End Select
End If
End Sub
Sub 洛洛妮_貪婪之刃與嗜血之槍()
If FormMainMode.comaiatk(4).Caption = "貪婪之刃與嗜血之槍" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(87, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "洛洛妮" Then
   Select Case atkingckai(87, 1)
        Case 1
             If movecp = 1 Then
                     If atkingpagetot(2, 1) >= 5 And atkingpagetot(2, 5) >= 5 And atkingckai(87, 2) = 0 Then
'                     If pageqlead(2) >= 1 And atkingckai(87, 2) = 0 Then
                       atkingckai(87, 2) = 1
                       atkingtrn(2) = Val(atkingtrn(2)) + 1
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                    ElseIf (atkingpagetot(2, 1) < 5 Or atkingpagetot(2, 5) < 5) And atkingckai(87, 2) = 1 Then
'                    ElseIf pageqlead(2) < 1 And atkingckai(87, 2) = 1 Then
                       atkingckai(87, 2) = 0
                       atkingtrn(2) = Val(atkingtrn(2)) - 1
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                    End If
             End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9495
                   atkingno(i, 6) = 9360
                   atkingno(i, 7) = 87
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
             atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0
        Case 3
             For i = 1 To 106
                 If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                     目前數(20) = i
                     目前數(21) = 7
                     atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 + 1
                     FormMainMode.tr使用者牌_偷牌.Enabled = True
                     Exit Sub
                 End If
             Next
             If i = 107 Then
                 If atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0 Then
                     For k = 1 To 3
                         戰鬥系統類.傷害執行_技能直傷_使用者 3, k
                     Next
                     目前數(22) = 28
                     FormMainMode.等待時間.Enabled = True
                 ElseIf atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 1 Or atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 2 Then
                     目前數(24) = 29
                     FormMainMode.等待時間_2.Enabled = True
                 ElseIf atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 > 2 Then
                     atkingckai(87, 2) = 0
                     戰鬥系統類.執行動作_技能手動結束
                 End If
             End If
        Case 4
             If atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0 Then
                atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 99
                目前數(22) = 37
                FormMainMode.等待時間.Enabled = True
             ElseIf atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 1 Then
                For k = 1 To 3
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, k
                Next
                atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 99
                目前數(24) = 29
                 FormMainMode.等待時間_2.Enabled = True
             ElseIf atking_AI_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 2 Then
                For k = 1 To 3
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, k
                Next
                atkingckai(87, 2) = 0
                戰鬥系統類.執行動作_技能手動結束
             Else
                 atkingckai(87, 2) = 0
                戰鬥系統類.執行動作_技能手動結束
             End If
   End Select
End If
End Sub
Sub 艾蕾可_王座之炎()
Dim dge As Integer
If FormMainMode.comaiatk(1).Caption = "王座之炎" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(91, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾蕾可" Then
   Select Case atkingckai(91, 1)
        Case 1
             If atkingpagetot(2, 1) >= 5 And atkingckai(91, 2) = 0 Then
               atkingckai(91, 2) = 1
               atkingckai(91, 1) = 2
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + pageqlead(2) * 3
               atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 = pageqlead(2)
            ElseIf atkingpagetot(2, 4) < 2 And atkingckai(91, 2) = 1 Then
               atkingckai(91, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - pageqlead(2) * 3
               atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 = 0
            End If
        Case 2
                 If atkingpagetot(2, 1) < 5 Then
                     atkingckai(91, 2) = 0
                     atkingtrn(2) = Val(atkingtrn(2)) - 1
                     If pageqlead(2) = atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 Then
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - pageqlead(2) * 3
                     Else
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - pageqlead(2) * 3 - 3
                     End If
                     atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 = 0
                     atkingckai(91, 1) = 1
                  End If
                  If atkingckai(91, 2) = 1 Then
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (pageqlead(2) - Val(atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數)) * 3
                     atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 = pageqlead(2)
                  End If
        Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10725
                   atkingno(i, 7) = 91
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾蕾可\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            dge = Val(FormMainMode.pagecomglead.Caption)
            If dge > 4 Then dge = 4
            擲骰後骰傷害數 = Val(擲骰後骰傷害數) - dge
            擲骰表單溝通暫時變數(2) = 擲骰後骰傷害數
            atking_AI_艾蕾可_王座之炎計算出牌張數紀錄數 = 0
            atkingckai(91, 2) = 0
   End Select
End If
End Sub
Sub 艾蕾可_白百合()
Dim dge As Integer
If FormMainMode.comaiatk(2).Caption = "白百合" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(92, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾蕾可" Then
   Select Case atkingckai(92, 1)
        Case 1
             If movecp < 3 Then
                 If pageqlead(2) >= 2 And atkingckai(92, 2) = 0 Then
                   atkingckai(92, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If pageqlead(2) < 2 And atkingckai(92, 2) = 1 Then
                   atkingckai(92, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
             End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking2-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 92
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾蕾可\atking2-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(92, 2) = 0
            '===================
            If 擲骰後骰傷害數 > 0 Then
               戰鬥系統類.執行動作_清除所有異常狀態_使用者
               '==================
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
   End Select
End If
End Sub
Sub 艾蕾可_聖王威光()
Dim dge As Integer
If FormMainMode.comaiatk(3).Caption = "聖王威光" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(93, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾蕾可" Then
   Select Case atkingckai(93, 1)
        Case 1
             If atkingpagetot(2, 4) >= 3 And atkingckai(93, 2) = 0 Then
               atkingckai(93, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 3 And atkingckai(93, 2) = 1 Then
               atkingckai(93, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 93
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\艾蕾可\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atking_AI_艾蕾可_聖王威光紀錄數(1) = Val(FormMainMode.顯示列1.goi1)
             atking_AI_艾蕾可_聖王威光紀錄數(2) = pageqlead(1)
        Case 4
            atkingckai(93, 2) = 0
            '===================
            If 擲骰後骰傷害數 <= 0 Then
               dge = Int(atking_AI_艾蕾可_聖王威光紀錄數(1) / 4 + 0.9)
               戰鬥系統類.傷害執行_技能直傷_使用者 dge, 1
            End If
            '===================
            If atking_AI_艾蕾可_聖王威光紀錄數(2) = 0 Then
                戰鬥系統類.傷害執行_技能直傷_使用者 2, 1
            End If
            '===================
            Erase atking_AI_艾蕾可_聖王威光紀錄數
   End Select
End If
End Sub
Sub 艾蕾可_救濟天使()
Dim dge As Integer
If FormMainMode.comaiatk(4).Caption = "救濟天使" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(94, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾蕾可" Then
   Select Case atkingckai(94, 1)
        Case 1
             If atkingpagetot(2, 4) >= 5 And atkingckai(94, 2) = 0 Then
               atkingckai(94, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 5 And atkingckai(94, 2) = 1 Then
               atkingckai(94, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾蕾可\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10590
                   atkingno(i, 7) = 94
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(94, 2) = 0
            '===================
            If livecom(角色待機人物紀錄數(2, 2)) = 0 And livecom(角色待機人物紀錄數(2, 3)) = 0 Then
                Do
                    For j = 14 * (角色待機人物紀錄數(2, 1) - 1) + 1 To 14 * 角色待機人物紀錄數(2, 1)
                          If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 7
                              FormMainMode.personcomspe(j).person_turn = 4
                              人物異常狀態資料庫(2, j, 1) = 7
                              人物異常狀態資料庫(2, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色待機人物紀錄數(2, 1) - 1) + 1 To 14 * 角色待機人物紀錄數(2, 1)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 7, 4
                          異常狀態檢查數(1, 1) = 1
                          異常狀態檢查數(1, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '=================================
                Do
                    For j = 14 * (角色待機人物紀錄數(2, 1) - 1) + 1 To 14 * 角色待機人物紀錄數(2, 1)
                          If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 7
                              FormMainMode.personcomspe(j).person_turn = 4
                              人物異常狀態資料庫(2, j, 1) = 7
                              人物異常狀態資料庫(2, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色待機人物紀錄數(2, 1) - 1) + 1 To 14 * 角色待機人物紀錄數(2, 1)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 2, app_path & "gif\異常狀態\defup.gif", 7, 4
                          異常狀態檢查數(2, 1) = 1
                          異常狀態檢查數(2, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '================
                Do
                    For j = 14 * (角色待機人物紀錄數(2, 1) - 1) + 1 To 14 * 角色待機人物紀錄數(2, 1)
                          If 人物異常狀態資料庫(2, j, 3) = 38 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_turn = 4
                              人物異常狀態資料庫(2, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色待機人物紀錄數(2, 1) - 1) + 1 To 14 * 角色待機人物紀錄數(2, 1)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 38, app_path & "gif\異常狀態\再生.gif", 0, 4
                          異常狀態檢查數(38, 1) = 1
                          異常狀態檢查數(38, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '================
            Else
                '================
                For i = 2 To 3
                     If livecom(角色待機人物紀錄數(2, i)) > 0 Then
                        Do
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                                  If 人物異常狀態資料庫(2, j, 3) = 36 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_turn = 1
                                      人物異常狀態資料庫(2, j, 2) = 1
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                               If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, j, 36, app_path & "gif\異常狀態\庇護.png", 0, 1
                                  異常狀態檢查數(36, 1) = 1
                                  異常狀態檢查數(36, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        戰鬥系統類.回復執行_電腦 1, i
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub 露緹亞_腐朽之靈()
If FormMainMode.comaiatk(1).Caption = "腐朽之靈" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(95, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "露緹亞" Then
   Select Case atkingckai(95, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingckai(95, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(95, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 1) < 3 And atkingckai(95, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   atkingckai(95, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking1_2.jpg"
                   atkingno(i, 2) = 2
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
                For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                  If 人物異常狀態資料庫(1, i, 3) = 33 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                      FormMainMode.personusspe(i).person_turn = 3
                      人物異常狀態資料庫(1, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                   If 人物異常狀態資料庫(1, i, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 1, i, 33, app_path & "gif\異常狀態\咒縛.gif", 0, 3
                      異常狀態檢查數(33, 1) = 1
                      異常狀態檢查數(33, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            atkingckai(95, 2) = 0
   End Select
End If
End Sub
Sub 露緹亞_朦朧之暗()
If FormMainMode.comaiatk(2).Caption = "朦朧之暗" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(96, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "露緹亞" Then
   Select Case atkingckai(96, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(96, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(96, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(96, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   atkingckai(96, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 96
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(96, 2) = 0
            '=================
            戰鬥系統類.執行動作_距離變更 1
   End Select
End If
End Sub
Sub 露緹亞_暗影之翼()
If FormMainMode.comaiatk(3).Caption = "暗影之翼" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(97, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "露緹亞" Then
   Select Case atkingckai(97, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(97, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(97, 2) = 1
                End If
                If (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(97, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   atkingckai(97, 2) = 0
                 End If
          End If
      Case 2
           atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 97
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            atkingckai(97, 2) = 0
            '=================
            戰鬥系統類.執行動作_距離變更 3
            If Val(擲骰後骰傷害數) < 0 Then
                回復執行_電腦 1, 1
            End If
   End Select
End If
End Sub
Sub 露緹亞_渦騎劍閃(ByVal Index As Integer)
Dim aw As Integer
If FormMainMode.comaiatk(4).Caption = "渦騎劍閃" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(98, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "露緹亞" Then
   Select Case atkingckai(98, 1)
      Case 1
             If movecp = 3 Then
                 If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(98, 2) = 0 Then
'                     aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                     For i = 1 To 106
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                           aw = aw + 1
                        End If
                     Next
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_露緹亞_渦騎劍閃計算張數紀錄數) * 5 + 8
                     atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = aw
                     atkingckai(98, 2) = 1
                     atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
            End If
      Case 2
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 And atkingckai(98, 2) = 1 Then
                   For i = 1 To 106
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                           aw = aw + 1
                        End If
                   Next
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_露緹亞_渦騎劍閃計算張數紀錄數) * 5
                   atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = aw
            End If
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 And atkingckai(98, 2) = 1 Then
                   If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(98, 2) = 1 Then
                        For i = 1 To 106
                             If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                                aw = aw + 1
                             End If
                        Next
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_露緹亞_渦騎劍閃計算張數紀錄數) * 5
                        atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = aw
                   ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) And atkingckai(98, 2) = 1 Then
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - (atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 * 5) - 8
                        atkingckai(98, 2) = 0
                        atkingckai(98, 1) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) - 1
                        atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = 0
                    End If
            End If
'            formmainmode.trgoi2.Enabled = True
    Case 3
        If Val(pagecardnum(Index, 5)) = 1 And atkingckai(98, 2) = 1 Then
               If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 Then
                    For i = 1 To 106
                         If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                            aw = aw + 1
                         End If
                    Next
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + (aw - atking_AI_露緹亞_渦騎劍閃計算張數紀錄數) * 5
                    atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = aw
               ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) Then
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - (atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 * 5) - 8
                    atkingckai(98, 2) = 0
                    atkingckai(98, 1) = 1
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
                    atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = 0
                End If
        End If
'        formmainmode.trgoi2.Enabled = True
      Case 4
             atkingckai(98, 2) = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\露緹亞\atking4_2.jpg"
                   atkingno(i, 2) = 2
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
             atking_AI_露緹亞_渦騎劍閃計算張數紀錄數 = 0
   End Select
End If
End Sub
Sub 梅莉_夢幻魔杖()
Dim m As Integer, n As Integer, bd As Integer
If FormMainMode.comaiatk(1).Caption = "夢幻魔杖" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(99, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅莉" Then
   Select Case atkingckai(99, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 3 And atkingckai(99, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
                   atkingckai(99, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 5) < 3 And atkingckai(99, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   atkingckai(99, 2) = 0
                 End If
          End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9405
                   atkingno(i, 6) = 10245
                   atkingno(i, 7) = 99
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
            Randomize
            m = Int(Rnd() * 100) + 1
            If livecom(角色人物對戰人數(2, 2)) <= livecom41(角色人物對戰人數(2, 2)) Then
                Randomize
                bd = Int(Rnd() * 2) + 1
            End If
            If m Mod (2 - bd) = 0 Then '===相當於50~100%機率
                 Randomize
                 n = Int(Rnd() * 100) + 1
                 If livecom(角色人物對戰人數(2, 2)) <= livecommax(角色人物對戰人數(2, 2)) Then
                     bd = livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2))
                     If bd > 8 Then bd = 8
                 End If
                 If n Mod (10 - bd) = 0 Then '===相當於10~50%機率
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) * 4
                     FormMainMode.messageus.AddItem "夢幻魔杖效果發動!  攻擊力變為4倍"
                     戰鬥系統類.自動捲軸捲動
                Else
                     攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) * 2
                     FormMainMode.messageus.AddItem "夢幻魔杖效果發動!  攻擊力變為2倍"
                     戰鬥系統類.自動捲軸捲動
                End If
                FormMainMode.trgoi2_Timer
            End If
            atkingckai(99, 1) = 4
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(99, 1) = 5
        Case 5
             atkingckai(99, 2) = 0
             If Val(擲骰後骰傷害數) <= 0 Then
                 Do
                    For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                        If 人物異常狀態資料庫(1, j, 3) = 11 And 人物異常狀態資料庫(1, j, 2) > 0 Then
                         FormMainMode.personusspe(j).person_num = 3
                         FormMainMode.personusspe(j).person_turn = 3
                         人物異常狀態資料庫(1, j, 1) = 3
                         人物異常狀態資料庫(1, j, 2) = 3
                         Exit Do
                        End If
                    Next
                   For j = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, j, 2) = 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 1, j, 11, app_path & "gif\異常狀態\defdown.gif", 3, 3
                         異常狀態檢查數(11, 1) = 1
                         異常狀態檢查數(11, 2) = 1
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
If FormMainMode.comaiatk(2).Caption = "徬徨夢羽" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(100, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅莉" Then
   Select Case atkingckai(100, 1)
      Case 1
            If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 2) >= 3 And atkingckai(100, 2) = 0 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
               atkingckai(100, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 2) < 3) And atkingckai(100, 2) = 1 Then
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               atkingckai(100, 2) = 0
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8865
                   atkingno(i, 6) = 9210
                   atkingno(i, 7) = 100
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
            bd = 0
            Randomize
            m = Int(Rnd() * 100) + 1
            If livecom(角色人物對戰人數(2, 2)) <= livecom41(角色人物對戰人數(2, 2)) Then
                bd = 1
            End If
            If m Mod (3 - bd) = 0 Then '===相當於33~50%機率
                 Randomize
                 n = Int(Rnd() * 100) + 1
                 If livecom(角色人物對戰人數(2, 2)) <= livecommax(角色人物對戰人數(2, 2)) Then
                     bd = livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2))
                     If bd > 8 Then bd = 8
                 End If
                 If n Mod (10 - bd) = 0 Then '===相當於10~50%機率
                     攻擊防禦骰子總數(1) = 0
                     FormMainMode.messageus.AddItem "徬徨夢羽效果發動!  我方攻擊力變為0"
                     戰鬥系統類.自動捲軸捲動
                Else
                     攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) \ 2
                     FormMainMode.messageus.AddItem "徬徨夢羽效果發動!  我方攻擊力變為1/2"
                     戰鬥系統類.自動捲軸捲動
                End If
            Else
                攻擊防禦骰子總數(1) = Int((攻擊防禦骰子總數(1) * 2) / 3)
                FormMainMode.messageus.AddItem "徬徨夢羽效果發動!  我方攻擊力變為2/3"
                戰鬥系統類.自動捲軸捲動
            End If
            FormMainMode.trgoi1_Timer
            '=====================
            戰鬥系統類.回復執行_電腦 1, 1
            '=====================
            atkingckai(100, 1) = 4
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(100, 2) = 0
   End Select
End If
End Sub
Sub 梅莉_綿羊幻夢()
Dim bloodnum As Integer
If FormMainMode.comaiatk(3).Caption = "綿羊幻夢" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(101, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅莉" Then
   Select Case atkingckai(101, 1)
      Case 1
           If movecp < 3 Then
                If pageqlead(2) >= 2 And atkingckai(101, 2) = 0 Then
                    atkingckai(101, 2) = 1
                 End If
                 If pageqlead(2) < 2 And atkingckai(101, 2) = 1 Then
                    atkingckai(101, 2) = 0
                  End If
            End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6930
                   atkingno(i, 6) = 9540
                   atkingno(i, 7) = 101
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_AI_梅莉_綿羊幻夢_抽牌紀錄數
                   Select Case livecom(角色人物對戰人數(2, 2))
                       Case Is >= 5
                           atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(2) = 4
                           atkingno(i, 11) = 1
                       Case Else
                           atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(2) = 2
                           atkingno(i, 11) = 0
                    End Select
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(2) And atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(1) = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(1) = atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(1) > atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(2)
                    目前數(15) = 30
                    FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(1) > atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_AI_梅莉_綿羊幻夢_抽牌紀錄數(2)) <= 2 Then
                   戰鬥系統類.傷害執行_技能直傷_電腦 1, 1
                   atkingckai(101, 2) = 0
               Else
                   目前數(24) = 32
                   FormMainMode.等待時間_2.Enabled = True
               End If
            End If
        Case 5
            atkingckai(101, 2) = 0
            戰鬥系統類.傷害執行_技能直傷_電腦 1, 1
            戰鬥系統類.執行動作_技能手動結束
   End Select
End If
End Sub
Sub 梅莉_夢境搖籃()
If FormMainMode.comaiatk(4).Caption = "夢境搖籃" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(102, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅莉" Then
   Select Case atkingckai(102, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(102, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
                   atkingckai(102, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(102, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
                   atkingckai(102, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\梅莉\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7710
                   atkingno(i, 6) = 9030
                   atkingno(i, 7) = 102
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If livecom(角色人物對戰人數(2, 2)) > 2 Then
                 For i = 1 To 3
                     戰鬥系統類.傷害執行_技能直傷_使用者 1, i
                 Next
            Else
                 For i = 1 To 3
                     戰鬥系統類.傷害執行_技能直傷_使用者 4, i
                 Next
             End If
             atkingckai(102, 2) = 0
   End Select
End If
End Sub
Sub 古魯瓦爾多_必殺架勢()
Dim i As Integer, j As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "必殺架勢" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(104, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "古魯瓦爾多" Then
   Select Case atkingckai(104, 1)
        Case 1
            If atkingpagetot(2, 4) >= 2 And atkingpagetot(2, 3) = 0 And atkingckai(104, 2) = 0 Then
               atkingckai(104, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 4) < 2 Or atkingpagetot(2, 3) <> 0) And atkingckai(104, 2) = 1 Then
               atkingckai(104, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\古魯瓦爾多-必殺架勢2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 240
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9345
                   atkingno(i, 7) = 104
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
         Do
            atkingckai(104, 2) = 0
            For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                     FormMainMode.personcomspe(j).person_num = 5
                     FormMainMode.personcomspe(j).person_turn = 1
                     人物異常狀態資料庫(2, j, 1) = 5
                     人物異常狀態資料庫(2, j, 2) = 1
                     Exit Do
                 End If
            Next
           For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
              If 人物異常狀態資料庫(2, j, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 5, 1
                 異常狀態檢查數(1, 1) = 1
                 異常狀態檢查數(1, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub 古魯瓦爾多_精神力吸收()
Dim rrr(1 To 3) As Integer '牌判斷暫時變數
If FormMainMode.comaiatk(4).Caption = "精神力吸收" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(105, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "古魯瓦爾多" Then
   Select Case atkingckai(105, 1)
        Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
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
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(105, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(105,  2) = 0 Then
                atkingckai(105, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(105, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(105,  2) = 1 Then
                atkingckai(105, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
              End If
        Case 2
              For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\古魯瓦爾多\Grunwaldatking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1080
                   atkingno(i, 5) = 8025
                   atkingno(i, 6) = 9525
                   atkingno(i, 7) = 105
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
            Erase atking_AI_古魯瓦爾多_精神力吸收紀錄數
            '=====================
            For i = 1 To 106
                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                    If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                         atking_AI_古魯瓦爾多_精神力吸收紀錄數(i) = 1
                         atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) = atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) + 1
                     End If
                End If
            Next
            If atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) > 0 Then
                atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) = 0
                For i = 1 To 106
                    If atking_AI_古魯瓦爾多_精神力吸收紀錄數(i) = 1 Then
                        atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) = Val(atking_AI_古魯瓦爾多_精神力吸收紀錄數(0)) + 1
                        目前數(20) = i
                        目前數(21) = 8
                        atking_AI_古魯瓦爾多_精神力吸收紀錄數(i) = 0
                        FormMainMode.tr使用者牌_偷牌.Enabled = True
                        Exit Sub
                    End If
                Next
            Else
               目前數(22) = 31
               FormMainMode.等待時間.Enabled = True
            End If
        Case 4
'            FormMainMode.tr電腦牌_偷牌.Enabled = True
'            目前數(17) = 5
        Case 5
            If atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) > 0 Then
                For i = 1 To 106
                    If atking_AI_古魯瓦爾多_精神力吸收紀錄數(i) = 1 And atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) < 3 Then
                        atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) = Val(atking_AI_古魯瓦爾多_精神力吸收紀錄數(0)) + 1
                        目前數(20) = i
                        目前數(21) = 8
                        atking_AI_古魯瓦爾多_精神力吸收紀錄數(i) = 0
                        FormMainMode.tr使用者牌_偷牌.Enabled = True
                        Exit Sub
                    End If
                Next
                If i = 107 Then
                    atkingckai(105, 2) = 0
                    戰鬥系統類.執行動作_技能手動結束
                 End If
            Else
               目前數(22) = 31
               FormMainMode.等待時間.Enabled = True
            End If
        Case 6
            If atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) = 0 Then
                atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) = 99
               目前數(22) = 31
               FormMainMode.等待時間.Enabled = True
            ElseIf atking_AI_古魯瓦爾多_精神力吸收紀錄數(0) > 0 Then
               atkingckai(105, 2) = 0
               戰鬥系統類.執行動作_技能手動結束
            End If
   End Select
End If

End Sub
Sub 帕茉_憤怒之爪()
If FormMainMode.comaiatk(1).Caption = "憤怒之爪" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(106, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "帕茉" Then
   Select Case atkingckai(106, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(106, 2) = 0 Then
               atkingckai(106, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 4) < 1 And atkingckai(106, 2) = 1 Then
               atkingckai(106, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\帕茉\帕茉_憤怒之爪_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -600
                   atkingno(i, 4) = 1440
                   atkingno(i, 5) = 7050
                   atkingno(i, 6) = 9090
                   atkingno(i, 7) = 106
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
    Case 3
        Do
           atkingckai(106, 2) = 0
           For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
             If 人物異常狀態資料庫(2, i, 2) >= 9 And 人物異常狀態資料庫(2, i, 3) = 26 Then
                Exit Do
             End If
             If 人物異常狀態資料庫(2, i, 3) = 26 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 9 Then
                 FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) + 1
                 人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 1
                 Exit Do
             End If
           Next
           For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
              If 人物異常狀態資料庫(2, i, 2) = 0 Then
                 戰鬥系統類.人物異常狀態表設定_初設 2, i, 26, app_path & "gif\異常狀態\聖痕.gif", 0, 1
                 異常狀態檢查數(26, 1) = 1
                 異常狀態檢查數(26, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub 伊芙琳_怠惰的墓表()
Dim cardp(1 To 106) As Boolean '紀錄暫時變數
Dim cardpn As Integer '紀錄牌總數暫時變數
If FormMainMode.comaiatk(1).Caption = "怠惰的墓表" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(107, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "伊芙琳" Then
   Select Case atkingckai(107, 1)
        Case 1
           If movecp < 3 Then
                 If atkingpagetot(2, 4) >= 2 And atkingckai(107, 2) = 0 Then
'                 If pageqlead(2) >= 1 And atkingckai(107, 2) = 0 Then
                   atkingckai(107, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf atkingpagetot(2, 4) < 2 And atkingckai(107, 2) = 1 Then
'                ElseIf pageqlead(2) < 1 And atkingckai(107, 2) = 1 Then
                   atkingckai(107, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
           End If
        Case 2
              For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6345
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 107
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
            Do
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, i, 2) >= 9 And 人物異常狀態資料庫(2, i, 3) = 25 Then
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(2, i, 3) = 25 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 9 Then
                     FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2) + 1
                     人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 2, i, 25, app_path & "gif\異常狀態\能力低下.gif", 0, 1
                     異常狀態檢查數(25, 1) = 1
                     異常狀態檢查數(25, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
            '=====================
            cardpn = 0
            Erase cardp
            Erase atking_AI_伊芙琳_怠惰的墓表紀錄數
            '=====================
            Do
               Randomize
               i = Int(Rnd() * 106) + 1
               If cardp(i) = False Then
                    cardp(i) = True
                    cardpn = cardpn + 1
                    If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                      Select Case movecp
                         Case 1
                             If pagecardnum(i, 1) = a1a Or pagecardnum(i, 3) = a1a Then
                                  atking_AI_伊芙琳_怠惰的墓表紀錄數(atking_AI_伊芙琳_怠惰的墓表紀錄數(0) + 1) = i
                                  atking_AI_伊芙琳_怠惰的墓表紀錄數(0) = atking_AI_伊芙琳_怠惰的墓表紀錄數(0) + 1
                              End If
                         Case Is > 1
                             If pagecardnum(i, 1) = a5a Or pagecardnum(i, 3) = a5a Then
                                  atking_AI_伊芙琳_怠惰的墓表紀錄數(atking_AI_伊芙琳_怠惰的墓表紀錄數(0) + 1) = i
                                  atking_AI_伊芙琳_怠惰的墓表紀錄數(0) = atking_AI_伊芙琳_怠惰的墓表紀錄數(0) + 1
                              End If
                        End Select
                    End If
               End If
               If atking_AI_伊芙琳_怠惰的墓表紀錄數(0) >= 2 Then
                   Exit Do
               End If
            Loop While cardpn < 106
            If atking_AI_伊芙琳_怠惰的墓表紀錄數(0) > 0 Then
                目前數(20) = atking_AI_伊芙琳_怠惰的墓表紀錄數(1)
                atkingckai(107, 1) = 4
                目前數(21) = 9
                FormMainMode.tr使用者牌_偷牌.Enabled = True
            Else
'               atkingckai(107, 2) = 0
               atkingckai(107, 1) = 4
               目前數(22) = 32
               FormMainMode.等待時間.Enabled = True
            End If
        Case 4
             If atking_AI_伊芙琳_怠惰的墓表紀錄數(0) < 2 Then
                目前數(22) = 32
                FormMainMode.等待時間.Enabled = True
            Else
                目前數(20) = atking_AI_伊芙琳_怠惰的墓表紀錄數(2)
                atkingckai(107, 1) = 5
                目前數(21) = 9
                FormMainMode.tr使用者牌_偷牌.Enabled = True
            End If
        Case 5
            If atking_AI_伊芙琳_怠惰的墓表紀錄數(0) = 0 Then
               atking_AI_伊芙琳_怠惰的墓表紀錄數(0) = 3
               目前數(22) = 32
               FormMainMode.等待時間.Enabled = True
               Exit Sub
            ElseIf atking_AI_伊芙琳_怠惰的墓表紀錄數(0) = 2 Then
               目前數(24) = 34
               FormMainMode.等待時間_2.Enabled = True
               Exit Sub
            ElseIf atking_AI_伊芙琳_怠惰的墓表紀錄數(0) > 0 And atking_AI_伊芙琳_怠惰的墓表紀錄數(0) <> 2 Then
               atkingckai(107, 2) = 0
               執行動作_技能手動結束
            End If
        Case 6
            atkingckai(107, 2) = 0
            執行動作_技能手動結束
   End Select
End If

End Sub
Sub 伊芙琳_赤紅石榴()
Dim mkp As Integer '暫時變數
Dim cardp(1 To 106) As Boolean '紀錄暫時變數
Dim cardpn(1 To 2) As Integer '紀錄牌總數暫時變數(1.牌紀錄目前總數/2.牌選定目前總數)
If FormMainMode.comaiatk(4).Caption = "赤紅石榴" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(108, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "伊芙琳" Then
   If Formsetting.checktest.Value = 1 Then Debug.Print "經過赤紅石榴主名字判斷"
   Select Case atkingckai(108, 1)
        Case 1
            If movecp = 3 Then
                 If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 _
                    And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 5) >= 1 And atkingckai(108, 2) = 0 Then
'                 If atkingpagetot(2, 3) >= 1 And atkingckai(108, 2) = 0 Then
                   atkingckai(108, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1 _
                   Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 5) < 1) And atkingckai(108, 2) = 1 Then
'                ElseIf atkingpagetot(2, 3) < 1 And atkingckai(108, 2) = 1 Then
                   atkingckai(108, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
          End If
        Case 2
             '==================================
             Erase atking_AI_伊芙琳_赤紅石榴階段紀錄數
             Randomize
             mkp = Int(Rnd() * 16) + 1
             atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = mkp
             If Formsetting.checktest.Value = 1 Then Debug.Print "技能 - AI - 伊芙琳 - 赤紅石榴效果值" & mkp
             '===================================
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\伊芙琳\Evelynatking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9330
                   atkingno(i, 6) = 9165
                   atkingno(i, 7) = 108
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                    If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) <= 9 Or atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) >= 13 Then
                       atkingno(i, 11) = 0
                    Else
                       atkingno(i, 11) = 1
                    End If
                   Exit For
                 End If
             Next
        Case 3
            '======================
               戰鬥系統類.執行動作_清除所有異常狀態_電腦
            '======================
            Select Case atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1)
                Case 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 1, 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 1, 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 1, 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 1, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 1, 2
                    戰鬥系統類.傷害執行_技能直傷_電腦 1, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員1點傷害。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                Case 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 3, 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, 2
                    戰鬥系統類.傷害執行_技能直傷_電腦 3, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員3點傷害。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                Case 3
                    戰鬥系統類.傷害執行_技能直傷_使用者 5, 1
                    戰鬥系統類.傷害執行_技能直傷_使用者 5, 2
                    戰鬥系統類.傷害執行_技能直傷_使用者 5, 3
                    戰鬥系統類.傷害執行_技能直傷_電腦 5, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 5, 2
                    戰鬥系統類.傷害執行_技能直傷_電腦 5, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員5點傷害。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
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
                    atkingckai(108, 2) = 0
                Case 5
                    回復執行_使用者 3, 1
                    回復執行_使用者 3, 2
                    回復執行_使用者 3, 3
                    回復執行_電腦 3, 1
                    回復執行_電腦 3, 2
                    回復執行_電腦 3, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員HP回復3點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                Case 6
                    回復執行_使用者 5, 1
                    回復執行_使用者 5, 2
                    回復執行_使用者 5, 3
                    回復執行_電腦 5, 1
                    回復執行_電腦 5, 2
                    回復執行_電腦 5, 3
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己隊伍與對方隊伍全員HP回復5點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                '===============================================
                Case 7
                    戰鬥系統類.傷害執行_技能直傷_使用者 Val(liveus(角色人物對戰人數(1, 2))) - 1, 1
                    戰鬥系統類.傷害執行_技能直傷_電腦 Val(livecom(角色人物對戰人數(2, 2))) - 1, 1
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己與對方的HP變為1點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                '============================================
                Case 8
                    If Val(liveus(角色人物對戰人數(1, 2))) > 5 Then
                        戰鬥系統類.傷害執行_技能直傷_使用者 Val(liveus(角色人物對戰人數(1, 1))) - 5, 1
                    Else
                        回復執行_使用者 5 - Val(liveus(角色人物對戰人數(1, 1))), 1
                    End If
                    If Val(livecom(角色人物對戰人數(2, 2))) > 5 Then
                        戰鬥系統類.傷害執行_技能直傷_電腦 Val(livecom(角色人物對戰人數(2, 2))) - 5, 1
                    Else
                        回復執行_電腦 5 - Val(livecom(角色人物對戰人數(2, 2))), 1
                    End If
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己與對方的HP變為5點。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                '===============================================
                Case 9
                    回復執行_使用者 Val(liveusmax(角色人物對戰人數(1, 2))) - Val(liveus(角色人物對戰人數(1, 2))), 1
                    回復執行_電腦 Val(livecommax(角色人物對戰人數(2, 2))) - Val(livecom(角色人物對戰人數(2, 2))), 1
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  自己與對方的HP完全恢復。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                '===============================================
                Case 10
                    目前數(20) = 1
                    atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1
                '==========使用者棄牌階段
                    Do
                        If Val(pagecardnum(目前數(20), 5)) = 1 And Val(pagecardnum(目前數(20), 6)) = 1 Then
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                            目前數(21) = 10
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
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pageusglead) - 8 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(21) = 10
                                FormMainMode.tr使用者_棄牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(20) = 目前數(20) + 1
                        Loop Until 目前數(20) > 106
                    ElseIf Val(FormMainMode.pageusglead) < 8 Then
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 3
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 8
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        目前數(15) = 33
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
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pageusglead) - 15 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(21) = 10
                                FormMainMode.tr使用者_棄牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(20) = 目前數(20) + 1
                        Loop Until 目前數(20) > 106
                    ElseIf Val(FormMainMode.pageusglead) < 15 Then
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 3
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 15
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        目前數(15) = 33
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
                    atkingckai(108, 2) = 0
                Case 14
                    執行動作_距離變更 (2)
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  距離變為中距離。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                Case 15
                    執行動作_距離變更 (3)
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  距離變為遠距離。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
                Case 16
                    FormMainMode.messageus.AddItem "赤紅石榴發動!  什麼都沒有發生。"
                    戰鬥系統類.自動捲軸捲動
                    atkingckai(108, 2) = 0
            End Select
        '=====================================================
       Case 4
             If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 10 Then
                    '==========使用者棄牌階段2
                    If 目前數(20) <= 106 Then
                        Do
                            If Val(pagecardnum(目前數(20), 5)) = 1 And Val(pagecardnum(目前數(20), 6)) = 1 Then
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(21) = 10
                                FormMainMode.tr使用者_棄牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(20) = 目前數(20) + 1
                        Loop Until 目前數(20) > 106
                    End If
效果10_使用者棄牌階段直接跳過:
                    If 目前數(20) > 106 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 1 Then
                        目前數(16) = 1
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2
                        '=============電腦方棄牌階段1
                        Do
                            If Val(pagecardnum(目前數(16), 5)) = 2 And Val(pagecardnum(目前數(16), 6)) = 1 Then
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(17) = 12
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                        If 目前數(16) > 106 Then
'                            FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為0張。"
'                            戰鬥系統類.自動捲軸捲動
'                            atkingckai(108, 2) = 0
'                            執行動作_技能手動結束
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                                GoTo 效果結束實行_手牌變化類
                            Else
                                目前數(22) = 33
                                FormMainMode.等待時間.Enabled = True
                            End If
                        End If
                    End If
                    If 目前數(16) <= 106 And atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2 Then
                        '==============電腦方棄牌階段2
                        Do
                            If Val(pagecardnum(目前數(16), 5)) = 2 And Val(pagecardnum(目前數(16), 6)) = 1 Then
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                目前數(17) = 12
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                        If 目前數(16) > 106 Then
'                            FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為0張。"
'                            戰鬥系統類.自動捲軸捲動
'                            atkingckai(108, 2) = 0
'                            執行動作_技能手動結束
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
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
            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
                    Do
                        If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                            目前數(21) = 10
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                        目前數(20) = 目前數(20) + 1
                    Loop Until 目前數(20) > 106
                    '=========電腦牌數判斷及選擇
效果11_移至電腦判斷:
                    目前數(16) = 1
                    If Val(FormMainMode.pagecomglead) > 8 Then
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2
                        Erase cardp
                        Erase cardpn
                        For i = 1 To 106
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 0
                        Next
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pagecomglead) - 8 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                                目前數(17) = 12
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                    ElseIf Val(FormMainMode.pagecomglead) < 8 Then
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 4
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 8
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        If Val(FormMainMode.pageul) < 8 - Val(FormMainMode.pagecomglead) Then
                            戰鬥系統類.執行動作_洗牌
                        End If
                        目前數(15) = 33
                        FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Else
'                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為8張。"
'                        戰鬥系統類.自動捲軸捲動
'                        atkingckai(108, 2) = 0
'                        執行動作_技能手動結束
                        If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                            GoTo 效果結束實行_手牌變化類
                        Else
                            目前數(22) = 33
                            FormMainMode.等待時間.Enabled = True
                        End If
                    End If
            End If
        '=====================================================
        Case 7
            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
                 Do
                     If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                         目前數(17) = 12
                         atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                         atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                         FormMainMode.tr電腦牌_翻牌.Enabled = True
                         Exit Sub
                     End If
                     目前數(16) = 目前數(16) + 1
                 Loop Until 目前數(16) > 106
                 If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
'                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為8張。"
'                        戰鬥系統類.自動捲軸捲動
'                        atkingckai(108, 2) = 0
'                        執行動作_技能手動結束
                        GoTo 效果結束實行_手牌變化類
                 Else
                        目前數(22) = 33
                        FormMainMode.等待時間.Enabled = True
                 End If
             End If
        '=====================================================
        Case 8
            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 11 Then
                Select Case atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2)
                    Case 3
                        If Val(FormMainMode.pageusglead) < atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 33
                           atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                        End If
                        If Val(FormMainMode.pageusglead) >= atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           GoTo 效果11_移至電腦判斷
                        End If
                    Case 4
                        If Val(FormMainMode.pagecomglead) < atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 33
                           atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                        End If
                        If Val(FormMainMode.pagecomglead) >= atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
'                           FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為8張。"
'                           戰鬥系統類.自動捲軸捲動
'                            atkingckai(108, 2) = 0
'                            執行動作_技能手動結束
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
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
            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
                    Do
                        If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 1 Then
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(20), 1) = 0
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                            目前數(21) = 10
                            FormMainMode.tr使用者_棄牌.Enabled = True
                            Exit Sub
                        End If
                        目前數(20) = 目前數(20) + 1
                    Loop Until 目前數(20) > 106
                    '=========電腦牌數判斷及選擇
效果12_移至電腦判斷:
                    目前數(16) = 1
                    If Val(FormMainMode.pagecomglead) > 15 Then
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 2
                        Erase cardp
                        Erase cardpn
                        For i = 1 To 106
                            atking_AI_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 0
                        Next
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_伊芙琳_赤紅石榴階段紀錄數(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pagecomglead) - 15 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                                目前數(17) = 12
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                                atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                                FormMainMode.tr電腦牌_翻牌.Enabled = True
                                Exit Sub
                            End If
                            目前數(16) = 目前數(16) + 1
                        Loop Until 目前數(16) > 106
                    ElseIf Val(FormMainMode.pagecomglead) < 15 Then
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2) = 4
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) = 15
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        If Val(FormMainMode.pageul) < 15 - Val(FormMainMode.pagecomglead) Then
                            戰鬥系統類.執行動作_洗牌
                        End If
                        目前數(15) = 33
                        FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Else
'                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為15張。"
'                        戰鬥系統類.自動捲軸捲動
'                        atkingckai(108, 2) = 0
'                        執行動作_技能手動結束
                        If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                            GoTo 效果結束實行_手牌變化類
                        Else
                            目前數(22) = 33
                            FormMainMode.等待時間.Enabled = True
                        End If
                    End If
            End If
        '=====================================================
        Case 10
            If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
                Do
                    If atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 1 Then
                        目前數(17) = 12
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(目前數(16), 1) = 0
                        atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                        FormMainMode.tr電腦牌_翻牌.Enabled = True
                        Exit Sub
                    End If
                    目前數(16) = 目前數(16) + 1
                Loop Until 目前數(16) > 106
'                FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為15張。"
'                戰鬥系統類.自動捲軸捲動
'                atkingckai(108, 2) = 0
'                執行動作_技能手動結束
                If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
                    GoTo 效果結束實行_手牌變化類
                Else
                    目前數(22) = 33
                    FormMainMode.等待時間.Enabled = True
                End If
            End If
        '=====================================================
       Case 11
           If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1) = 12 Then
                Select Case atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 2)
                    Case 3
                        If Val(FormMainMode.pageusglead) < atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 33
                           atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_使用者.Enabled = True
                        End If
                        If Val(FormMainMode.pageusglead) >= atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           GoTo 效果12_移至電腦判斷
                        End If
                    Case 4
                        If Val(FormMainMode.pagecomglead) < atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           目前數(15) = 33
                           atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) = atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) + 1
                           FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                        End If
                        If Val(FormMainMode.pagecomglead) >= atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
'                           FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為15張。"
'                           戰鬥系統類.自動捲軸捲動
'                           atkingckai(108, 2) = 0
'                           執行動作_技能手動結束
                           If atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 4) >= 2 Then
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
            Select Case atking_AI_伊芙琳_赤紅石榴階段紀錄數(0, 1)
                 Case 10
                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為0張。"
                        戰鬥系統類.自動捲軸捲動
                        atkingckai(108, 2) = 0
                        執行動作_技能手動結束
                 Case 11
                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為8張。"
                        戰鬥系統類.自動捲軸捲動
                        atkingckai(108, 2) = 0
                        執行動作_技能手動結束
                 Case 12
                        FormMainMode.messageus.AddItem "赤紅石榴發動! 自己與對方的手牌變為15張。"
                        戰鬥系統類.自動捲軸捲動
                        atkingckai(108, 2) = 0
                        執行動作_技能手動結束
            End Select
   End Select
End If
End Sub
Sub 布勞_發條機構()
Dim tn As Integer
If FormMainMode.comaiatk(1).Caption = "發條機構" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(109, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "布勞" Then
   Select Case atkingckai(109, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 4) >= 2 And atkingckai(109, 2) = 0 Then
                   atkingckai(109, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 2 And atkingckai(109, 2) = 1 Then
                   atkingckai(109, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7455
                   atkingno(i, 6) = 9075
                   atkingno(i, 7) = 109
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
                    If tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgrecom.Value = 0 Then
                        If pageeventnum(2, tn, 1) <> "" Then
                            ay = Split(一般系統類.事件卡資料庫(pageeventnum(2, tn, 1), 3), "=")
                            pagecardnum(88 + tn, 1) = ay(0)
                            pagecardnum(88 + tn, 2) = ay(1)
                            pagecardnum(88 + tn, 3) = ay(2)
                            pagecardnum(88 + tn, 4) = ay(3)
                            pagecardnum(88 + tn, 5) = 2
                            pagecardnum(88 + tn, 6) = 1
                            pagecardnum(88 + tn, 8) = pageeventnum(2, tn, 2)
                            FormMainMode.card(88 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(2, tn, 2) & "-1.bmp")
                            pagecardnum(88 + tn, 11) = 0
                            pageonin(88 + tn) = 1
                        End If
                    End If
                End If
             '=====================================
             If Val(FormMainMode.turni) < 18 And (tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgrecom.Value = 0) Then
                目前數(16) = 88 + Val(FormMainMode.turni) + 1
                atking_AI_布勞_發條機構紀錄數 = 1
                目前數(15) = 34
                FormMainMode.tr牌組_回牌_電腦.Enabled = True
            Else
                atkingckai(109, 2) = 0
            End If
        Case 4
            If Val(FormMainMode.turni) + atking_AI_布勞_發條機構紀錄數 < 18 And atking_AI_布勞_發條機構紀錄數 < 2 And _
               (tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgrecom.Value = 0) Then
                '=====================================分派下一張事件卡
                tn = Val(FormMainMode.turni) + 2
                If tn <= 18 Then
                        If tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgrecom.Value = 0 Then
                            If pageeventnum(2, tn, 1) <> "" Then
                                ay = Split(一般系統類.事件卡資料庫(pageeventnum(2, tn, 1), 3), "=")
                                pagecardnum(88 + tn, 1) = ay(0)
                                pagecardnum(88 + tn, 2) = ay(1)
                                pagecardnum(88 + tn, 3) = ay(2)
                                pagecardnum(88 + tn, 4) = ay(3)
                                pagecardnum(88 + tn, 5) = 2
                                pagecardnum(88 + tn, 6) = 1
                                pagecardnum(88 + tn, 8) = pageeventnum(2, tn, 2)
                                FormMainMode.card(88 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(2, tn, 2) & "-1.bmp")
                                pagecardnum(88 + tn, 11) = 0
                                pageonin(88 + tn) = 1
                            End If
                        End If
                End If
                atking_AI_布勞_發條機構紀錄數 = atking_AI_布勞_發條機構紀錄數 + 1
                目前數(16) = 88 + Val(FormMainMode.turni) + 2
                目前數(15) = 34
                FormMainMode.tr牌組_回牌_電腦.Enabled = True
            Else
                FormMainMode.turni = Val(FormMainMode.turni) + atking_AI_布勞_發條機構紀錄數
                turn = Val(FormMainMode.turni)
                atking_AI_布勞_發條機構紀錄數 = 0
                atkingckai(109, 2) = 0
            End If
   End Select
End If
End Sub
Sub 布勞_夜幕時分()
Dim tn(1 To 3) As Boolean
If FormMainMode.comaiatk(4).Caption = "夜幕時分" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(110, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "布勞" Then
   Select Case atkingckai(110, 1)
        Case 1
             If pageqlead(2) >= 3 And atkingckai(110, 2) = 0 Then
               atkingckai(110, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf pageqlead(2) < 3 And atkingckai(110, 2) = 1 Then
               atkingckai(110, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\布勞\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -840
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6870
                   atkingno(i, 6) = 10365
                   atkingno(i, 7) = 110
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
            atkingckai(110, 2) = 0
            '======================
            For i = 1 To 3
                If VBEPerson(2, 角色待機人物紀錄數(2, i), 1, 2, 1) = "R" Then
                     tn(i) = True
                Else
                     tn(i) = False
                End If
                 If tn(i) = True Then
                     Select Case Val(VBEPerson(2, 角色待機人物紀錄數(2, i), 1, 2, 2))
                         Case Is <= 2
                              戰鬥系統類.回復執行_電腦 1, i
                         Case Is > 2, Is <= 4
                              戰鬥系統類.回復執行_電腦 2, i
                         Case 5
                              戰鬥系統類.回復執行_電腦 3, i
                     End Select
                 End If
            Next
            '=============================
   End Select
End If
End Sub
Sub 阿貝爾_抽刀斷水計()
If FormMainMode.comaiatk(4).Caption = "抽刀斷水計" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(113, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿貝爾" Then
   Select Case atkingckai(113, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(113, 2) = 0 Then
                   atkingckai(113, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(113, 2) = 1 Then
                   atkingckai(113, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\阿貝爾\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 8745
                   atkingno(i, 6) = 10200
                   atkingno(i, 7) = 113
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(113, 2) = 0
             '======================
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
   End Select
End If
End Sub

Sub 夏洛特_夜未央()
If FormMainMode.comaiatk(3).Caption = "夜未央" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(114, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "夏洛特" Then
   Select Case atkingckai(114, 1)
      Case 1
            If movecp < 3 Then
                If atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(114, 2) = 0 Then
                   atkingckai(114, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(114, 2) = 1 Then
                   atkingckai(114, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
            End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\夏洛特\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7575
                   atkingno(i, 6) = 9660
                   atkingno(i, 7) = 114
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(114, 2) = 0
             '========================
             戰鬥系統類.回復執行_電腦 1, 1
             '========================
             For i = 18 To (turn + 3) Step -1
                  pageeventnum(2, i, 1) = pageeventnum(2, i - 2, 1)
                  pageeventnum(2, i, 2) = pageeventnum(2, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 2)
                  pageeventnum(2, i, 1) = "HP回復3"
                  pageeventnum(2, i, 2) = 一般系統類.事件卡資料庫("HP回復3", 2)
             Next
   End Select
End If
End Sub
Sub 瑪格莉特_末日幻影()
Dim m As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "末日幻影" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(116, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "瑪格莉特" Then
   Select Case atkingckai(116, 1)
        Case 1
            If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 3) = 0 And atkingckai(116, 2) = 0 Then
'            If atkingpagetot(2, 3) >= 1 And atkingckai(116, 2) = 0 Then
               atkingckai(116, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 3) > 0) And atkingckai(116, 2) = 1 Then
'            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(116, 2) = 1 Then
               atkingckai(116, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\瑪格莉特\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 10455
                   atkingno(i, 7) = 116
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
                atkingckai(116, 2) = 0
                Select Case movecp
                    Case 1
                       Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 29 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 3
                                  人物異常狀態資料庫(1, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 29, app_path & "gif\異常狀態\恐怖.gif", 0, 3
                                  異常狀態檢查數(29, 1) = 1
                                  異常狀態檢查數(29, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 20 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 3
                                  人物異常狀態資料庫(1, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 20, app_path & "gif\異常狀態\damage.gif", 0, 3
                                  異常狀態檢查數(20, 1) = 1
                                  異常狀態檢查數(20, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                              If 人物異常狀態資料庫(1, i, 3) = 27 And 人物異常狀態資料庫(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 3
                                  人物異常狀態資料庫(1, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                               If 人物異常狀態資料庫(1, i, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 1, i, 27, app_path & "gif\異常狀態\狂戰士.gif", 0, 3
                                  異常狀態檢查數(27, 1) = 1
                                  異常狀態檢查數(27, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
   End Select
End If
End Sub
Sub 蕾格烈芙_SSS()
Dim rrr(1 To 3) As Integer '暫時變數
If FormMainMode.comaiatk(4).Caption = "S.S.S" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(117, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾格烈芙" Then
   Select Case atkingckai(117, 1)
        Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
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
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(117, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(117, 2) = 0 Then
                atkingckai(117, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(117, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(117, 2) = 1 Then
                atkingckai(117, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
              End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾格烈芙\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6135
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 117
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
                atkingckai(117, 2) = 0
                Do
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                          If 人物異常狀態資料庫(2, j, 3) = 32 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_turn = 3
                              人物異常狀態資料庫(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                       If 人物異常狀態資料庫(2, j, 2) = 0 Then
                          戰鬥系統類.人物異常狀態表設定_初設 2, j, 32, app_path & "gif\異常狀態\混沌.gif", 0, 3
                          異常狀態檢查數(32, 1) = 1
                          異常狀態檢查數(32, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
   End Select
End If
End Sub
Sub 多妮妲_超級女主角()
If FormMainMode.comaiatk(3).Caption = "超級女主角" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(118, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "多妮妲" Then
   Select Case atkingckai(118, 1)
      Case 1
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingpagetot(2, 4) >= 2 And atkingckai(118, 2) = 0 Then
               atkingckai(118, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3 Or atkingpagetot(2, 4) < 2) And atkingckai(118, 2) = 1 Then
               atkingckai(118, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\多妮妲\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 5970
                   atkingno(i, 6) = 10365
                   atkingno(i, 7) = 118
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(118, 2) = 0
            '==================
            Do
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                          FormMainMode.personcomspe(j).person_num = 6
                          FormMainMode.personcomspe(j).person_turn = 5
                          人物異常狀態資料庫(2, j, 1) = 6
                          人物異常狀態資料庫(2, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                   If 人物異常狀態資料庫(2, j, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 6, 5
                      異常狀態檢查數(1, 1) = 1
                      異常狀態檢查數(1, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
            Do
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                          FormMainMode.personcomspe(j).person_num = 4
                          FormMainMode.personcomspe(j).person_turn = 5
                          人物異常狀態資料庫(2, j, 1) = 4
                          人物異常狀態資料庫(2, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                   If 人物異常狀態資料庫(2, j, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 2, j, 2, app_path & "gif\異常狀態\defup.gif", 4, 5
                      異常狀態檢查數(2, 1) = 1
                      異常狀態檢查數(2, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
            Do
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, j, 3) = 3 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                          FormMainMode.personcomspe(j).person_num = 1
                          FormMainMode.personcomspe(j).person_turn = 5
                          人物異常狀態資料庫(2, j, 1) = 1
                          人物異常狀態資料庫(2, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                   If 人物異常狀態資料庫(2, j, 2) = 0 Then
                      戰鬥系統類.人物異常狀態表設定_初設 2, j, 3, app_path & "gif\異常狀態\movup.gif", 1, 5
                      異常狀態檢查數(3, 1) = 1
                      異常狀態檢查數(3, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
   End Select
End If
End Sub
Sub 傑多_因果之線()
Dim m As Integer
If FormMainMode.comaiatk(1).Caption = "因果之線" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(119, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "傑多" Then
   Select Case atkingckai(119, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(119, 2) = 0 Then
               atkingckai(119, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 1 And atkingckai(119, 2) = 1 Then
               atkingckai(119, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9255
                   atkingno(i, 7) = 119
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(119, 2) = 0
             Do
                Randomize
                m = Int(Rnd() * 106) + 1
                If Val(pagecardnum(m, 6)) = 1 And Val(pagecardnum(m, 5)) = 1 Then
                     目前數(20) = m
                     目前數(21) = 1
                     FormMainMode.tr使用者牌_偷牌.Enabled = True
                     Exit Do
                End If
            Loop
   End Select
End If
End Sub
Sub 傑多_因果之輪()
Dim m, n As Integer
Dim aw(1 To 2) As Integer
If FormMainMode.comaiatk(2).Caption = "因果之輪" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(120, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "傑多" Then
   Select Case atkingckai(120, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(120, 2) = 0 Then
               atkingckai(120, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(120, 2) = 1 Then
               atkingckai(120, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7260
                   atkingno(i, 6) = 8925
                   atkingno(i, 7) = 120
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
                階段狀態數 = 1
                For m = 1 To 106
                    If Val(pagecardnum(m, 6)) = 2 And Val(pagecardnum(m, 5)) = 1 Then
                        Randomize
                        n = Int(Rnd() * 6) + 1
                        If n Mod 2 = 0 Then
                            FormMainMode.cqen_Click (m)
                        End If
                    End If
                Next
              atkingckai(120, 1) = 4
              FormMainMode.trgoi1_Timer
              FormMainMode.trgoi2_Timer
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(120, 2) = 0
   End Select
End If
End Sub
Sub 傑多_因果之刻()
Dim m As Integer
If FormMainMode.comaiatk(3).Caption = "因果之刻" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(121, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "傑多" Then
   Select Case atkingckai(121, 1)
      Case 1
            If atkingpagetot(2, 4) >= 4 And atkingckai(121, 2) = 0 Then
               atkingckai(121, 2) = 1
            End If
            If atkingpagetot(2, 4) < 4 And atkingckai(121, 2) = 1 Then
               atkingckai(121, 2) = 0
             End If
      Case 2
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_傑多_因果之刻記錄數(i) = 1
                   atking_AI_傑多_因果之刻記錄數(107) = atking_AI_傑多_因果之刻記錄數(107) + 1
               End If
            Next
            atking_AI_傑多_因果之刻記錄數(108) = 1
      Case 3
            atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 4
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\傑多\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7425
                   atkingno(i, 6) = 9570
                   atkingno(i, 7) = 121
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
            Do Until atking_AI_傑多_因果之刻記錄數(108) > 106
                If atking_AI_傑多_因果之刻記錄數(atking_AI_傑多_因果之刻記錄數(108)) = 1 Then
                    目前數(16) = atking_AI_傑多_因果之刻記錄數(108)
                    目前數(15) = 35
                    FormMainMode.tr牌組_回牌_電腦.Enabled = True
                    atking_AI_傑多_因果之刻記錄數(目前數(16)) = 0
                    Exit Do
                End If
                atking_AI_傑多_因果之刻記錄數(108) = atking_AI_傑多_因果之刻記錄數(108) + 1
            Loop
            If atking_AI_傑多_因果之刻記錄數(108) >= 106 Then
                If atking_AI_傑多_因果之刻記錄數(107) < 2 Then
                    atking_AI_傑多_因果之刻記錄數(107) = atking_AI_傑多_因果之刻記錄數(107) + 1
                    目前數(22) = 34
                    FormMainMode.等待時間.Enabled = True
                ElseIf atking_AI_傑多_因果之刻記錄數(107) >= 2 Then
                    atkingckai(121, 2) = 0
                    Erase atking_AI_傑多_因果之刻記錄數
                    戰鬥系統類.執行動作_技能手動結束
                End If
            End If
   End Select
End If
End Sub

Sub 貝琳達_雪光()
Dim bloodnum As Integer
If FormMainMode.comaiatk(1).Caption = "雪光" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(122, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "貝琳達" Then
   Select Case atkingckai(122, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(122, 2) = 0 Then
                atkingckai(122, 2) = 1
             End If
             If atkingpagetot(2, 4) < 2 And atkingckai(122, 2) = 1 Then
                atkingckai(122, 2) = 0
              End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7035
                   atkingno(i, 6) = 9510
                   atkingno(i, 7) = 122
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_AI_貝琳達_雪光_抽牌紀錄數
                   Select Case liveus(角色人物對戰人數(1, 2))
                       Case Is = liveusmax(角色人物對戰人數(1, 2))
                           atking_AI_貝琳達_雪光_抽牌紀錄數(2) = 4
                           atkingno(i, 11) = 1
                       Case Else
                           atking_AI_貝琳達_雪光_抽牌紀錄數(2) = 2
                           atkingno(i, 11) = 0
                    End Select
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_AI_貝琳達_雪光_抽牌紀錄數(2) And atking_AI_貝琳達_雪光_抽牌紀錄數(1) = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_AI_貝琳達_雪光_抽牌紀錄數(1) = atking_AI_貝琳達_雪光_抽牌紀錄數(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_貝琳達_雪光_抽牌紀錄數(1) > atking_AI_貝琳達_雪光_抽牌紀錄數(2)
                    目前數(15) = 36
                    FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_貝琳達_雪光_抽牌紀錄數(1) > atking_AI_貝琳達_雪光_抽牌紀錄數(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_AI_貝琳達_雪光_抽牌紀錄數(2)) <= 2 Then
                   atkingckai(122, 2) = 0
               Else
                   目前數(24) = 35
                   FormMainMode.等待時間_2.Enabled = True
               End If
            End If
        Case 5
            atkingckai(122, 2) = 0
            戰鬥系統類.執行動作_技能手動結束
   End Select
End If
End Sub
Sub 貝琳達_水晶幻鏡()
If FormMainMode.comaiatk(2).Caption = "水晶幻鏡" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(123, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "貝琳達" Then
   Select Case atkingckai(123, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(2, 2) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(123, 2) = 0 Then
               atkingckai(123, 2) = 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
            ElseIf (atkingpagetot(2, 2) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(123, 2) = 1 Then
               atkingckai(123, 2) = 0
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
            End If
          End If
        Case 2
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_貝琳達_水晶幻鏡紀錄狀態數(i) = True
               End If
            Next
            目前數(30) = 1
        Case 3
            atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 4
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7290
                   atkingno(i, 6) = 9120
                   atkingno(i, 7) = 123
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
            Do
                If atking_AI_貝琳達_水晶幻鏡紀錄狀態數(目前數(30)) = True Then
                    目前數(16) = 目前數(30)
                    目前數(15) = 37
                    FormMainMode.tr牌組_回牌_電腦.Enabled = True
                    atking_AI_貝琳達_水晶幻鏡紀錄狀態數(目前數(16)) = False
                    Exit Do
                End If
                目前數(30) = 目前數(30) + 1
            Loop Until 目前數(30) >= 106
            If 目前數(30) >= 106 Then
                If 目前數(30) < 2 Then
                    目前數(30) = 目前數(30) + 1
                    目前數(22) = 35
                    FormMainMode.等待時間.Enabled = True
                ElseIf 目前數(30) >= 2 Then
                    atkingckai(123, 2) = 0
                    Erase atking_AI_貝琳達_水晶幻鏡紀錄狀態數
                    戰鬥系統類.執行動作_技能手動結束
                End If
            End If
   End Select
End If
End Sub
Sub 貝琳達_裂地冰牙()
Dim wtr As Integer, wert(1 To 3) As Boolean, wery As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "裂地冰牙" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(124, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "貝琳達" Then
   Select Case atkingckai(124, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(124, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
                   atkingckai(124, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) And atkingckai(124, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
                   atkingckai(124, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking3_2.jpg"
                   atkingno(i, 2) = 2
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
                            If liveus(角色待機人物紀錄數(1, wtr)) > 0 Then
                                戰鬥系統類.傷害執行_技能直傷_使用者 2, wtr
                                 Exit Do
                            End If
                        End If
                 Loop Until wery > 3
             End If
             atkingckai(124, 2) = 0
   End Select
End If
End Sub
Sub 貝琳達_溶魂之雨()
If FormMainMode.comaiatk(4).Caption = "溶魂之雨" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(125, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "貝琳達" Then
   Select Case atkingckai(125, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 1 _
                   And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(125, 2) = 0 Then
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 10
                        atkingckai(125, 2) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 1 _
                    Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(125, 2) = 1 Then
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 10
                        atkingckai(125, 2) = 0
                        atkingtrn(2) = Val(atkingtrn(2)) - 1
                        If atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = True Then
                              攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 15
                              atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = False
                        End If
                        If atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = True Then
                             攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 10
                             atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = False
                        End If
                End If
                 '=====================
                 If atkingckai(125, 2) = 1 Then
                     If pageqlead(2) >= 10 And atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = False Then
                         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 10
                         atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = True
                     ElseIf pageqlead(2) < 10 And atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = True Then
                         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 10
                         atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(1) = False
                     End If
                     If pageqlead(2) >= 15 And atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = False Then
                         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 15
                         atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = True
                     ElseIf pageqlead(2) < 15 And atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = True Then
                         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 15
                         atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數(2) = False
                     End If
                 End If
          End If
      Case 2
             atkingckai(125, 2) = 0
             Erase atking_AI_貝琳達_溶魂之雨_攻擊力加成紀錄數
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\貝琳達\atking4_2.jpg"
                   atkingno(i, 2) = 2
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
Sub 蕾_安魂曲_死神的鎮魂歌()
Dim rrr As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "安魂曲-死神的鎮魂歌" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(126, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(126, 1)
        Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          If rrr >= 1 And atkingckai(126, 2) = 0 Then
             atkingckai(126, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 1 And atkingckai(126, 2) = 1 Then
             atkingckai(126, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-安魂曲-死神的鎮魂歌_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(126, 2) = 0
             If livecom(角色人物對戰人數(2, 2)) <= 0 Then
                 For i = 2 To 3
                     If livecom(角色待機人物紀錄數(2, i)) > 0 Then
                        Do
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                                  If 人物異常狀態資料庫(2, j, 3) = 1 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 5
                                      FormMainMode.personcomspe(j).person_turn = 3
                                      人物異常狀態資料庫(2, j, 1) = 5
                                      人物異常狀態資料庫(2, j, 2) = 3
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                               If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, j, 1, app_path & "gif\異常狀態\atkup.gif", 5, 3
                                  異常狀態檢查數(1, 1) = 1
                                  異常狀態檢查數(1, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        Do
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                                  If 人物異常狀態資料庫(2, j, 3) = 2 And 人物異常狀態資料庫(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 5
                                      FormMainMode.personcomspe(j).person_turn = 3
                                      人物異常狀態資料庫(2, j, 1) = 5
                                      人物異常狀態資料庫(2, j, 2) = 3
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (角色待機人物紀錄數(2, i) - 1) + 1 To 14 * 角色待機人物紀錄數(2, i)
                               If 人物異常狀態資料庫(2, j, 2) = 0 Then
                                  戰鬥系統類.人物異常狀態表設定_初設 2, j, 2, app_path & "gif\異常狀態\defup.gif", 5, 3
                                  異常狀態檢查數(2, 1) = 1
                                  異常狀態檢查數(2, 2) = 1
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
Sub 蕾_EX_終曲_無盡輪迴的終結()
Dim num(1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.comaiatk(4).Caption = "Ex終曲-無盡輪迴的終結" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(127, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾" Then
   Select Case atkingckai(127, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(2, 4) >= 6 And atkingckai(127, 2) = 0 Then
               atkingckai(127, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 18
            ElseIf atkingpagetot(2, 4) < 6 And atkingckai(127, 2) = 1 Then
               atkingckai(127, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 18
            End If
          End If
        Case 2
             atking_AI_蕾_終曲_無盡輪迴的終結紀錄數 = 0
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\蕾\蕾-EX-終曲-無盡輪迴的終結_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8655
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 127
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=============
             atking_AI_蕾_終曲_無盡輪迴的終結紀錄數 = atkingpagetot(1, 2)
        Case 3
             atkingckai(127, 2) = 0
             If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                 num(1) = 1
                 num(2) = liveus(角色人物對戰人數(2, 2))
                 For i = 2 To 3
                    If liveus(角色待機人物紀錄數(1, i)) > 0 And liveus(角色待機人物紀錄數(1, i)) < num(2) Then
                        num(1) = i
                        num(2) = liveus(角色待機人物紀錄數(1, i))
                    End If
                Next
                戰鬥系統類.傷害執行_技能直傷_使用者 Val(擲骰表單溝通暫時變數(2)), num(1)
            End If
            '=================
            戰鬥系統類.傷害執行_技能直傷_使用者 Val(atking_AI_蕾_終曲_無盡輪迴的終結紀錄數), 1
            擲骰表單溝通暫時變數(2) = 0
            擲骰後骰傷害數 = 0
   End Select
End If
End Sub
Sub 羅莎琳_黑霧幻影()
If FormMainMode.comaiatk(1).Caption = "黑霧幻影" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(128, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羅莎琳" Then
   Select Case atkingckai(128, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 2 And atkingckai(128, 2) = 0 Then
               atkingckai(128, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
            ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 2) And atkingckai(128, 2) = 1 Then
               atkingckai(128, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_黑霧幻影_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 128
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_羅莎琳_黑霧幻影紀錄狀態數(i) = True
               End If
            Next
            目前數(18) = 1
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) <= 0 Then
                Do
                    If atking_AI_羅莎琳_黑霧幻影紀錄狀態數(目前數(18)) = True Then
                        目前數(16) = 目前數(18)
                        目前數(15) = 38
                        FormMainMode.tr牌組_回牌_電腦.Enabled = True
                        atking_AI_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
                        Exit Do
                    End If
                    目前數(18) = 目前數(18) + 1
                Loop Until 目前數(18) >= 106
            End If
            If 目前數(18) >= 106 Or Val(擲骰表單溝通暫時變數(2)) > 0 Then
                atkingckai(128, 1) = 6
                FormMainMode.骰子執行完啟動.Enabled = True
            End If
        Case 5
'            tr牌組_回牌_使用者.Enabled = True
'            atking_AI_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
        Case 6
            atkingckai(128, 2) = 0
            Erase atking_AI_羅莎琳_黑霧幻影紀錄狀態數
   End Select
End If
End Sub
Sub 羅莎琳_EX_黑霧幻影()
If FormMainMode.comaiatk(1).Caption = "Ex黑霧幻影" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(129, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羅莎琳" Then
   Select Case atkingckai(129, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 2) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(129, 2) = 0 Then
               atkingckai(129, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 9
            ElseIf (atkingpagetot(2, 2) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(129, 2) = 1 Then
               atkingckai(129, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 9
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\羅莎琳_Ex-黑霧幻影_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 129
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_羅莎琳_黑霧幻影紀錄狀態數(i) = True
               End If
            Next
            目前數(18) = 1
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) <= 0 Then
                Do
                    If atking_AI_羅莎琳_黑霧幻影紀錄狀態數(目前數(18)) = True Then
                        目前數(16) = 目前數(18)
                        目前數(15) = 38
                        FormMainMode.tr牌組_回牌_電腦.Enabled = True
                        atking_AI_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
                        Exit Do
                    End If
                    目前數(18) = 目前數(18) + 1
                Loop Until 目前數(18) >= 106
            End If
            If 目前數(18) >= 106 Or Val(擲骰表單溝通暫時變數(2)) > 0 Then
                atkingckai(129, 1) = 6
                FormMainMode.骰子執行完啟動.Enabled = True
            End If
        Case 5
'            tr牌組_回牌_使用者.Enabled = True
'            atking_AI_羅莎琳_黑霧幻影紀錄狀態數(目前數(16)) = False
        Case 6
            atkingckai(129, 2) = 0
            Erase atking_AI_羅莎琳_黑霧幻影紀錄狀態數
   End Select
End If
End Sub
Sub 洛洛妮_逆轉戰局的槍響()
Dim bloodnum As Integer
If FormMainMode.comaiatk(1).Caption = "逆轉戰局的槍響" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(130, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "洛洛妮" Then
   Select Case atkingckai(130, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(130, 2) = 0 Then
                   atkingckai(130, 2) = 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(130, 2) = 1 Then
                   atkingckai(130, 2) = 0
                 End If
          End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\洛洛妮\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -120
                   atkingno(i, 5) = 7035
                   atkingno(i, 6) = 9540
                   atkingno(i, 7) = 130
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數
                   atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) = livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2))
                   If atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) > 2 Then
                       atkingno(i, 11) = 1
                   Else
                       atkingno(i, 11) = 0
                   End If
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) And atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) = 0 Then
               戰鬥系統類.執行動作_洗牌
            End If
            atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) = atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) > atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2)
                    目前數(15) = 39
                    FormMainMode.tr牌組_抽牌_電腦.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(1) > atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_AI_洛洛妮_逆轉戰局的槍響_抽牌紀錄數(2)) <= 2 Then
                   atkingckai(130, 2) = 0
               Else
                   目前數(24) = 36
                   FormMainMode.等待時間_2.Enabled = True
               End If
            End If
        Case 5
            atkingckai(130, 2) = 0
            戰鬥系統類.執行動作_技能手動結束
   End Select
End If
End Sub
Sub 克頓_竊取資料()
If FormMainMode.comaiatk(1).Caption = "竊取資料" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(131, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "克頓" Then
   Select Case atkingckai(131, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(131, 2) = 0 Then
'            If pageqlead(1) >= 1 And atkingckai(131, 2) = 0 Then
               atkingckai(131, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(131, 2) = 1 Then
'            If pageqlead(1) < 1 And atkingckai(131, 2) = 1 Then
               atkingckai(131, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5655
                   atkingno(i, 6) = 9855
                   atkingno(i, 7) = 131
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
                階段狀態數 = 1
                目前數(21) = 1
                If pageqlead(1) > 0 Then
                    Do
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 6)) = 2 And Val(pagecardnum(m, 5)) = 1 Then
                            atking_AI_克頓_竊取資料_奪牌紀錄數(1) = m
                            turnpageoninatking = 1
                            atkingckai(131, 1) = 5
                            FormMainMode.card_Click (m)
                            Exit Do
                        End If
                    Loop
                    FormMainMode.trgoi1_Timer
                    FormMainMode.trgoi2_Timer
                 Else
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
                    atkingckai(131, 2) = 0
                    turnpageoninatking = 0
                    Erase atking_AI_克頓_竊取資料_奪牌紀錄數
                End If
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(131, 2) = 0
             turnpageoninatking = 0
             Erase atking_AI_克頓_竊取資料_奪牌紀錄數
        Case 5
             目前數(21) = 1
             atkingckai(131, 1) = 4
             atking_AI_克頓_竊取資料_奪牌紀錄數(2) = 目前數(5)
             '=========將座標指定至電腦手牌
             戰鬥系統類.座標計算_電腦手牌
             戰鬥系統類.執行動作_使用者牌_偷牌_電腦 atking_AI_克頓_竊取資料_奪牌紀錄數(1)
'             FormMainMode.card(atking_AI_克頓_竊取資料_奪牌紀錄數(1)).Width = 810
'             FormMainMode.card(atking_AI_克頓_竊取資料_奪牌紀錄數(1)).Height = 1260
'             FormMainMode.card(atking_AI_克頓_竊取資料_奪牌紀錄數(1)).Picture = LoadPicture(app_path & "card\" & pagecardnum(atking_AI_克頓_竊取資料_奪牌紀錄數(1), 8) & "-" & pageonin(atking_AI_克頓_竊取資料_奪牌紀錄數(1)) & ".bmp")
             目前數(5) = atking_AI_克頓_竊取資料_奪牌紀錄數(2)
             目前數(15) = 0
   End Select
End If
End Sub
Sub 克頓_逃亡計畫()
Dim rrr(1 To 2) As Integer '牌判斷暫時變數
Dim au As Integer
If FormMainMode.comaiatk(2).Caption = "逃亡計畫" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(132, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "克頓" Then
   Select Case atkingckai(132, 1)
      Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                    If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(2) = rrr(2) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1) And atkingckai(132, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingckai(132, 2) = 0 Then
                atkingckai(132, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1) And atkingckai(132, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingckai(132, 2) = 1 Then
                atkingckai(132, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
              End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7200
                   atkingno(i, 6) = 9990
                   atkingno(i, 7) = 132
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
                    If livecom(角色待機人物紀錄數(2, m)) > 0 Then
                        戰鬥系統類.傷害執行_技能直傷_電腦 3, m
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
                        戰鬥系統類.傷害執行_技能直傷_電腦 3, 1
                        Exit Do
                    End If
               Loop
               擲骰後骰傷害數 = 0
               擲骰表單溝通暫時變數(2) = 0
               atkingckai(132, 2) = 0
   End Select
End If
End Sub
Sub 克頓_隱蔽射擊()
Dim p, i, j As Integer
If FormMainMode.comaiatk(3).Caption = "隱蔽射擊" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(133, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "克頓" Then
   Select Case atkingckai(133, 1)
      Case 1
         If movecp > 1 Then
            If atkingpagetot(2, 5) >= 2 And atkingpagetot(2, 3) >= 1 And atkingckai(133, 2) = 0 Then
               atkingckai(133, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 4
            End If
            If (atkingpagetot(2, 5) < 2 Or atkingpagetot(2, 3) < 1) And atkingckai(133, 2) = 1 Then
               atkingckai(133, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 4
             End If
         End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking3_2.jpg"
                   atkingno(i, 2) = 2
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
             atkingckai(133, 1) = 3
        Case 3
             If livecom(角色人物對戰人數(2, 2)) = livecommax(角色人物對戰人數(2, 2)) Then
                    atking_AI_克頓_隱蔽射擊骰量紀錄數(1) = 擲骰後骰傷害數
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
                       atkingckai(133, 1) = 4
                       FormMainMode.骰子執行完啟動.Enabled = False
                       目前數(22) = 12
                       FormMainMode.等待時間.Enabled = True
                Else
                       atkingckai(133, 2) = 0
                       FormMainMode.骰子執行完啟動.Enabled = True
                       Erase atking_AI_克頓_隱蔽射擊骰量紀錄數
                End If
          Case 4
                atking_AI_克頓_隱蔽射擊骰量紀錄數(2) = 擲骰後骰傷害數
                '==========================
                擲骰表單溝通暫時變數(2) = atking_AI_克頓_隱蔽射擊骰量紀錄數(1) + atking_AI_克頓_隱蔽射擊骰量紀錄數(2)
                擲骰後骰傷害數 = Val(擲骰表單溝通暫時變數(2))
                atkingckai(133, 2) = 0
                Erase atking_AI_克頓_隱蔽射擊骰量紀錄數
   End Select
End If
End Sub
Sub 克頓_惡意情報()
Dim rrr(1 To 2) As Integer '牌判斷暫時變數
Dim au As Integer
If FormMainMode.comaiatk(4).Caption = "惡意情報" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(134, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "克頓" Then
   Select Case atkingckai(134, 1)
      Case 1
            If movecp > 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 3 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 3 Then
                           rrr(2) = rrr(2) + 1
                        End If
                    End If
                 Next
                 '========================
                 If (rrr(1) >= 1 And rrr(2) >= 1) And atkingpagetot(2, 4) >= 2 And atkingckai(134, 2) = 0 Then
'                 If pageqlead(2) >= 1 And atkingckai(134, 2) = 0 Then
                    atkingckai(134, 2) = 1
                 ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or atkingpagetot(2, 4) < 2) And atkingckai(134, 2) = 1 Then
'                 ElseIf pageqlead(2) < 1 And atkingckai(134, 2) = 1 Then
                    atkingckai(134, 2) = 0
                  End If
            End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\克頓\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7050
                   atkingno(i, 6) = 10005
                   atkingno(i, 7) = 134
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             '=====================
              For i = 1 To 106
                   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                      atking_AI_克頓_惡意情報紀錄數(i) = 1
                      atking_AI_克頓_惡意情報紀錄數(0) = Val(atking_AI_克頓_惡意情報紀錄數(0)) + 1
                   End If
               Next
        Case 4
               For i = 1 To 106
                   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                      turnpageoninatking = 1
                      階段狀態數 = 1
                      FormMainMode.card_Click (i)
                      目前數(15) = 40
                      FormMainMode.對齊完成檢查.Enabled = False
                      Exit Sub
                    End If
               Next
               If i = 107 And atking_AI_克頓_惡意情報紀錄數(0) > 0 Then
                    For k = 1 To 106
                         If atking_AI_克頓_惡意情報紀錄數(k) = 1 Then
                             atking_AI_克頓_惡意情報紀錄數(k) = 0
                             turnpageoninatking = 1
                             階段狀態數 = 1
                             FormMainMode.card_Click (k)
                             目前數(21) = 11
                             FormMainMode.對齊完成檢查.Enabled = False
                             Exit Sub
                         End If
                    Next
                End If
         Case 5
               turnpageonin = 0
               For k = 1 To 106
                     If atking_AI_克頓_惡意情報紀錄數(k) = 1 Then
                         atking_AI_克頓_惡意情報紀錄數(k) = 0
                         turnpageoninatking = 1
                         階段狀態數 = 1
                         FormMainMode.card_Click (k)
                         目前數(21) = 11
                         FormMainMode.對齊完成檢查.Enabled = False
                         Exit Sub
                     End If
                Next
                If k = 107 Then
                    atkingckai(134, 2) = 0
                    turnpageoninatking = 0
                    turnpageonin = 0
                    階段狀態數 = 4
                    Erase atking_AI_克頓_惡意情報紀錄數
                    戰鬥系統類.執行動作_技能手動結束
                End If
   End Select
End If
End Sub
Sub 艾茵_一顆心()
Dim cardnum(1 To 2) As Integer '暫時變數
If FormMainMode.comaiatk(1).Caption = "一顆心" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(135, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾茵" Then
   Select Case atkingckai(135, 1)
        Case 1
           If movecp = 2 Then
                 If atkingpagetot(2, 4) >= 3 And atkingckai(135, 2) = 0 Then
'                 If pageqlead(1) >= 1 And atkingckai(135, 2) = 0 Then
                   atkingckai(135, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf atkingpagetot(2, 4) < 3 And atkingckai(135, 2) = 1 Then
'                ElseIf pageqlead(1) < 1 And atkingckai(135, 2) = 1 Then
                   atkingckai(135, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
           End If
        Case 2
              For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\艾茵\艾茵_一顆心_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8925
                   atkingno(i, 6) = 9105
                   atkingno(i, 7) = 135
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
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
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
            目前數(20) = cardnum(2)
            FormMainMode.tr使用者牌_偷牌.Enabled = True
            目前數(21) = 1
            atkingckai(135, 2) = 0
   End Select
End If
End Sub
Sub 尤莉卡_奸佞的鐵鎚()
Dim wert As Integer '暫時變數
If FormMainMode.comaiatk(1).Caption = "奸佞的鐵鎚" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(136, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "尤莉卡" Then
   Select Case atkingckai(136, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 4) >= 1 And atkingckai(136, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
                   atkingckai(136, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   '==========
                   If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 5
                   End If
                   '==========
                End If
                If (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 4) < 1) And atkingckai(136, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
                   atkingckai(136, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   '==========
                   If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                       攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 5
                   End If
                   '==========
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9420
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 136
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then wert = 2 Else wert = 1
             '====================
             Do
                  For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
                      If 人物異常狀態資料庫(1, i, 3) > 0 Then
                             人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
                             If 人物異常狀態資料庫(1, i, 2) = 0 Then
                               '===繼承下一狀態資料
                                戰鬥系統類.異常狀態繼承_使用者
                                If 人物異常狀態資料庫(1, i, 3) = 15 Then
                                    戰鬥系統類.傷害執行_立即死亡_使用者 1  '自壞回合數歸0時執行死亡動作
                                End If
                             Else
                                FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
                             End If
                     End If
                  Next
                  '=====================
                  wert = Val(wert) - 1
             Loop Until wert <= 0
        Case 4
            If Val(擲骰表單溝通暫時變數(2)) > 0 And 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                Randomize
                wert = Int(Rnd() * 3) + 1
                Do
                   For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                     If 人物異常狀態資料庫(2, i, 2) >= 3 And 人物異常狀態資料庫(2, i, 3) = 40 Then
                        Exit Do
                     End If
                     If 人物異常狀態資料庫(2, i, 3) = 40 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 3 Then
                         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + (Val(wert) - 1)
                         If 人物異常狀態資料庫(2, i, 2) > 3 Then 人物異常狀態資料庫(2, i, 2) = 3
                         FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 2) = 0 And (Val(wert) - 1) > 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 2, i, 40, app_path & "gif\異常狀態\臨界.gif", 0, (Val(wert) - 1)
                         異常狀態檢查數(40, 1) = 1
                         異常狀態檢查數(40, 2) = 1
                         Exit Do
                     End If
                   Next
                   If i = 14 * 角色人物對戰人數(2, 2) + 1 And (Val(wert) - 1) = 0 Then Exit Do
                Loop
            End If
            atkingckai(136, 2) = 0
            '===============超載技能使用結束
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                atkingckai(139, 1) = 6
                AI技能.尤莉卡_超載 '(階段6)
            End If
            '===============
   End Select
End If
End Sub
Sub 尤莉卡_不善的信仰()
Dim wert As Integer '暫時變數
If FormMainMode.comaiatk(2).Caption = "不善的信仰" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(137, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "尤莉卡" Then
   Select Case atkingckai(137, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 2) >= 3 And atkingckai(137, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 3
                   atkingckai(137, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 2) < 3 And atkingckai(137, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 3
                   atkingckai(137, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7170
                   atkingno(i, 6) = 10440
                   atkingno(i, 7) = 137
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then wert = 2 Else wert = 3
                '=============
                If Val(擲骰表單溝通暫時變數(2)) Mod Val(wert) = 0 Then
                    擲骰表單溝通暫時變數(2) = 0
                    擲骰後骰傷害數 = 擲骰表單溝通暫時變數(2)
                End If
            End If
            '======================================
            If Val(擲骰表單溝通暫時變數(2)) <= 0 And 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                Randomize
                wert = Int(Rnd() * 3) + 1
                Do
                   For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                     If 人物異常狀態資料庫(2, i, 2) >= 3 And 人物異常狀態資料庫(2, i, 3) = 40 Then
                        Exit Do
                     End If
                     If 人物異常狀態資料庫(2, i, 3) = 40 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 3 Then
                         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + (Val(wert) - 1)
                         If 人物異常狀態資料庫(2, i, 2) > 3 Then 人物異常狀態資料庫(2, i, 2) = 3
                         FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
                         Exit Do
                     End If
                   Next
                   For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                      If 人物異常狀態資料庫(2, i, 2) = 0 And (Val(wert) - 1) > 0 Then
                         戰鬥系統類.人物異常狀態表設定_初設 2, i, 40, app_path & "gif\異常狀態\臨界.gif", 0, (Val(wert) - 1)
                         異常狀態檢查數(40, 1) = 1
                         異常狀態檢查數(40, 2) = 1
                         Exit Do
                     End If
                   Next
                   If i = 14 * 角色人物對戰人數(2, 2) + 1 And (Val(wert) - 1) = 0 Then Exit Do
                Loop
            End If
            atkingckai(137, 2) = 0
            '===============超載技能使用結束
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                atkingckai(139, 1) = 6
                AI技能.尤莉卡_超載 '(階段6)
            End If
            '===============
   End Select
End If
End Sub
Sub 尤莉卡_曲惡的安寧()
Dim wert As Integer '暫時變數
If FormMainMode.comaiatk(3).Caption = "曲惡的安寧" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(138, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "尤莉卡" Then
   Select Case atkingckai(138, 1)
      Case 1
           If movecp = 3 Then
                If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 1 And atkingckai(138, 2) = 0 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 6
                   atkingckai(138, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 1) And atkingckai(138, 2) = 1 Then
                   攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 6
                   atkingckai(138, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6450
                   atkingno(i, 6) = 10215
                   atkingno(i, 7) = 138
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                For k = 1 To 3
                    戰鬥系統類.回復執行_電腦 2, k
                Next
            Else
                戰鬥系統類.回復執行_電腦 2, 1
            End If
            '======================================
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                    If 人物異常狀態資料庫(2, i, 3) = 40 Then
                      人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
                      If 人物異常狀態資料庫(2, i, 2) = 0 Then
                        '===繼承下一狀態資料
                         戰鬥系統類.異常狀態繼承_使用者
                         異常狀態檢查數(40, 2) = 0
                     Else
                         FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
                         異常狀態檢查數(40, 1) = 1
                     End If
                   End If
                Next
            End If
            atkingckai(138, 2) = 0
            '===============超載技能使用結束
            If 戰鬥系統類.特殊_尤莉卡_檢查超載是否啟動_電腦 = True Then
                atkingckai(139, 1) = 6
                AI技能.尤莉卡_超載 '(階段6)
            End If
            '===============
   End Select
End If
End Sub
Sub 尤莉卡_超載()
Dim wert As Integer '暫時變數
If FormMainMode.comaiatk(4).Caption = "超載" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(139, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "尤莉卡" Then
   Select Case atkingckai(139, 1)
      Case 1
                If atkingpagetot(2, 4) >= 1 And atkingckai(139, 2) = 0 Then
                   atkingckai(139, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 1 And atkingckai(139, 2) = 1 Then
                   atkingckai(139, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
      Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\尤莉卡\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7920
                   atkingno(i, 6) = 10005
                   atkingno(i, 7) = 139
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             Do
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                 If 人物異常狀態資料庫(2, i, 2) >= 3 And 人物異常狀態資料庫(2, i, 3) = 40 Then
                    atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2
                    Exit Do
                 End If
                 If 人物異常狀態資料庫(2, i, 3) = 40 And 人物異常狀態資料庫(2, i, 2) > 0 And 人物異常狀態資料庫(2, i, 2) < 3 Then
                     人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) + 1
                     FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
                     atking_AI_尤莉卡_超載目前階段紀錄數(3) = 1
                     Exit Do
                 End If
               Next
               For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
                  If 人物異常狀態資料庫(2, i, 2) = 0 Then
                     戰鬥系統類.人物異常狀態表設定_初設 2, i, 40, app_path & "gif\異常狀態\臨界.gif", 0, 1
                     異常狀態檢查數(40, 1) = 1
                     異常狀態檢查數(40, 2) = 1
                     atking_AI_尤莉卡_超載目前階段紀錄數(3) = 1
                     Exit Do
                 End If
               Next
            Loop
            '========================超載3時執行封印
            If atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2 Then
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
        Case 4
            '========================超載3時攻防2倍階段-執行
            If atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2 Then
                If Val(atking_AI_尤莉卡_超載目前階段紀錄數(4)) = 0 Then
                    atking_AI_尤莉卡_超載目前階段紀錄數(1) = 攻擊防禦骰子總數(2)
                    atking_AI_尤莉卡_超載目前階段紀錄數(2) = 攻擊防禦骰子總數(2) * 2
                    atking_AI_尤莉卡_超載目前階段紀錄數(4) = 1
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) * 2
                ElseIf Val(atking_AI_尤莉卡_超載目前階段紀錄數(4)) = 1 Then
                    atking_AI_尤莉卡_超載目前階段紀錄數(1) = atking_AI_尤莉卡_超載目前階段紀錄數(1) + (攻擊防禦骰子總數(2) - atking_AI_尤莉卡_超載目前階段紀錄數(2))
                    攻擊防禦骰子總數(2) = atking_AI_尤莉卡_超載目前階段紀錄數(1) * 2
                    atking_AI_尤莉卡_超載目前階段紀錄數(2) = atking_AI_尤莉卡_超載目前階段紀錄數(1) * 2
                End If
            End If
        Case 5
            '========================超載3時攻防2倍階段-開始階段時清除資料
            atking_AI_尤莉卡_超載目前階段紀錄數(1) = 0
            atking_AI_尤莉卡_超載目前階段紀錄數(2) = 0
            atking_AI_尤莉卡_超載目前階段紀錄數(4) = 0
        Case 6
            '========================超載技能結束(普通)
            atkingckai(139, 2) = 0
            Erase atking_AI_尤莉卡_超載目前階段紀錄數
        Case 7
            '========================更換角色時重新載入技能
            If atking_AI_尤莉卡_超載目前階段紀錄數(3) > 0 Then
                atking_AI_尤莉卡_超載目前階段紀錄數(1) = 0
                atking_AI_尤莉卡_超載目前階段紀錄數(2) = 0
                atking_AI_尤莉卡_超載目前階段紀錄數(4) = 0
            End If
        Case 8
            '========================超載技能結束(回合結束階段)
            atkingckai(139, 2) = 0
            If atking_AI_尤莉卡_超載目前階段紀錄數(3) = 2 Then
                戰鬥系統類.執行動作_清除所有異常狀態_電腦
            End If
            Erase atking_AI_尤莉卡_超載目前階段紀錄數
   End Select
End If
End Sub
Sub 羅莎琳_EX_染血之刃()
If FormMainMode.comaiatk(2).Caption = "Ex染血之刃" And (執行動作_檢查是否有指定異常狀態(2, 23) = False Or atkingckai(140, 2) = 1) _
   And FormMainMode.compi1(角色人物對戰人數(2, 2)) = "羅莎琳" Then
   Select Case atkingckai(140, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 3) >= 2 And atkingckai(140, 2) = 0 Then
               atkingckai(140, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 9
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 3) < 2) And atkingckai(140, 2) = 1 Then
               atkingckai(140, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 9
            End If
          End If
        Case 2
             For i = 人物技能數字指示 To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\羅莎琳\atkingEX2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 140
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            回復執行_電腦 1, 1
        Case 4
            atkingckai(140, 2) = 0
            If Val(擲骰表單溝通暫時變數(2)) > 0 Then
                回復執行_電腦 1, 1
            End If
   End Select
End If
End Sub

