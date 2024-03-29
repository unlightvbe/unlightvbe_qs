Attribute VB_Name = "戰鬥系統類"
Option Explicit
Public Const a1a As String = "ATK-劍"
Public Const a2a As String = "DEF"
Public Const a3a As String = "MOV"
Public Const a4a As String = "SPE"
Public Const a5a As String = "ATK-槍"
Public Const a6a As String = "DRAW"
Public Const a7a As String = "BRK"
Public Const a8a As String = "HPL"
Public Const a9a As String = "HPW"
Public Const b1b As Integer = 1
Public Const b2b As Integer = 2
Public Const b3b As Integer = 3
Public Const b4b As Integer = 4
Public Const b5b As Integer = 5
Public Const b6b As Integer = 6
Public Const b7b As Integer = 7
Public Const b8b As Integer = 8
Public Const b9b As Integer = 9

Public goicheck(1 To 2) As Integer   '攻擊/防禦模式加骰數值檢查碼
Public liveus(1 To 3) As Integer, livecom(1 To 3) As Integer, liveusmax(1 To 3) As Integer, livecommax(1 To 3) As Integer
Public BattleTurn As Integer, BattleCardNum As Integer, atkus(1 To 3) As Integer, atkcom(1 To 3) As Integer, defus(1 To 3) As Integer, defcom(1 To 3) As Integer, pagecheckus As Integer, pagecheckcom As Integer, pagegive As Integer, goidefus As Integer, movecom As Integer, moveus As Integer, movecp As Integer, chkcomck As Integer, uslevel(1 To 3) As Integer, comlevel(1 To 3) As Integer, liveus41(1 To 3) As Integer, livecom41(1 To 3) As Integer, movecheckcom As Integer, movecheckus As Integer
Attribute BattleCardNum.VB_VarUserMemId = 1073741829
Public nameus(1 To 3) As String, namecom(1 To 3) As String
Attribute nameus.VB_VarUserMemId = 1073741849
Public moveturn As Integer  '攻擊／防禦模式先後檢查碼(1.使用者先攻/2.電腦先攻)
Attribute moveturn.VB_VarUserMemId = 1073741851
Public atkinghelpxy(1 To 2, 1 To 4, 1 To 2) As Integer    '技能說明欄座標指定資料(1.使用者方/2.電腦方,第1~4個技能,1.Left/2.Top(座標))
Attribute atkinghelpxy.VB_VarUserMemId = 1073741852
Public pageusleadmax(0 To 1) As Integer   '使用者牌順序計數表(0.手牌/1.出牌)
Attribute pageusleadmax.VB_VarUserMemId = 1073741853
Public pagecomleadmax(0 To 1) As Integer   '電腦牌順序計數表(0.手牌/1.出牌)
Attribute pagecomleadmax.VB_VarUserMemId = 1073741854
Public pageqlead(1 To 2) As Integer   '出牌計數變數(1.使用者/2.電腦)
Attribute pageqlead.VB_VarUserMemId = 1073741855
Public pageglead(1 To 2) As Integer   '手牌計數變數(1.使用者/2.電腦)
Attribute pageglead.VB_VarUserMemId = 1073741856
Public movedsus As Integer   '使用者移動階段決定值變數
Attribute movedsus.VB_VarUserMemId = 1073741857
Public turnpageonin As Integer  '階段是否可出牌變數(一般)
Attribute turnpageonin.VB_VarUserMemId = 1073741858
Public turnpageoninatking As Integer  '階段是否可出牌變數(技能使用)
Attribute turnpageoninatking.VB_VarUserMemId = 1073741859
Public goickus As Integer    '牌值一次檢查碼
Attribute goickus.VB_VarUserMemId = 1073741860
Public atkingck(1 To 2, 1 To 3, 1 To 8, 1 To 3) As Integer    '技能階段啟動碼(1.使用者/2.電腦,1~3.人物編號/1~4人物自身技能項目;5~8人物自身被動技項目,1.技能啟動標記/2.這回合間啟動次數(主動技->動畫執行)/3.這場戰鬥間啟動次數(主動技->動畫執行))
Attribute atkingck.VB_VarUserMemId = 1073741861
Public atkingckdice(1 To 2, 1 To 2, 1 To 4) As String    '人物技能骰子影響紀錄暫時變數(1.使用者/2.電腦,1.對使用者/2.對電腦,1.主動技/2.被動技/3.異常狀態/4.人物實際狀態,對總骰數之影響量變化串)
Attribute atkingckdice.VB_VarUserMemId = 1073741862
Public atkingtrn(1 To 4) As Integer    '技能計數器暫時儲存變數(1.使用者(現)/2.電腦(現)/3.使用者(備份)/4.電腦(備份))
Attribute atkingtrn.VB_VarUserMemId = 1073741863
Public akhpnm As Integer  '技能說明暫時變數
Attribute akhpnm.VB_VarUserMemId = 1073741864
Public turnatk As Integer  '攻擊／防禦階段變數(1.使用者攻擊、電腦防禦,2.使用者防禦、電腦攻擊,3.發牌、移動)
Attribute turnatk.VB_VarUserMemId = 1073741865
Public trend暫時變數 As Integer    '結束階段計數器暫時變數
Attribute trend暫時變數.VB_VarUserMemId = 1073741866
Public HP檢查變數 As Boolean    'HP檢查階段是否已檢查變數
Attribute HP檢查變數.VB_VarUserMemId = 1073741867
Public HP檢查階段數 As Integer    'HP檢查階段變數(1.移動階段後,2.攻擊/防禦階段前,3.攻/防禦階段後)
Attribute HP檢查階段數.VB_VarUserMemId = 1073741868
Public 距離單位(1 To 2, 1 To 2, 1 To 2) As Integer  '距離單位暫時儲存資料(1.HP血條/2.牌移動,1.使用者/2.電腦,1.Left單位/2.Top單位)
Attribute 距離單位.VB_VarUserMemId = 1073741869
Public personminixy(1 To 2, 1 To 3, 1 To 3, 1 To 2) As Integer    '小人物圖片座標指定資料(1.使用者/2.電腦,第n位,1.近距離/2.中距離/3.遠距離,1.Left/2.Top(座標))
Attribute personminixy.VB_VarUserMemId = 1073741870
Public 異常狀態檢查數(1 To 40, 1 To 2) As Integer    '異常狀態啟動碼(x.異常狀態編號,1.狀態執行階段/2.狀態啟動檢查值)
Attribute 異常狀態檢查數.VB_VarUserMemId = 1073741871
Public 技能動畫顯示階段數 As Integer    '技能動畫計數器階段碼(1.攻擊/防禦階段-普通,2.移動階段-普通/3.發牌階段後、移動階段前/4.移動階段後/5.攻擊階段後/6.防禦階段後/7.回合結束時)
Attribute 技能動畫顯示階段數.VB_VarUserMemId = 1073741872
Public 攻擊防禦骰子總數(1 To 4) As Integer    '攻擊/防禦模式骰子數量資料(1.使用者(總)/2.電腦(總)/3.使用者(原)/4.電腦(原))
Attribute 攻擊防禦骰子總數.VB_VarUserMemId = 1073741873
Public atkingpagetot(1 To 2, 1 To 5) As Integer  '每階段出牌種類及數值統計資料(1.使用者/2.電腦,1.劍/2.防/3.移/4.特/5.槍)
Attribute atkingpagetot.VB_VarUserMemId = 1073741874
Public 骰數零檢查值(1 To 2) As Boolean    '當前階段骰子數量是否為零檢查數(1.使用者/2.電腦)
Attribute 骰數零檢查值.VB_VarUserMemId = 1073741875
Public 牌總階段數(1 To 3) As Integer    '牌擁有總階段數(1.使用者/2.電腦/3.總計)
Attribute 牌總階段數.VB_VarUserMemId = 1073741876
Public 牌移動暫時變數(1 To 3) As Long    '牌移動計數器暫時變數(1.Left單位/2.Top單位/3.牌張編號)
Attribute 牌移動暫時變數.VB_VarUserMemId = 1073741877
Public 目前數(1 To 33) As Integer    '總暫時變數
Attribute 目前數.VB_VarUserMemId = 1073741878
'2.(1)變成使用者發牌階段-(2)變成電腦發牌階段-(3)變成發牌檢查階段
'3.使用者出牌對齊距離統計,
'4.使用者手牌靠左對齊距離統計,
'5.使用者牌對齊篩選數,
'6.電腦出牌/亮牌計數暫時數,
'7.電腦出牌對齊距離統計,
'8.電腦手牌對齊距離統計,
'9.電腦牌對齊篩選數,
'10.收牌階段數,
'11.收牌統計暫時值,
'12.收牌牌減值暫時數,
'13.使用者手牌對齊2列第1張牌編號移動暫時數,
'14.等待時間計數器(1,2)暫時數,
'15牌移動計數器執行階段數-(1)發牌階段-(2)偷牌階段-電腦手牌對齊,
'16.牌翻牌/回牌/棄牌牌編號暫時數,
'17.電腦手牌對齊階段數-(1)-電腦出牌階段-(2)偷牌後對齊階段
'20.使用者翻牌牌編號暫時數
'21.使用者手牌對齊階段數
'22.等待時間計數器執行階段數
'25.電腦移動階段出牌計數暫時變數
'26.骰子執行後技能啟動階段HP檢查是否完成數
'29.技能啟動動畫計數暫時數
'30.牌堆洗牌時既有卡牌數量暫時紀錄數
'31.移動階段初始Timer-是否第一次啟動暫時紀錄數
'32.使用者出牌-AI出牌控制牌目前數
'33.使用者出牌-AI出牌移動階段選擇行動數
Public 距離單位_收牌暫時數() As Integer  '收牌個別距離單位暫時儲存變數(第x順序,1.Left單位/2.Top單位/3.牌張編號)
Attribute 距離單位_收牌暫時數.VB_VarUserMemId = 1073741879
Public 階段狀態數 As Integer    '每階段開始結束狀態檢查數(1.開始階段(使用者)/2.結束階段(使用者)/3.開始階段(電腦)/4.結束階段(電腦)/5.交換角色)
Attribute 階段狀態數.VB_VarUserMemId = 1073741880
Public 小人物頭像移動方向數(1 To 2) As Integer    '小人物頭像移動方向狀態數(1.使用者/2.電腦[1.向內,2.向外])
Attribute 小人物頭像移動方向數.VB_VarUserMemId = 1073741881
Public 血量計數器動畫暫時變數(1 To 2, 1 To 2) As Integer    '開始初始階段-血量動畫計數器暫時變數(1.使用者血條/2.電腦血條,1.每次移動量/2.是否已完成)
Attribute 血量計數器動畫暫時變數.VB_VarUserMemId = 1073741882
Public 時間軸顏色變化紀錄暫時變數(1 To 4, 1 To 3) As Integer    '時間軸進行顏色變化階段紀錄暫時變數(1~3(1)單位變化量(1(1).時間軸(內外))/2.目前累計量/3.目前顏色(R,G,B),4.(1)時間軸(外)階段數-(1)黑變紅-(2)紅變黑/2.目前累計量/3.目前顏色(R))
Attribute 時間軸顏色變化紀錄暫時變數.VB_VarUserMemId = 1073741883
Public 開始卡片移動動畫完成數(1 To 2, 1 To 4) As Integer   '開始時每張卡片移動動畫完成紀錄數(1.使用者/2.電腦,1~3.卡片/4.目前第幾張)
Attribute 開始卡片移動動畫完成數.VB_VarUserMemId = 1073741884
Public 交換角色紀錄暫時變數(1 To 4) As Integer    '交換角色雙方紀錄暫時數(1.使用者/2.電腦/3.是否當下首次/4.交換角色完執行階段數)
Attribute 交換角色紀錄暫時變數.VB_VarUserMemId = 1073741885
Public 戰鬥模式勝敗紀錄數 As Integer    '戰鬥系統當前勝敗紀錄暫時變數(1.使用者方勝利/2.使用者方敗北/3.平手)
Attribute 戰鬥模式勝敗紀錄數.VB_VarUserMemId = 1073741886
Public 電腦方移動階段選擇數 As Integer    '移動階段電腦方選擇之行動暫時變數
Attribute 電腦方移動階段選擇數.VB_VarUserMemId = 1073741887
Public 電腦方事件卡是否出完選擇數 As Boolean    '電腦方先出事件卡是否出完暫時紀錄
Attribute 電腦方事件卡是否出完選擇數.VB_VarUserMemId = 1073741888
Public 人物卡面背面編號紀錄數(1 To 7) As Integer    '人物卡片背面技能說明人物編號暫時變數(1.(1).使用者/(2).電腦,2.第n位,3.目前使用者方使用人物編號/4.目前選擇之技能編號(使用者方使用人物)/5.目前選擇之技能編號(其他)/6~7.目前選擇之技能編號(交換角色)
Attribute 人物卡面背面編號紀錄數.VB_VarUserMemId = 1073741889
Public 擲骰表單溝通暫時變數(1 To 10) As Integer    '擲骰介面溝通暫時變數(1.一回合中先後判斷(1.前/2.後),2.擲骰後有效傷害數,3.擲骰後傷害對象(1.使用者/2.電腦),4.(1.使用者先攻/2.電腦先攻)/5.當前骰值(使用者)/6.當前骰值(電腦)/7.系統公用骰值(使用者)/8.系統公用骰值(電腦)/9.擲骰前骰值-總骰(使用者)/10.擲骰前骰值-總骰(電腦))
Attribute 擲骰表單溝通暫時變數.VB_VarUserMemId = 1073741890
Public 人物消失檢查暫時變數(1 To 3) As Integer    '人物消失檢查計數器紀錄暫時變數(1.目前計數/2.使用者標記/3.電腦標記)
Attribute 人物消失檢查暫時變數.VB_VarUserMemId = 1073741891
Public 公用牌各牌類型紀錄數(0 To 29, 1 To 2) As Integer    '各場景公用牌牌類型紀錄暫時變數(0.(1)目前已發牌總數量/(2)目前場景牌總數量,1~29.(1)目前已使用之牌數/(2)該牌型能使用之總數量)
Attribute 公用牌各牌類型紀錄數.VB_VarUserMemId = 1073741892
Public 卡片人物資訊檔案讀取失敗紀錄串 As String    '卡片人物資訊檔案讀取失敗時檔案名紀錄暫時變數
Attribute 卡片人物資訊檔案讀取失敗紀錄串.VB_VarUserMemId = 1073741893
Public 顯示列雙方數值鎖定紀錄數(1 To 2) As Boolean    '戰鬥系統顯示列雙方數值鎖定表示紀錄變數(1.使用者方/2.電腦方)
Attribute 顯示列雙方數值鎖定紀錄數.VB_VarUserMemId = 1073741894
Public 是否系統公骰 As Boolean    '是否為系統公骰紀錄數
Attribute 是否系統公骰.VB_VarUserMemId = 1073741895
Public 戰鬥擲骰介面人物立繪圖路徑紀錄數(1 To 2) As String    '戰鬥系統擲骰介面雙方人物立繪圖路徑紀錄數(1.使用者方/2.電腦方)
Attribute 戰鬥擲骰介面人物立繪圖路徑紀錄數.VB_VarUserMemId = 1073741896
Public 人物實際狀態資料庫(1 To 2, 1 To 3, 1 To 9) As String
Attribute 人物實際狀態資料庫.VB_VarUserMemId = 1073741897
'人物實際狀態資料(1.使用者/2.電腦,1~3.第n位,1.技能統一識別碼/2.人物形象是否更換(1.是/0.否)/
'3.欲套用之大人物形象圖片路徑/4.欲套用之小人物形象圖片路徑/5.欲套用之小人物形象影子圖片路徑/
'6.欲套用之顯示列人物頭像圖片路徑/7.小人物形象影子座標Left/8.小人物形象影子座標Top/9.人物實際狀態名稱)
Public 系統顯示界面紀錄數 As Integer    '戰鬥系統顯示介面設定紀錄數(1.舊版/2.新版)
Attribute 系統顯示界面紀錄數.VB_VarUserMemId = 1073741898
Public 等待時間佇列(1 To 2) As New Collection    '戰鬥系統等待時間計數器工作佇列
Attribute 等待時間佇列.VB_VarUserMemId = 1073741899
Public 人物異常狀態列表(1 To 2, 1 To 3) As Collection    '異常狀態列表(1.使用者/2.電腦,第n位)
Attribute 人物異常狀態列表.VB_VarUserMemId = 1073741900
Public ActiveSkillObj(1 To 2, 1 To 4) As clsPersonActiveSkill    '戰鬥系統主動技能說明物件(1.使用者方/2.電腦方,第n個)
Attribute ActiveSkillObj.VB_VarUserMemId = 1073741901
Public PersonCardShowOnMode(1 To 2, 1 To 3) As Boolean    '戰鬥系統人物卡片資訊是否展示(1.使用者方/2.電腦方,第n個)
Attribute PersonCardShowOnMode.VB_VarUserMemId = 1073741902
Public CardDeckCollection(0 To 9) As Collection    '戰鬥系統卡牌牌堆集合(0.卡牌索引/1.牌堆/2.墓地牌/3.事件卡牌堆(使用者方)/4.事件卡牌堆(電腦方)/5.手牌(使用者方)/6.出牌(使用者方)/7.手牌(電腦方)/8.出牌(電腦方)/9.棄牌)
Attribute CardDeckCollection.VB_VarUserMemId = 1073741903
Public ActionCardTotNum As Integer    '戰鬥系統卡牌總發行紀錄數
Attribute ActionCardTotNum.VB_VarUserMemId = 1073741904
Sub 人物技能欄燈開關(ByVal isOn As Boolean, ByVal num As Integer)
    FormMainMode.PEAFInterface.ActiveSkillLight 1, num, isOn
End Sub
Function 執行動作_路徑使用新式異常狀態圖案(ByVal ph As String) As String
    Dim i As Integer
    For i = 1 To Len(ph)
        If Mid(ph, i, 1) = "." Then
            ph = Mid(ph, 1, i - 1) & "new" & Right(ph, 4)
            Exit For
        End If
    Next
    執行動作_路徑使用新式異常狀態圖案 = ph
End Function
Sub 傷害執行_技能直傷_使用者(ByVal tot As Integer, ByVal num As Integer, ByVal isEvent As Boolean)
    If tot <= 0 Then Exit Sub
    If isEvent = True Then
        Dim stageInfoListObj As clsVSStageObj
        Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
        '===============================
        VBEStageNum(0) = 46
        VBEStageNum(1) = -1    '受到傷害方(1.使用者/2.電腦)
        VBEStageNum(2) = num    '受到傷害人物編號
        VBEStageNum(3) = 2    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
        VBEStageNum(4) = tot    '受到傷害之數值
        stageInfoListObj.Argument = tot  '受到傷害之數值
        stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害方(1.使用者/2.電腦)
        stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num)    '受到傷害人物編號
        stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2"    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
        '===========================執行階段插入點(46)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 46, 1
        '============================
        If stageInfoListObj.CommandStr = "PersonBloodControl" Then
            If stageInfoListObj.Value = "BLOODOFF" Then
                Exit Sub
            Else
                Dim tmpstr() As String
                tmpstr = Split(stageInfoListObj.Value, "%")
                If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                    tot = Val(tmpstr(1))
                End If
            End If
        End If
    End If
    Select Case num
        Case 1
            If tot > 0 And liveus(角色人物對戰人數(1, 2)) > 0 Then
                If tot >= liveus(角色人物對戰人數(1, 2)) Then
                    戰鬥系統類.廣播訊息 "您受到了" & liveus(角色人物對戰人數(1, 2)) & "點傷害。"
                    FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = 0
                    liveus(角色人物對戰人數(1, 2)) = 0
                    FormMainMode.bloodnumus1.Caption = 0
                    FormMainMode.bloodlineout1.Width = 0
                    牌總階段數(1) = 牌總階段數(1) + 1
                Else
                    FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = liveus(角色人物對戰人數(1, 2)) - tot
                    liveus(角色人物對戰人數(1, 2)) = liveus(角色人物對戰人數(1, 2)) - tot
                    FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
                    FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (距離單位(1, 1, 1) * tot)
                    戰鬥系統類.廣播訊息 "您受到了" & tot & "點傷害。"
                End If
                FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).CurrentHP = liveus(角色人物對戰人數(1, 2))
                戰鬥系統類.播放傷害音樂
            End If
        Case Is > 1
            If tot > 0 And liveus(角色待機人物紀錄數(1, num)) > 0 Then
                If tot >= liveus(角色待機人物紀錄數(1, num)) Then
                    liveus(角色待機人物紀錄數(1, num)) = 0
                    FormMainMode.cardus(角色待機人物紀錄數(1, num)).CardMain_角色HP = 0
                    牌總階段數(1) = 牌總階段數(1) + 1
                Else
                    liveus(角色待機人物紀錄數(1, num)) = liveus(角色待機人物紀錄數(1, num)) - tot
                    FormMainMode.cardus(角色待機人物紀錄數(1, num)).CardMain_角色HP = liveus(角色待機人物紀錄數(1, num))
                End If
                FormMainMode.PEAFpersoncardus(角色待機人物紀錄數(1, num)).CurrentHP = liveus(角色待機人物紀錄數(1, num))
            End If
    End Select

End Sub
Sub 骰量更新顯示()
    攻擊防禦骰子總數(1) = 0
    攻擊防禦骰子總數(2) = 0
    Erase 顯示列雙方數值鎖定紀錄數
    Erase atkingckdice
    Erase Vss_EventPersonAbilityDiceChangeNum
    Dim uscom As Integer
    '===========================執行階段插入點(45)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 45, 1
    '============================
    For uscom = 1 To 2
        Select Case uscom
            Case 1
                If turnatk = 1 Then
                    If atkingpagetot(1, 1) > 0 And movecp = 1 Then
                        If Vss_EventPersonAbilityDiceChangeNum(1, 2) = 0 Then
                            攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkus(角色人物對戰人數(1, 2))
                        End If
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Vss_EventPersonAbilityDiceChangeNum(1, 1)
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkingpagetot(1, 1)
                    ElseIf atkingpagetot(1, 5) > 0 And movecp > 1 Then
                        If Vss_EventPersonAbilityDiceChangeNum(1, 2) = 0 Then
                            攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkus(角色人物對戰人數(1, 2))
                        End If
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Vss_EventPersonAbilityDiceChangeNum(1, 1)
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkingpagetot(1, 5)
                    End If
                ElseIf turnatk = 2 Then
                    If Vss_EventPersonAbilityDiceChangeNum(1, 2) = 0 Then
                        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + defus(角色人物對戰人數(1, 2))
                    End If
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + Vss_EventPersonAbilityDiceChangeNum(1, 1)
                    攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkingpagetot(1, 2)
                End If
                '=======主動技
                解析骰量變化 atkingckdice(1, 1, 1), 1
                '=======被動技
                解析骰量變化 atkingckdice(1, 1, 2), 1
                '=======異常狀態
                解析骰量變化 atkingckdice(1, 1, 3), 1
                '=======人物實際狀態
                解析骰量變化 atkingckdice(1, 1, 4), 1
                '=================================對手
                '=======主動技
                解析骰量變化 atkingckdice(2, 1, 1), 1
                '=======被動技
                解析骰量變化 atkingckdice(2, 1, 2), 1
                '=======異常狀態
                解析骰量變化 atkingckdice(2, 1, 3), 1
                '=======人物實際狀態
                解析骰量變化 atkingckdice(2, 1, 4), 1
                '===================================
                '            FormMainMode.trgoi1_Timer
            Case 2
                If turnatk = 2 Then
                    If atkingpagetot(2, 1) > 0 And movecp = 1 Then
                        If Vss_EventPersonAbilityDiceChangeNum(2, 2) = 0 Then
                            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkcom(角色人物對戰人數(2, 2))
                        End If
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Vss_EventPersonAbilityDiceChangeNum(2, 1)
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkingpagetot(2, 1)
                    ElseIf atkingpagetot(2, 5) > 0 And movecp > 1 Then
                        If Vss_EventPersonAbilityDiceChangeNum(2, 2) = 0 Then
                            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkcom(角色人物對戰人數(2, 2))
                        End If
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Vss_EventPersonAbilityDiceChangeNum(2, 1)
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkingpagetot(2, 5)
                    End If
                ElseIf turnatk = 1 Then
                    If Vss_EventPersonAbilityDiceChangeNum(2, 2) = 0 Then
                        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + defcom(角色人物對戰人數(2, 2))
                    End If
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Vss_EventPersonAbilityDiceChangeNum(2, 1)
                    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkingpagetot(2, 2)
                End If
                '=======主動技
                解析骰量變化 atkingckdice(2, 2, 1), 2
                '=======被動技
                解析骰量變化 atkingckdice(2, 2, 2), 2
                '=======異常狀態
                解析骰量變化 atkingckdice(2, 2, 3), 2
                '=======人物實際狀態
                解析骰量變化 atkingckdice(2, 2, 4), 2
                '=================================對手
                '=======主動技
                解析骰量變化 atkingckdice(1, 2, 1), 2
                '=======被動技
                解析骰量變化 atkingckdice(1, 2, 2), 2
                '=======異常狀態
                解析骰量變化 atkingckdice(1, 2, 3), 2
                '=======人物實際狀態
                解析骰量變化 atkingckdice(1, 2, 4), 2
                '===================================
        End Select
    Next
End Sub

Sub 播放傷害音樂()
    Select Case movecp
        Case 1
            一般系統類.音效播放 2
        Case Is >= 2
            一般系統類.音效播放 8
    End Select
End Sub
Sub 回復執行_使用者(ByVal tot As Integer, ByVal num As Integer, ByVal statusfrom As Integer, ByVal isEvent As Boolean, ByVal isSysCall As Boolean)
    If isEvent = True Then
        Dim stageInfoListObj As clsVSStageObj
        Dim tmpflagoff As Boolean
        If isSysCall = True Then
            Set stageInfoListObj = New clsVSStageObj
            stageInfoListObj.StageNum = 0
            stageInfoListObj.CommandStr = "@SystemHPLAction"
            stageInfoListObj.Value = "0"
            執行階段系統類.VBEVSStageInfoList.Add stageInfoListObj
        Else
            Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
        End If
        '===============================
        If statusfrom = 0 Then
            ReDim VBEStageNum(0 To 5) As Integer
            VBEStageNum(4) = 0    '觸發事件方
            VBEStageNum(5) = 0    '觸發事件體系
        End If
        VBEStageNum(0) = 48
        VBEStageNum(1) = -1    '回復方(1.使用者/2.電腦)
        VBEStageNum(2) = num    '回復人物編號
        VBEStageNum(3) = tot    '回復之數值
        stageInfoListObj.Argument = tot    '回復之數值
        '===========================執行階段插入點(48)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 48, 1
        '============================
        tmpflagoff = False
        If stageInfoListObj.CommandStr = "PersonBloodControl" Or (stageInfoListObj.CommandStr = "@SystemHPLAction" And isSysCall = True) Then
            If stageInfoListObj.Value = "HPLOFF" Then
                tmpflagoff = True
            Else
                Dim tmpstr() As String
                tmpstr = Split(stageInfoListObj.Value, "%")
                If UBound(tmpstr) = 1 And tmpstr(0) = "HPLCHANGE" Then
                    tot = Val(tmpstr(1))
                End If
            End If
        End If
        If isSysCall = True Then
            執行階段系統類.VBEVSStageInfoList.Remove 執行階段系統類.VBEVSStageInfoList.Count
        End If
        '===============================
        If tmpflagoff = True Then Exit Sub
        '===============================
    End If

    Select Case num
        Case 1
            If liveus(角色人物對戰人數(1, 2)) > 0 And tot > 0 Then
                If liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2)) >= tot Then
                    戰鬥系統類.廣播訊息 "你的HP恢復了" & tot & "點。"
                    FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + 距離單位(1, 1, 1) * tot
                    liveus(角色人物對戰人數(1, 2)) = Val(liveus(角色人物對戰人數(1, 2))) + tot
                ElseIf liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2)) < tot Then
                    If liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2)) > 0 Then
                        戰鬥系統類.廣播訊息 "你的HP恢復了" & Val(liveusmax(角色人物對戰人數(1, 2))) - Val(liveus(角色人物對戰人數(1, 2))) & "點。"
                        FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + 距離單位(1, 1, 1) * (Val(liveusmax(角色人物對戰人數(1, 2))) - Val(liveus(角色人物對戰人數(1, 2))))
                        liveus(角色人物對戰人數(1, 2)) = Val(liveusmax(角色人物對戰人數(1, 2)))
                    End If
                End If
                FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = liveus(角色人物對戰人數(1, 2))
                FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).CurrentHP = liveus(角色人物對戰人數(1, 2))
                FormMainMode.bloodnumus1.Caption = liveus(角色人物對戰人數(1, 2))
            End If
        Case Is > 1
            If liveus(角色待機人物紀錄數(1, num)) > 0 And tot > 0 Then
                If liveusmax(角色待機人物紀錄數(1, num)) - liveus(角色待機人物紀錄數(1, num)) >= tot Then
                    liveus(角色待機人物紀錄數(1, num)) = Val(liveus(角色待機人物紀錄數(1, num))) + tot
                ElseIf liveusmax(角色待機人物紀錄數(1, num)) - liveus(角色待機人物紀錄數(1, num)) < tot Then
                    liveus(角色待機人物紀錄數(1, num)) = Val(liveusmax(角色待機人物紀錄數(1, num)))
                End If
                FormMainMode.cardus(角色待機人物紀錄數(1, num)).CardMain_角色HP = liveus(角色待機人物紀錄數(1, num))
                FormMainMode.PEAFpersoncardus(角色待機人物紀錄數(1, num)).CurrentHP = liveus(角色待機人物紀錄數(1, num))
            End If
    End Select
End Sub
Sub 回復執行_電腦(ByVal tot As Integer, ByVal num As Integer, ByVal statusfrom As Integer, ByVal isEvent As Boolean, ByVal isSysCall As Boolean)
    If isEvent = True Then
        Dim stageInfoListObj As clsVSStageObj
        Dim tmpflagoff As Boolean
        If isSysCall = True Then
            Set stageInfoListObj = New clsVSStageObj
            stageInfoListObj.StageNum = 0
            stageInfoListObj.CommandStr = "@SystemHPLAction"
            stageInfoListObj.Value = "0"
            執行階段系統類.VBEVSStageInfoList.Add stageInfoListObj
        Else
            Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
        End If
        '===============================
        If statusfrom = 0 Then
            ReDim VBEStageNum(0 To 5) As Integer
            VBEStageNum(4) = 0    '觸發事件方
            VBEStageNum(5) = 0    '觸發事件體系
        End If
        VBEStageNum(0) = 48
        VBEStageNum(1) = -2    '回復方(系統代號)
        VBEStageNum(2) = num    '回復人物編號
        VBEStageNum(3) = tot    '回復之數值
        stageInfoListObj.Argument = tot    '回復之數值
        '===========================執行階段插入點(48)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 48, 1
        '============================
        tmpflagoff = False
        If stageInfoListObj.CommandStr = "PersonBloodControl" Or (stageInfoListObj.CommandStr = "@SystemHPLAction" And isSysCall = True) Then
            If stageInfoListObj.Value = "HPLOFF" Then
                tmpflagoff = True
            Else
                Dim tmpstr() As String
                tmpstr = Split(stageInfoListObj.Value, "%")
                If UBound(tmpstr) = 1 And tmpstr(0) = "HPLCHANGE" Then
                    tot = Val(tmpstr(1))
                End If
            End If
        End If
        If isSysCall = True Then
            執行階段系統類.VBEVSStageInfoList.Remove 執行階段系統類.VBEVSStageInfoList.Count
        End If
        '===============================
        If tmpflagoff = True Then Exit Sub
        '===============================
    End If

    Select Case num
        Case 1
            If livecom(角色人物對戰人數(2, 2)) > 0 And tot > 0 Then
                If livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2)) >= tot Then
                    戰鬥系統類.廣播訊息 "對方的HP恢復了" & tot & "點。"
                    FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - 距離單位(1, 2, 1) * tot
                    livecom(角色人物對戰人數(2, 2)) = Val(livecom(角色人物對戰人數(2, 2))) + tot
                ElseIf livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2)) < tot Then
                    If livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2)) > 0 Then
                        戰鬥系統類.廣播訊息 "對方的HP恢復了" & Val(livecommax(角色人物對戰人數(2, 2))) - Val(livecom(角色人物對戰人數(2, 2))) & "點。"
                        FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - 距離單位(1, 2, 1) * (Val(livecommax(角色人物對戰人數(2, 2))) - Val(livecom(角色人物對戰人數(2, 2))))
                        livecom(角色人物對戰人數(2, 2)) = Val(livecommax(角色人物對戰人數(2, 2)))
                    End If
                End If
                FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).CurrentHP = livecom(角色人物對戰人數(2, 2))
                FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = livecom(角色人物對戰人數(2, 2))
                FormMainMode.bloodnumcom1.Caption = livecom(角色人物對戰人數(2, 2))
            End If
        Case Is > 1
            If livecom(角色待機人物紀錄數(2, num)) > 0 And tot > 0 Then
                If livecommax(角色待機人物紀錄數(2, num)) - livecom(角色待機人物紀錄數(2, num)) >= tot Then
                    livecom(角色待機人物紀錄數(2, num)) = Val(livecom(角色待機人物紀錄數(2, num))) + tot
                ElseIf livecommax(角色待機人物紀錄數(2, num)) - livecom(角色待機人物紀錄數(2, num)) < tot Then
                    livecom(角色待機人物紀錄數(2, num)) = Val(livecommax(角色待機人物紀錄數(2, num)))
                End If
                FormMainMode.cardcom(角色待機人物紀錄數(2, num)).CardMain_角色HP = livecom(角色待機人物紀錄數(2, num))
                FormMainMode.PEAFpersoncardcom(角色待機人物紀錄數(2, num)).CurrentHP = livecom(角色待機人物紀錄數(2, num))
            End If
    End Select
End Sub
Sub 傷害執行_使用者(ByVal tot As Integer)
    If tot <= 0 Then Exit Sub
    '===============================
    Dim stageInfoListObj As New clsVSStageObj
    Dim tmpflagoff As Boolean
    stageInfoListObj.StageNum = 0
    stageInfoListObj.CommandStr = "@SystemBloodAction"
    stageInfoListObj.Value = "0"
    執行階段系統類.VBEVSStageInfoList.Add stageInfoListObj
    '===============================
    ReDim VBEStageNum(0 To 6) As Integer
    VBEStageNum(0) = 46
    VBEStageNum(1) = -1    '受到傷害方(1.使用者/2.電腦)
    VBEStageNum(2) = 1    '受到傷害人物編號
    VBEStageNum(3) = 1    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    VBEStageNum(4) = tot    '受到傷害之數值
    VBEStageNum(5) = 0    '來自系統的傷害
    VBEStageNum(6) = 0    '來自系統的傷害
    stageInfoListObj.Argument = tot  '受到傷害之數值
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害方(1.使用者/2.電腦)
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害人物編號
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    '===========================執行階段插入點(46)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 46, 1
    '============================
    tmpflagoff = False
    If stageInfoListObj.CommandStr = "@SystemBloodAction" Then
        If stageInfoListObj.Value = "BLOODOFF" Then
            tmpflagoff = True
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                tot = Val(tmpstr(1))
            End If
        End If
    End If
    執行階段系統類.VBEVSStageInfoList.Remove 執行階段系統類.VBEVSStageInfoList.Count
    '===============================
    If tmpflagoff = True Then Exit Sub
    '===============================
    If tot > 0 And liveus(角色人物對戰人數(1, 2)) > 0 Then
        If tot >= liveus(角色人物對戰人數(1, 2)) Then
            戰鬥系統類.廣播訊息 "您受到了" & liveus(角色人物對戰人數(1, 2)) & "點傷害。"
            FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = 0
            liveus(角色人物對戰人數(1, 2)) = 0
            FormMainMode.bloodnumus1.Caption = 0
            FormMainMode.bloodlineout1.Width = 0
            牌總階段數(1) = 牌總階段數(1) + 1
        Else
            FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = liveus(角色人物對戰人數(1, 2)) - tot
            liveus(角色人物對戰人數(1, 2)) = liveus(角色人物對戰人數(1, 2)) - tot
            FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
            FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (距離單位(1, 1, 1) * tot)
            戰鬥系統類.廣播訊息 "您受到了" & tot & "點傷害。"
        End If
        FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).CurrentHP = liveus(角色人物對戰人數(1, 2))
        戰鬥系統類.播放傷害音樂
    End If
End Sub
Sub 傷害執行_技能直傷_電腦(ByVal tot As Integer, ByVal num As Integer, ByVal isEvent As Boolean)
    If tot <= 0 Then Exit Sub
    If isEvent = True Then
        Dim stageInfoListObj As clsVSStageObj
        Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
        '===============================
        VBEStageNum(0) = 46
        VBEStageNum(1) = -2    '受到傷害方(1.使用者/2.電腦)
        VBEStageNum(2) = num    '受到傷害人物編號
        VBEStageNum(3) = 2    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
        VBEStageNum(4) = tot    '受到傷害之數值
        stageInfoListObj.Argument = tot  '受到傷害之數值
        stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2"    '受到傷害方(1.使用者/2.電腦)
        stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num)    '受到傷害人物編號
        stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2"    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
        '===========================執行階段插入點(46)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 46, 1
        '============================
        If stageInfoListObj.CommandStr = "PersonBloodControl" Then
            If stageInfoListObj.Value = "BLOODOFF" Then
                Exit Sub
            Else
                Dim tmpstr() As String
                tmpstr = Split(stageInfoListObj.Value, "%")
                If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                    tot = Val(tmpstr(1))
                End If
            End If
        End If
    End If
    Select Case num
        Case 1
            If tot > 0 And livecom(角色人物對戰人數(2, 2)) > 0 Then
                If tot >= livecom(角色人物對戰人數(2, 2)) Then
                    戰鬥系統類.廣播訊息 "對方受到了" & livecom(角色人物對戰人數(2, 2)) & "點傷害。"
                    FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = 0
                    FormMainMode.bloodnumcom1.Caption = 0
                    livecom(角色人物對戰人數(2, 2)) = 0
                    FormMainMode.bloodlineout2.Left = 11580
                    牌總階段數(2) = 牌總階段數(2) + 1
                Else
                    戰鬥系統類.廣播訊息 "對方受到了" & Val(tot) & "點傷害。"
                    FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = livecom(角色人物對戰人數(2, 2)) - tot
                    FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
                    livecom(角色人物對戰人數(2, 2)) = livecom(角色人物對戰人數(2, 2)) - tot
                    FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (距離單位(1, 2, 1) * tot)
                End If
                FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).CurrentHP = livecom(角色人物對戰人數(2, 2))
                戰鬥系統類.播放傷害音樂
            End If
        Case Is > 1
            If tot > 0 And livecom(角色待機人物紀錄數(2, num)) > 0 Then
                If tot >= livecom(角色待機人物紀錄數(2, num)) Then
                    livecom(角色待機人物紀錄數(2, num)) = 0
                    FormMainMode.cardcom(角色待機人物紀錄數(2, num)).CardMain_角色HP = 0
                    牌總階段數(2) = 牌總階段數(2) + 1
                Else
                    livecom(角色待機人物紀錄數(2, num)) = livecom(角色待機人物紀錄數(2, num)) - tot
                    FormMainMode.cardcom(角色待機人物紀錄數(2, num)).CardMain_角色HP = livecom(角色待機人物紀錄數(2, num))
                End If
                FormMainMode.PEAFpersoncardcom(角色待機人物紀錄數(2, num)).CurrentHP = livecom(角色待機人物紀錄數(2, num))
            End If
    End Select
End Sub
Sub 傷害執行_電腦(ByVal tot As Integer)
    If tot <= 0 Then Exit Sub
    '===============================
    Dim stageInfoListObj As New clsVSStageObj
    Dim tmpflagoff As Boolean
    stageInfoListObj.StageNum = 0
    stageInfoListObj.CommandStr = "@SystemBloodAction"
    stageInfoListObj.Value = "0"
    執行階段系統類.VBEVSStageInfoList.Add stageInfoListObj
    '===============================
    ReDim VBEStageNum(0 To 6) As Integer
    VBEStageNum(0) = 46
    VBEStageNum(1) = -2    '受到傷害方(1.使用者/2.電腦)
    VBEStageNum(2) = 1    '受到傷害人物編號
    VBEStageNum(3) = 1    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    VBEStageNum(4) = tot    '受到傷害之數值
    VBEStageNum(5) = 0    '來自系統的傷害
    VBEStageNum(6) = 0    '來自系統的傷害
    stageInfoListObj.Argument = tot  '受到傷害之數值
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害方(1.使用者/2.電腦)
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害人物編號
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    '===========================執行階段插入點(46)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 46, 1
    '============================
    tmpflagoff = False
    If stageInfoListObj.CommandStr = "@SystemBloodAction" Then
        If stageInfoListObj.Value = "BLOODOFF" Then
            tmpflagoff = True
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                tot = Val(tmpstr(1))
            End If
        End If
    End If
    執行階段系統類.VBEVSStageInfoList.Remove 執行階段系統類.VBEVSStageInfoList.Count
    '===============================
    If tmpflagoff = True Then Exit Sub
    '============================
    If tot > 0 And livecom(角色人物對戰人數(2, 2)) > 0 Then
        If tot >= livecom(角色人物對戰人數(2, 2)) Then
            戰鬥系統類.廣播訊息 "對方受到了" & livecom(角色人物對戰人數(2, 2)) & "點傷害。"
            FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = 0
            FormMainMode.bloodnumcom1.Caption = 0
            livecom(角色人物對戰人數(2, 2)) = 0
            FormMainMode.bloodlineout2.Left = 11580
            牌總階段數(2) = 牌總階段數(2) + 1
        Else
            戰鬥系統類.廣播訊息 "對方受到了" & Val(tot) & "點傷害。"
            FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = livecom(角色人物對戰人數(2, 2)) - tot
            FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
            livecom(角色人物對戰人數(2, 2)) = livecom(角色人物對戰人數(2, 2)) - tot
            FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (距離單位(1, 2, 1) * tot)
        End If
        FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).CurrentHP = livecom(角色人物對戰人數(2, 2))
        戰鬥系統類.播放傷害音樂
    End If
End Sub
Sub 執行動作_使用者_棄牌(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)))(CStr(n))

    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) - 1
    目前數(5) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Location = 3
    牌移動暫時變數(1) = 240
    牌移動暫時變數(2) = 960
    牌移動暫時變數(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '指定目前Left(座標)
    tmpcard.XYTop = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    目前數(15) = 4
    Select Case tmpcard.CardType
        Case 1    '公用牌
            戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 2
        Case 2    '事件卡
            戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 9
    End Select
    FormMainMode.牌移動.Enabled = True
    一般系統類.音效播放 1
End Sub
Sub 執行動作_牌組_回牌_使用者(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)))(CStr(n))

    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    tmpcard.Owner = 1
    tmpcard.Location = 1
    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '指定目前Left(座標)
    tmpcard.XYTop = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.公用牌回復正面 n
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.牌順序增加_手牌_使用者
    戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 5
    FormMainMode.牌移動.Enabled = True
    一般系統類.音效播放 1
End Sub
Sub 執行動作_電腦牌_偷牌_使用者(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)))(CStr(n))

    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    目前數(9) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Owner = 1
    tmpcard.Location = 1
    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '指定目前Left(座標)
    tmpcard.XYTop = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.牌順序增加_手牌_使用者
    戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 5
    目前數(15) = 2
    FormMainMode.牌移動.Enabled = True
    一般系統類.音效播放 1
End Sub
Sub 執行動作_使用者牌_偷牌_電腦(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)))(CStr(n))

    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pageusglead = Val(FormMainMode.pageusglead) - 1
    目前數(5) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Owner = 2
    tmpcard.Location = 1
    戰鬥系統類.座標計算_電腦手牌
    牌移動暫時變數(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '指定目前Left(座標)
    tmpcard.XYTop = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.牌順序增加_手牌_電腦
    戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 7
    目前數(15) = 20
    戰鬥系統類.公用牌變背面
    FormMainMode.牌移動.Enabled = True
    一般系統類.音效播放 1
End Sub
Sub 執行動作_牌組_回牌_電腦(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)))(CStr(n))

    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    tmpcard.Owner = 2
    tmpcard.Location = 1
    戰鬥系統類.座標計算_電腦手牌
    牌移動暫時變數(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '指定目前Left(座標)
    tmpcard.XYTop = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.牌順序增加_手牌_電腦
    戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 7
    戰鬥系統類.公用牌變背面
    FormMainMode.牌移動.Enabled = True
    一般系統類.音效播放 1
End Sub
Sub 執行動作_翻牌(ByVal n As Integer)
    FormMainMode.card(n).Width = 810
    FormMainMode.card(n).Height = 1260
    FormMainMode.card(n).LocationType = 1
    FormMainMode.card(n).CardEventType = False
    FormMainMode.card(n).Visible = True
    一般系統類.音效播放 4
End Sub
Sub 座標計算_電腦出牌()
    Dim xy As Long  '暫時變數(首牌Left)
    If pageqlead(2) = 1 Then
        牌移動暫時變數(1) = 5260
        牌移動暫時變數(2) = 1120
    ElseIf pageqlead(2) > 1 Then
        xy = (pageqlead(2) - 1) * 460
        牌移動暫時變數(1) = (Val(5260) - xy) + ((pageqlead(2) - 1) * Val(960))
        牌移動暫時變數(2) = 1120
    End If

End Sub
Sub 座標計算_電腦手牌()
    牌移動暫時變數(1) = 10560 - 240 * (Val(FormMainMode.pagecomglead) - 1)    '計算Left座標
    牌移動暫時變數(2) = -600    '指定Top座標
End Sub
Sub 座標計算_使用者出牌()
    Dim xy As Long   '暫時變數(首牌Left)
    If pageqlead(1) = 1 Then
        牌移動暫時變數(1) = 5260
        牌移動暫時變數(2) = 4840
    ElseIf pageqlead(1) > 1 Then
        xy = (pageqlead(1) - 1) * 460
        牌移動暫時變數(1) = (Val(5260) - xy) + ((pageqlead(1) - 1) * Val(960))
        牌移動暫時變數(2) = 4840
    End If

End Sub
Sub 座標計算_使用者手牌()
    If Val(FormMainMode.pageusglead) <= 9 Then
        牌移動暫時變數(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 1)    '計算Left座標
    Else
        牌移動暫時變數(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 10)
    End If

    If Val(FormMainMode.pageusglead) <= 9 Then
        牌移動暫時變數(2) = 6700    '指定Top座標
    Else
        牌移動暫時變數(2) = 7980    '指定Top座標
    End If
End Sub
Sub 牌順序增加_出牌_電腦()
    pagecomleadmax(1) = pagecomleadmax(1) + 1
End Sub
Sub 牌順序增加_手牌_電腦()
    pagecomleadmax(0) = pagecomleadmax(0) + 1
End Sub
Sub 牌順序增加_手牌_使用者()
    pageusleadmax(0) = pageusleadmax(0) + 1
End Sub
Sub 牌順序增加_出牌_使用者()
    pageusleadmax(1) = pageusleadmax(1) + 1
End Sub
Sub 執行動作_電腦_棄牌(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)))(CStr(n))

    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) - 1
    目前數(9) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Location = 3
    牌移動暫時變數(1) = 240
    牌移動暫時變數(2) = 960
    牌移動暫時變數(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '指定目前Left(座標)
    tmpcard.XYTop = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    Select Case tmpcard.CardType
        Case 1    '公用牌
            戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 2
        Case 2    '事件卡
            戰鬥系統類.卡牌牌堆集合更換 tmpcard, 戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(n)), 9
    End Select
    目前數(15) = 5
    FormMainMode.牌移動.Enabled = True
    一般系統類.音效播放 1
End Sub
Sub 執行動作_洗牌()
    戰鬥系統類.執行動作_洗牌_墓地牌回牌
    戰鬥系統類.執行動作_洗牌_牌堆洗牌
End Sub
Sub 執行動作_洗牌_墓地牌回牌()
    Dim tmpcard As clsActionCard

    目前數(30) = 戰鬥系統類.CardDeckCollection(1).Count

    Do While 戰鬥系統類.CardDeckCollection(2).Count > 0
        Set tmpcard = 戰鬥系統類.CardDeckCollection(2)(1)

        tmpcard.Owner = 0
        tmpcard.Location = 4
        戰鬥系統類.卡牌牌堆集合更換 tmpcard, 2, 1
    Loop

    BattleCardNum = 戰鬥系統類.CardDeckCollection(1).Count
    戰鬥系統類.執行動作_系統總卡牌張數更新
End Sub
Sub 執行動作_洗牌_牌堆洗牌()
    Dim tmpnewCollection As New Collection
    Dim tmpcard As clsActionCard
    Dim i As Integer

    For i = 1 To 目前數(30)    '將既有卡牌保留於牌堆最上層
        Set tmpcard = 戰鬥系統類.CardDeckCollection(1)(1)

        tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
        戰鬥系統類.CardDeckCollection(1).Remove 1
    Next

    Do While 戰鬥系統類.CardDeckCollection(1).Count > 0
        Randomize
        i = Int(Rnd() * 戰鬥系統類.CardDeckCollection(1).Count) + 1
        Set tmpcard = 戰鬥系統類.CardDeckCollection(1)(i)

        tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
        戰鬥系統類.CardDeckCollection(1).Remove i
    Loop

    Set 戰鬥系統類.CardDeckCollection(1) = Nothing
    Set 戰鬥系統類.CardDeckCollection(1) = tmpnewCollection

End Sub
Sub 執行動作_洗牌_事件卡牌堆洗牌()
    Dim tmpnewCollection As Collection
    Dim tmpcard As clsActionCard
    Dim i As Integer, m As Integer, tmptot As Integer

    For m = 3 To 4
        Set tmpnewCollection = New Collection

        If (m = 3 And 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgreus.Value = 1) Or _
           (m = 4 And 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1) Then    '1v1-遵守規則
            For tmptot = 6 To 1 Step -1
                Randomize
                i = Int(Rnd() * tmptot) + 1
                Set tmpcard = 戰鬥系統類.CardDeckCollection(m)(i)

                tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
                戰鬥系統類.CardDeckCollection(m).Remove i
            Next
        End If

        Do While 戰鬥系統類.CardDeckCollection(m).Count > 0
            Randomize
            i = Int(Rnd() * 戰鬥系統類.CardDeckCollection(m).Count) + 1
            Set tmpcard = 戰鬥系統類.CardDeckCollection(m)(i)

            tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
            戰鬥系統類.CardDeckCollection(m).Remove i
        Loop

        Set 戰鬥系統類.CardDeckCollection(m) = Nothing
        Set 戰鬥系統類.CardDeckCollection(m) = tmpnewCollection
    Next

End Sub
Sub 執行動作_抽牌_公用牌(ByVal uscom As Integer, Optional ByVal cardindex As Integer, Optional ByVal isEvent As Boolean)
    Dim tmpcard As clsActionCard
    Dim stageInfoListObj As clsVSStageObj
    Dim n As Integer, i As Integer, tmpflag As Boolean

    If cardindex > 0 And cardindex <= 戰鬥系統類.CardDeckCollection(1).Count Then
        Set tmpcard = 戰鬥系統類.CardDeckCollection(1)(cardindex)
    Else
        Set tmpcard = 戰鬥系統類.CardDeckCollection(1)(1)
    End If
    '==================================隨機轉牌
    Randomize
    n = Int(Rnd() * 2) + 1
    If n = 2 Then
        Call tmpcard.Reverse
    End If
    FormMainMode.card(tmpcard.CardNum).CardRotationType = tmpcard.CardOnIn

    If isEvent = True Then
        Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)

        ReDim VBEStageNum(0 To 5) As Integer
        VBEStageNum(0) = 103
        VBEStageNum(1) = 戰鬥系統類.卡牌牌堆集合索引_Index(CStr(tmpcard.CardNum))
        VBEStageNum(2) = 執行階段系統類.執行階段系統_準備變數統合_pagecardnum_type(tmpcard.UpperType)    '受影響之卡牌正面類型
        VBEStageNum(3) = tmpcard.UpperNum    '受影響之卡牌正面數值
        VBEStageNum(4) = 執行階段系統類.執行階段系統_準備變數統合_pagecardnum_type(tmpcard.LowerType)    '受影響之卡牌反面類型
        VBEStageNum(5) = tmpcard.LowerNum    '受影響之卡牌反面數值
        '===========================執行階段插入點(103)
        執行階段系統類.執行階段系統總主要程序_執行階段開始 uscom, 103, 2
        '============================
        tmpflag = False
        If stageInfoListObj.CommandStr = "AtkingDrawCardsEvent" Then
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")

            Dim tmpbool(1 To 2) As Boolean
            For i = 0 To UBound(tmpstr)
                If tmpstr(i) = "OFF" Then
                    Vss_AtkingDrawCardsNum(1) = Vss_AtkingDrawCardsNum(2)
                    tmpflag = True
                    Exit For
                ElseIf tmpstr(i) = "AddOnce" And tmpbool(1) = False Then
                    Vss_AtkingDrawCardsNum(2) = Vss_AtkingDrawCardsNum(2) + 1
                    tmpbool(1) = True
                ElseIf tmpstr(i) = "Continue" And tmpbool(2) = False Then
                    Vss_AtkingDrawCardsNum(3) = Vss_AtkingDrawCardsNum(3) + 1
                    tmpbool(2) = True
                    tmpflag = True
                End If
            Next
        End If

        If tmpflag = True Then
            If 執行階段系統_搜尋正在執行之執行階段("AtkingDrawCards") <> 0 Then
                vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingDrawCards")) = 2    '(階段2)
            End If
            Exit Sub
        End If
    End If
    '=======================
    Select Case uscom
        Case 1    '使用者
            tmpcard.ComMark = 0
            tmpcard.Owner = 1
            tmpcard.Location = 1
            戰鬥系統類.卡牌牌堆集合更換 tmpcard, 1, 5
            '=================
            BattleCardNum = BattleCardNum - 1
            戰鬥系統類.執行動作_系統總卡牌張數更新
            FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
            戰鬥系統類.座標計算_使用者手牌
            牌移動暫時變數(3) = tmpcard.CardNum
            tmpcard.XYLeft = 240    '指定目前Left(座標)
            tmpcard.XYTop = 960    '指定目前Top(座標)
            FormMainMode.card(tmpcard.CardNum).Left = 240
            FormMainMode.card(tmpcard.CardNum).Top = 960
            戰鬥系統類.計算牌移動距離單位
            戰鬥系統類.公用牌回復正面 (牌移動暫時變數(3))
            FormMainMode.card(tmpcard.CardNum).CardEventType = False
            FormMainMode.card(tmpcard.CardNum).Visible = True
            FormMainMode.card(tmpcard.CardNum).ZOrder
            戰鬥系統類.牌順序增加_手牌_使用者
            FormMainMode.牌移動.Enabled = True
            一般系統類.音效播放 1
        Case 2    '電腦
            tmpcard.ComMark = 0
            tmpcard.Owner = 2
            tmpcard.Location = 1
            戰鬥系統類.卡牌牌堆集合更換 tmpcard, 1, 7
            '=================
            BattleCardNum = BattleCardNum - 1
            戰鬥系統類.執行動作_系統總卡牌張數更新
            FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
            戰鬥系統類.座標計算_電腦手牌
            牌移動暫時變數(3) = tmpcard.CardNum
            tmpcard.XYLeft = 240    '指定目前Left(座標)
            tmpcard.XYTop = 960    '指定目前Top(座標)
            FormMainMode.card(tmpcard.CardNum).Left = 240
            FormMainMode.card(tmpcard.CardNum).Top = 960
            戰鬥系統類.計算牌移動距離單位
            戰鬥系統類.公用牌變背面
            FormMainMode.card(tmpcard.CardNum).CardEventType = False
            FormMainMode.card(tmpcard.CardNum).Visible = True
            FormMainMode.card(tmpcard.CardNum).ZOrder
            戰鬥系統類.牌順序增加_手牌_電腦
            FormMainMode.牌移動.Enabled = True
            一般系統類.音效播放 1
    End Select
End Sub
Sub 執行動作_清除所有異常狀態_聖水(ByVal uscom As Integer, ByVal num As Integer)
    If 人物異常狀態列表(uscom, 角色待機人物紀錄數(uscom, num)).Count > 0 Then
        '==================
        執行階段系統類.執行階段73_指令_異常狀態控制_全部清除 uscom, num, True
        '==================
        Dim tempnum As Integer, i As Integer
        tempnum = 1
        For i = 1 To 人物異常狀態列表(uscom, 角色待機人物紀錄數(uscom, num)).Count
            If VBEStageRemoveBuffAllNum(i) = False Then
                人物異常狀態列表(uscom, 角色待機人物紀錄數(uscom, num)).Remove tempnum
            Else
                tempnum = tempnum + 1
            End If
        Next
        戰鬥系統類.異常狀態顯示更新 uscom
    End If
End Sub
Sub 執行動作_距離變更(ByVal m As Integer, ByVal isEvent As Boolean, ByVal isSysCall As Boolean)
    '===========================執行階段插入點(47)
    If isEvent = True Then
        Dim stageInfoListObj As clsVSStageObj
        Dim tmpflagoff As Boolean
        Dim tmpuscom As Integer
        If isSysCall = True Then
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(3) = 0  '觸發事件方
            '======================
            Set stageInfoListObj = New clsVSStageObj
            stageInfoListObj.StageNum = 0
            stageInfoListObj.CommandStr = "@SystemBattleMove"
            stageInfoListObj.Value = "0"
            執行階段系統類.VBEVSStageInfoList.Add stageInfoListObj
        Else
            Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
        End If
        VBEStageNum(0) = 47
        VBEStageNum(1) = movecp    '變更前距離
        VBEStageNum(2) = m  '變更後距離
        If isSysCall = True Then
            tmpuscom = 1
        Else
            tmpuscom = Abs(VBEStageNum(3))
        End If
        執行階段系統類.執行階段系統總主要程序_執行階段開始 tmpuscom, 47, 1
        '=====================
        tmpflagoff = False
        If stageInfoListObj.CommandStr = "BattleMoveControl" Or (stageInfoListObj.CommandStr = "@SystemBattleMove" And isSysCall = True) Then
            If stageInfoListObj.Value = "BMCOFF" Then
                tmpflagoff = True
            End If
        End If
        If isSysCall = True Then
            執行階段系統類.VBEVSStageInfoList.Remove 執行階段系統類.VBEVSStageInfoList.Count
        End If
        '===============================
        If tmpflagoff = True Then Exit Sub
        '===============================
    End If
    '============================
    Dim anw(1 To 2) As Integer
    Dim anh(1 To 2) As Integer
    anw(1) = Val(FormMainMode.personusminijpg.小人物圖片width) / 2
    anw(2) = Val(FormMainMode.personcomminijpg.小人物圖片width) / 2
    anh(1) = Val(FormMainMode.personusminijpg.小人物圖片height)
    anh(2) = Val(FormMainMode.personcomminijpg.小人物圖片height)
    Select Case m
        Case 1
            FormMainMode.PEAFMoveRange.LoadImage_FromFile app_path & "\gif\system\short.png"
            FormMainMode.PEAFMoveRange.Left = 4440
            FormMainMode.PEAFMoveRange.Top = 2520
            FormMainMode.personusminijpg.Left = 4320 - anw(1)
            FormMainMode.personusminijpg.Top = 5880 - anh(1)
            FormMainMode.personcomminijpg.Left = 7080 - anw(2)
            FormMainMode.personcomminijpg.Top = 5880 - anh(2)
        Case 2
            FormMainMode.PEAFMoveRange.LoadImage_FromFile app_path & "\gif\system\middle.png"
            FormMainMode.PEAFMoveRange.Left = 2880
            FormMainMode.PEAFMoveRange.Top = 2000
            FormMainMode.personusminijpg.Left = 2640 - anw(1)
            FormMainMode.personusminijpg.Top = 5880 - anh(1)
            FormMainMode.personcomminijpg.Left = 8680 - anw(2)
            FormMainMode.personcomminijpg.Top = 5880 - anh(2)
        Case 3
            FormMainMode.PEAFMoveRange.LoadImage_FromFile app_path & "\gif\system\long.png"
            FormMainMode.PEAFMoveRange.Left = 1080
            FormMainMode.PEAFMoveRange.Top = 2360
            FormMainMode.personusminijpg.Left = 1040 - anw(1)
            FormMainMode.personusminijpg.Top = 5880 - anh(1)
            FormMainMode.personcomminijpg.Left = 10320 - anw(2)
            FormMainMode.personcomminijpg.Top = 5880 - anh(2)
    End Select

    movecp = m
End Sub
Sub 計算牌移動距離單位()
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(牌移動暫時變數(3))))(CStr(牌移動暫時變數(3)))

    If 牌移動暫時變數(1) >= tmpcard.XYLeft Then
        距離單位(2, 1, 1) = (牌移動暫時變數(1) - tmpcard.XYLeft) \ 8
    Else
        距離單位(2, 1, 1) = -((tmpcard.XYLeft - 牌移動暫時變數(1)) \ 8)
    End If

    If 牌移動暫時變數(2) >= tmpcard.XYTop Then
        距離單位(2, 1, 2) = (牌移動暫時變數(2) - tmpcard.XYTop) \ 8
    Else
        距離單位(2, 1, 2) = -((tmpcard.XYTop - 牌移動暫時變數(2)) \ 8)
    End If
End Sub
Sub 異常狀態顯示更新(ByVal uscom As Integer)
    Dim numNow As Integer, obj As clsStatus
    Dim i As Integer, k As Integer

    For i = 1 To 角色人物對戰人數(uscom, 1)
        numNow = 1
        For Each obj In 人物異常狀態列表(uscom, 角色待機人物紀錄數(uscom, i))
            Select Case uscom
                Case 1
                    FormMainMode.cardus(角色待機人物紀錄數(1, i)).更改異常狀態資料 numNow, obj.ImagePath, obj.Value, obj.Total, True
                Case 2
                    FormMainMode.cardcom(角色待機人物紀錄數(2, i)).更改異常狀態資料 numNow, obj.ImagePath, obj.Value, obj.Total, True
            End Select
            numNow = numNow + 1
            If numNow > 14 Then Exit For
        Next
        If numNow <= 14 Then
            For k = numNow To 14
                Select Case uscom
                    Case 1
                        FormMainMode.cardus(角色待機人物紀錄數(1, i)).更改異常狀態資料 k, 0, 0, 0, False
                    Case 2
                        FormMainMode.cardcom(角色待機人物紀錄數(2, i)).更改異常狀態資料 k, 0, 0, 0, False
                End Select
            Next
        End If
    Next

End Sub

Sub 直接寫入顯示列數值(ByVal n As Integer, ByVal num As Integer)
    If num < 0 Then num = 0
    Select Case n
        Case 1
            FormMainMode.顯示列1.goi1 = num
        Case 2
            FormMainMode.顯示列1.goi2 = num
    End Select
End Sub
Sub 小人物頭像執行完判斷_使用者()
    Dim ckl As Integer

    If turnatk = 1 Or turnatk = 2 Then
        turnpageonin = 1
        If Vss_EventPlayerAllActionOffNum(1) = 1 Then
            For ckl = 1 To 戰鬥系統類.ActionCardTotNum
                FormMainMode.card(ckl).CardEnabledType = False
            Next
            FormMainMode.PEAFInterface.BnOKEnabled False
            等待時間佇列(2).Add 47
            FormMainMode.等待時間_2.Enabled = True
        ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
            For ckl = 1 To 戰鬥系統類.ActionCardTotNum
                FormMainMode.card(ckl).CardEnabledType = False
            Next
            FormMainMode.PEAFInterface.BnOKEnabled False
            等待時間佇列(2).Add 45
            FormMainMode.等待時間_2.Enabled = True
        End If
    End If
    If turnatk = 3 Then
        FormMainMode.trtimeline.Enabled = True
    End If
End Sub
Sub 小人物頭像執行完判斷_電腦()
    If turnatk = 1 Or turnatk = 2 Or turnatk = 3 Then
        If Vss_EventPlayerAllActionOffNum(2) = 0 Then
            階段狀態數 = 3
            FormMainMode.電腦出牌.Enabled = True
        Else
            等待時間佇列(2).Add 48
            FormMainMode.等待時間_2.Enabled = True
        End If
    End If
End Sub
Sub 公用牌變背面()
    FormMainMode.card(牌移動暫時變數(3)).Width = 720
    FormMainMode.card(牌移動暫時變數(3)).Height = 990
    FormMainMode.card(牌移動暫時變數(3)).LocationType = 3
End Sub
Sub 公用牌回復正面(ByVal num As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(戰鬥系統類.卡牌牌堆集合索引_CollectionIndex(CStr(num)))(CStr(num))

    FormMainMode.card(num).Width = 810
    FormMainMode.card(num).Height = 1260
    FormMainMode.card(num).LocationType = 1
    FormMainMode.card(num).CardEventType = False
    FormMainMode.card(num).CardRotationType = tmpcard.CardOnIn
End Sub
Sub 收牌計算距離單位_使用者()
    Dim i As Integer
    Dim tmpcard As clsActionCard

    If 戰鬥系統類.CardDeckCollection(6).Count > 0 Then
        ReDim 距離單位_收牌暫時數(1 To 戰鬥系統類.CardDeckCollection(6).Count, 1 To 3) As Integer
    Else
        Erase 距離單位_收牌暫時數
    End If

    For i = 1 To 戰鬥系統類.CardDeckCollection(6).Count
        Set tmpcard = 戰鬥系統類.CardDeckCollection(6)(i)
        牌移動暫時變數(1) = 240
        牌移動暫時變數(2) = 960
        牌移動暫時變數(3) = tmpcard.CardNum
        tmpcard.XYLeft = FormMainMode.card(tmpcard.CardNum).Left  '指定目前Left(座標)
        tmpcard.XYTop = FormMainMode.card(tmpcard.CardNum).Top  '指定目前Top(座標)
        戰鬥系統類.計算牌移動距離單位
        距離單位_收牌暫時數(i, 1) = 距離單位(2, 1, 1)
        距離單位_收牌暫時數(i, 2) = 距離單位(2, 1, 2)
        距離單位_收牌暫時數(i, 3) = tmpcard.CardNum
    Next
End Sub
Sub 收牌計算距離單位_電腦()
    Dim i As Integer
    Dim tmpcard As clsActionCard

    If 戰鬥系統類.CardDeckCollection(8).Count > 0 Then
        ReDim 距離單位_收牌暫時數(1 To 戰鬥系統類.CardDeckCollection(8).Count, 1 To 3) As Integer
    Else
        Erase 距離單位_收牌暫時數
    End If

    For i = 1 To 戰鬥系統類.CardDeckCollection(8).Count
        Set tmpcard = 戰鬥系統類.CardDeckCollection(8)(i)
        牌移動暫時變數(1) = 240
        牌移動暫時變數(2) = 960
        牌移動暫時變數(3) = tmpcard.CardNum
        tmpcard.XYLeft = FormMainMode.card(tmpcard.CardNum).Left  '指定目前Left(座標)
        tmpcard.XYTop = FormMainMode.card(tmpcard.CardNum).Top  '指定目前Top(座標)
        戰鬥系統類.計算牌移動距離單位
        距離單位_收牌暫時數(i, 1) = 距離單位(2, 1, 1)
        距離單位_收牌暫時數(i, 2) = 距離單位(2, 1, 2)
        距離單位_收牌暫時數(i, 3) = tmpcard.CardNum
    Next
End Sub
Sub 技能說明載入_使用者()
    Dim i As Integer, ahmt As String, n As Integer
    Dim tmpobj As clsPersonActiveSkill

    For n = 1 To 4
        Set tmpobj = New clsPersonActiveSkill
        If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 1) = "" Then
            FormMainMode.PEAFInterface.ActiveDescription 1, n, tmpobj
            FormMainMode.PEAFInterface.ActiveSkillVisable 1, n, False
        Else
            tmpobj.name = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 1)

            If VBEPerson(1, 角色人物對戰人數(1, 2), 2, 3, 5) = 1 Then
                tmpobj.NameFontSize = 12
            Else
                tmpobj.NameFontSize = VBEPerson(1, 角色人物對戰人數(1, 2), 2, 3, n)
            End If

            tmpobj.Stage = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 2)
            tmpobj.Distance = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 3)
            tmpobj.card = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 4)
            ahmt = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 5)
            For i = 1 To Len(ahmt)
                If Mid(ahmt, i, 1) = "&" Then
                    Mid(ahmt, i, 1) = Chr(10)
                End If
            Next
            tmpobj.Effect = ahmt
            If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 6) <> "" Then
                tmpobj.cardFontSize = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 6))
            Else
                tmpobj.cardFontSize = 10
            End If
            If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 7) <> "" Then
                tmpobj.EffectFontSize = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 7))
            Else
                tmpobj.EffectFontSize = 10
            End If

            FormMainMode.PEAFInterface.ActiveDescription 1, n, tmpobj
            Set 戰鬥系統類.ActiveSkillObj(1, n) = tmpobj
            FormMainMode.PEAFInterface.ActiveSkillVisable 1, n, True
            If atkingck(1, 角色人物對戰人數(1, 2), n, 1) = 1 Then
                戰鬥系統類.人物技能欄燈開關 True, n
            End If
        End If
    Next
End Sub
Sub 技能說明載入_電腦()
    Dim i As Integer, ahmt As String, n As Integer
    Dim tmpobj As clsPersonActiveSkill

    For n = 1 To 4
        Set tmpobj = New clsPersonActiveSkill
        If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 1) = "" Then
            FormMainMode.PEAFInterface.ActiveDescription 2, n, tmpobj
            FormMainMode.PEAFInterface.ActiveSkillVisable 2, n, False
        Else
            tmpobj.name = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 1)

            If VBEPerson(2, 角色人物對戰人數(2, 2), 2, 3, 5) = 1 Then
                tmpobj.NameFontSize = 12
            Else
                tmpobj.NameFontSize = VBEPerson(2, 角色人物對戰人數(2, 2), 2, 3, n)
            End If

            tmpobj.Stage = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 2)
            tmpobj.Distance = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 3)
            tmpobj.card = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 4)
            ahmt = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 5)
            For i = 1 To Len(ahmt)
                If Mid(ahmt, i, 1) = "&" Then
                    Mid(ahmt, i, 1) = Chr(10)
                End If
            Next
            tmpobj.Effect = ahmt
            If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 6) <> "" Then
                tmpobj.cardFontSize = Val(VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 6))
            Else
                tmpobj.cardFontSize = 10
            End If
            If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 7) <> "" Then
                tmpobj.EffectFontSize = Val(VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 7))
            Else
                tmpobj.EffectFontSize = 10
            End If

            FormMainMode.PEAFInterface.ActiveDescription 2, n, tmpobj
            Set 戰鬥系統類.ActiveSkillObj(2, n) = tmpobj
            FormMainMode.PEAFInterface.ActiveSkillVisable 2, n, True
        End If
    Next
End Sub
Sub 音量靜音調節設定()
    Dim i As Integer

    If Formsetting.cksemute.Value = 1 Then
        For i = 1 To FormMainMode.cMusicPlayer.UBound
            FormMainMode.cMusicPlayer(i).Mute = True
        Next
    Else
        For i = 1 To FormMainMode.cMusicPlayer.UBound
            FormMainMode.cMusicPlayer(i).Mute = False
        Next
    End If
End Sub
Sub 時間軸_重設()
    FormMainMode.timelineout1.X1 = 0
    FormMainMode.timelineout2.X2 = 11310
    時間軸顏色變化紀錄暫時變數(1, 1) = 23
    時間軸顏色變化紀錄暫時變數(1, 2) = 77
    時間軸顏色變化紀錄暫時變數(1, 3) = 0
    時間軸顏色變化紀錄暫時變數(2, 1) = 0
    時間軸顏色變化紀錄暫時變數(2, 2) = 0
    時間軸顏色變化紀錄暫時變數(2, 3) = 0
    時間軸顏色變化紀錄暫時變數(3, 1) = 111
    時間軸顏色變化紀錄暫時變數(3, 2) = 251
    時間軸顏色變化紀錄暫時變數(3, 3) = 50
    FormMainMode.timelineout1.BorderColor = RGB(111, 251, 50)
    FormMainMode.timelineout2.BorderColor = RGB(111, 251, 50)
End Sub
Sub 時間軸_停止()
    FormMainMode.trtimeline.Enabled = False
    FormMainMode.timelinein1.BorderColor = RGB(0, 0, 0)
    FormMainMode.timelinein2.BorderColor = RGB(0, 0, 0)
End Sub
Sub 時間軸_隱藏()
    FormMainMode.timeup.Visible = False
    FormMainMode.timelinein1.Visible = False
    FormMainMode.timelinein2.Visible = False
    FormMainMode.timelineout1.Visible = False
    FormMainMode.timelineout2.Visible = False
End Sub
Sub 時間軸_顯示()
    FormMainMode.timeup.Visible = True
    FormMainMode.timelinein1.Visible = True
    FormMainMode.timelinein2.Visible = True
    FormMainMode.timelineout1.Visible = True
    FormMainMode.timelineout2.Visible = True
End Sub
Sub 階段執行判斷()
    If Val(擲骰表單溝通暫時變數(4)) = 1 Then
        Select Case Val(擲骰表單溝通暫時變數(1))
            Case 1
                If 擲骰表單溝通暫時變數(4) = 1 Then
                    擲骰表單溝通暫時變數(1) = 2
                    等待時間佇列(1).Add 14
                    FormMainMode.等待時間.Enabled = True
                Else
                    等待時間佇列(1).Add 15
                    FormMainMode.等待時間.Enabled = True
                End If
            Case 2
                If 擲骰表單溝通暫時變數(4) = 1 Then
                    等待時間佇列(1).Add 15
                    FormMainMode.等待時間.Enabled = True
                Else
                    擲骰表單溝通暫時變數(1) = 2
                    等待時間佇列(1).Add 13
                    FormMainMode.等待時間.Enabled = True
                End If
        End Select
    Else
        Select Case Val(擲骰表單溝通暫時變數(1))
            Case 1
                If 擲骰表單溝通暫時變數(4) = 1 Then
                    等待時間佇列(1).Add 15
                    FormMainMode.等待時間.Enabled = True
                Else
                    擲骰表單溝通暫時變數(1) = 2
                    等待時間佇列(1).Add 13
                    FormMainMode.等待時間.Enabled = True
                End If
            Case 2
                If 擲骰表單溝通暫時變數(4) = 1 Then
                    擲骰表單溝通暫時變數(1) = 2
                    等待時間佇列(1).Add 14
                    FormMainMode.等待時間.Enabled = True
                Else
                    等待時間佇列(1).Add 15
                    FormMainMode.等待時間.Enabled = True
                End If
        End Select
    End If
End Sub
Sub 電腦牌_模擬按牌(ByVal Index As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(CStr(Index))

    If tmpcard.Location = 1 And tmpcard.Owner = 2 Then
        tmpcard.Location = 2
        If tmpcard.UpperType = a1a Then
            atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + Val(tmpcard.UpperNum)
            If turnatk = 2 And movecp = 1 And 攻擊防禦骰子總數(4) = 0 Then
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
            End If
            If turnatk = 2 And movecp = 1 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(tmpcard.UpperNum)
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(tmpcard.UpperNum)
            End If
        End If
        If tmpcard.UpperType = a5a Then
            atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + Val(tmpcard.UpperNum)
            If turnatk = 2 And movecp > 1 And 攻擊防禦骰子總數(4) = 0 Then
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
            End If
            If turnatk = 2 And movecp > 1 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(tmpcard.UpperNum)
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(tmpcard.UpperNum)
            End If
        End If
        If tmpcard.UpperType = a2a Then
            atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + Val(tmpcard.UpperNum)
            If turnatk = 1 And 攻擊防禦骰子總數(4) = 0 Then
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + defcom(角色人物對戰人數(2, 2))
            End If
            If turnatk = 1 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(tmpcard.UpperNum)
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(tmpcard.UpperNum)
            End If
        End If
        If tmpcard.UpperType = a3a Then
            atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + Val(tmpcard.UpperNum)
        End If
        If tmpcard.UpperType = a4a Then
            atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + Val(tmpcard.UpperNum)
        End If
        '===================
        目前數(9) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(7), tmpcard)
        pagecomleadmax(1) = Val(pagecomleadmax(1)) + 1
        pageqlead(2) = Val(pageqlead(2)) + 1
        FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
        FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) + 1
        tmpcard.ComMark = 2
        '===================以下是出牌對齊
        目前數(7) = 0
        FormMainMode.電腦出牌_出牌對齊_靠左.Enabled = True
        '=============以下是牌移動(出牌)(電腦)
        戰鬥系統類.座標計算_電腦出牌
        牌移動暫時變數(3) = Index
        tmpcard.XYLeft = FormMainMode.card(Index).Left  '指定目前Left(座標)
        tmpcard.XYTop = FormMainMode.card(Index).Top  '指定目前Top(座標)
        戰鬥系統類.計算牌移動距離單位
        戰鬥系統類.卡牌牌堆集合更換 tmpcard, 7, 8
        目前數(15) = 0
        FormMainMode.牌移動.Enabled = True
        一般系統類.音效播放 1
        '================以下是手牌對齊
        目前數(8) = 0
        目前數(17) = 1
        FormMainMode.電腦出牌_手牌對齊.Enabled = True
        '===================以下是事件卡檢查及啟動
        If tmpcard.UpperType = a6a Then
            事件卡記錄暫時數(2, 3) = 1
            事件卡.機會_電腦 Index, tmpcard.UpperNum
        End If
        If turnatk = 1 Or turnatk = 2 Then
            If tmpcard.UpperType = a7a Then
                事件卡記錄暫時數(2, 3) = 1
                事件卡.詛咒術_電腦 Index, tmpcard.UpperNum
            End If
        End If
        If tmpcard.UpperType = a8a Then
            事件卡記錄暫時數(2, 3) = 1
            事件卡.HP回復_電腦 Index, tmpcard.UpperNum
        End If
        If tmpcard.UpperType = a9a Then
            事件卡記錄暫時數(2, 3) = 1
            事件卡.聖水_電腦 Index, tmpcard.UpperNum
        End If
        '==============================================
        Select Case turnatk
            Case 1
                '===========================執行階段插入點(ATK-42/DEF-43)
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 42
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 43
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
                '============================
            Case 2
                '===========================執行階段插入點(ATK-42/DEF-43)
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 42
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 43
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
                '============================
            Case 3
                '===========================執行階段插入點(44)
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 44
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
                '============================
        End Select
        戰鬥系統類.骰量更新顯示
    End If
End Sub
Sub 電腦牌_模擬按牌_外(ByVal Index As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(8)(CStr(Index))

    If tmpcard.Location = 2 And tmpcard.Owner = 2 Then
        tmpcard.Location = 1
        If tmpcard.UpperType = a1a Then
            atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - Val(tmpcard.UpperNum)
            If turnatk = 2 And movecp = 1 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(tmpcard.UpperNum)
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(tmpcard.UpperNum)
            End If
            If 攻擊防禦骰子總數(4) = atkcom(角色人物對戰人數(2, 2)) Then
                攻擊防禦骰子總數(4) = 0
            End If
        End If
        If tmpcard.UpperType = a5a Then
            atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - Val(tmpcard.UpperNum)
            If turnatk = 2 And movecp > 1 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(tmpcard.UpperNum)
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(tmpcard.UpperNum)
            End If
            If 攻擊防禦骰子總數(4) = atkcom(角色人物對戰人數(2, 2)) Then
                攻擊防禦骰子總數(4) = 0
            End If
        End If
        If tmpcard.UpperType = a2a Then
            atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - Val(tmpcard.UpperNum)
            If turnatk = 1 Then
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(tmpcard.UpperNum)
                攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(tmpcard.UpperNum)
            End If
        End If
        If tmpcard.UpperType = a3a Then
            atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - Val(tmpcard.UpperNum)
        End If
        If tmpcard.UpperType = a4a Then
            atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - Val(tmpcard.UpperNum)
        End If
        '================
        目前數(9) = Utils.IndexOf(戰鬥系統類.CardDeckCollection(8), tmpcard)
        pagecomleadmax(0) = Val(pagecomleadmax(0)) + 1
        pageqlead(2) = Val(pageqlead(2)) - 1
        FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) + 1
        FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
        tmpcard.ComMark = 0
        '=============以下是牌移動(回牌)(電腦)
        戰鬥系統類.座標計算_電腦手牌
        牌移動暫時變數(3) = Index
        tmpcard.XYLeft = FormMainMode.card(Index).Left  '指定目前Left(座標)
        tmpcard.XYTop = FormMainMode.card(Index).Top  '指定目前Top(座標)
        戰鬥系統類.計算牌移動距離單位
        戰鬥系統類.公用牌變背面
        戰鬥系統類.卡牌牌堆集合更換 tmpcard, 8, 7
        目前數(15) = 0
        FormMainMode.牌移動.Enabled = True
        一般系統類.音效播放 1
        '================以下是出牌對齊
        目前數(7) = 0
        FormMainMode.電腦出牌_出牌對齊_靠右.Enabled = True
        '=====================以下是技能檢查及啟動
        If 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards") <> 0 Then
            vbecommadnum(2, 執行階段系統_搜尋正在執行之執行階段("AtkingSeizeEnemyCards")) = 2    '(階段2)
        End If
        '==============================================
        Select Case turnatk
            Case 1
                '===========================執行階段插入點(ATK-42/DEF-43)
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 42
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 43
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
                '============================
            Case 2
                '===========================執行階段插入點(ATK-42/DEF-43)
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 42
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 43
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
                '============================
            Case 3
                '===========================執行階段插入點(44)
                ReDim VBEStageNum(0 To 1) As Integer
                VBEStageNum(0) = 44
                VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
                執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
                '============================
        End Select
        戰鬥系統類.骰量更新顯示
    End If
End Sub
Sub 電腦牌_模擬轉牌_外(ByVal Index As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = 戰鬥系統類.CardDeckCollection(8)(CStr(Index))

    Call tmpcard.Reverse
    一般系統類.音效播放 3

    FormMainMode.card(Index).CardRotationType = tmpcard.CardOnIn

    If tmpcard.UpperType = a1a Then
        atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + tmpcard.UpperNum
        If turnatk = 2 And movecp = 1 And 攻擊防禦骰子總數(4) = 0 Then
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
        End If
        If turnatk = 2 And movecp = 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(tmpcard.UpperNum)
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(tmpcard.UpperNum)
        End If
    End If
    If tmpcard.UpperType = a5a Then
        atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + tmpcard.UpperNum
        If turnatk = 2 And movecp > 1 And 攻擊防禦骰子總數(4) = 0 Then
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
        End If
        If turnatk = 2 And movecp > 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(tmpcard.UpperNum)
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(tmpcard.UpperNum)
        End If
    End If
    If tmpcard.UpperType = a2a Then
        atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + tmpcard.UpperNum
        If turnatk = 1 And 攻擊防禦骰子總數(4) = 0 Then
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + defcom(角色人物對戰人數(2, 2))
        End If
        If turnatk = 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(tmpcard.UpperNum)
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(tmpcard.UpperNum)
        End If
    End If
    If tmpcard.UpperType = a3a Then
        atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + tmpcard.UpperNum
    End If
    If tmpcard.UpperType = a4a Then
        atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + tmpcard.UpperNum
    End If
    '======================================
    If tmpcard.LowerType = a1a Then
        atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - tmpcard.LowerNum
        If turnatk = 2 And movecp = 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(tmpcard.LowerNum)
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(tmpcard.LowerNum)
        End If
        If 攻擊防禦骰子總數(4) = atkcom(角色人物對戰人數(2, 2)) Then
            攻擊防禦骰子總數(4) = 0
        End If
    End If
    If tmpcard.LowerType = a5a Then
        atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - tmpcard.LowerNum
        If turnatk = 2 And movecp > 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(tmpcard.LowerNum)
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(tmpcard.LowerNum)
        End If
        If 攻擊防禦骰子總數(4) = atkcom(角色人物對戰人數(2, 2)) Then
            攻擊防禦骰子總數(4) = 0
        End If
    End If
    If tmpcard.LowerType = a2a Then
        atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - tmpcard.LowerNum
        If turnatk = 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(tmpcard.LowerNum)
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(tmpcard.LowerNum)
        End If
    End If
    If tmpcard.LowerType = a3a Then
        atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - tmpcard.LowerNum
    End If
    If tmpcard.LowerType = a4a Then
        atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - tmpcard.LowerNum
    End If
    '==============================================
    Select Case turnatk
        Case 1
            '===========================執行階段插入點(ATK-42/DEF-43)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 42
            VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
            執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 42, 4
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 43
            VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
            執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 43, 4
            '============================
        Case 2
            '===========================執行階段插入點(ATK-42/DEF-43)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 42
            VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
            執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 42, 4
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 43
            VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
            執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 43, 4
            '============================
        Case 3
            '===========================執行階段插入點(44)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 44
            VBEStageNum(1) = -2    '觸發方(1.使用者/2.電腦)
            執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 44, 3
            '============================
    End Select
    戰鬥系統類.骰量更新顯示
End Sub
Sub 骰數零執行判斷()
    Dim ustruenum As Integer, comtruenum As Integer
    Dim p As Integer, i As Integer, j As Integer
    '==無介面表示，需自行擲骰
    For p = 1 To 擲骰表單溝通暫時變數(9)
        Randomize Timer
        i = Int(Rnd() * 6) + 1
        If i = 1 Or i = 6 Then ustruenum = ustruenum + 1
    Next
    For p = 1 To 擲骰表單溝通暫時變數(10)
        Randomize Timer
        j = Int(Rnd() * 6) + 1
        If j = 1 Or j = 6 Then comtruenum = comtruenum + 1
    Next
    If 是否系統公骰 = True Then
        擲骰表單溝通暫時變數(5) = ustruenum
        擲骰表單溝通暫時變數(6) = comtruenum
    Else
        Vss_BattleStartDiceNum(3) = ustruenum
        Vss_BattleStartDiceNum(4) = comtruenum
    End If
End Sub
Sub 擲骰表單顯示()
    If 骰數零檢查值(1) = False And 骰數零檢查值(2) = False Then
        If moveturn = 1 Then
            Select Case 擲骰表單溝通暫時變數(1)
                Case 1
                    FormMainMode.PEAFDiceInterface.DiceATKMode = 1
                    FormMainMode.PEAFDiceInterface.DiceInputMode = 2
                    FormMainMode.PEAFDiceInterface.diceusTotal = 擲骰表單溝通暫時變數(9)
                    FormMainMode.PEAFDiceInterface.dicecomTotal = 擲骰表單溝通暫時變數(10)
                    FormMainMode.PEAFDiceInterface.PersonImage = 戰鬥擲骰介面人物立繪圖路徑紀錄數(1)
                    FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 2, 5))
                    FormMainMode.PEAFDiceInterface.ZOrder
                    戰鬥系統類.擲骰時血量介面頂層顯示
                    FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
                    FormMainMode.PEAFDiceInterface.DiceStart = True
                Case 2
                    FormMainMode.PEAFDiceInterface.DiceATKMode = 2
                    FormMainMode.PEAFDiceInterface.DiceInputMode = 2
                    FormMainMode.PEAFDiceInterface.diceusTotal = 擲骰表單溝通暫時變數(9)
                    FormMainMode.PEAFDiceInterface.dicecomTotal = 擲骰表單溝通暫時變數(10)
                    FormMainMode.PEAFDiceInterface.PersonImage = 戰鬥擲骰介面人物立繪圖路徑紀錄數(2)
                    FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, 角色人物對戰人數(2, 2), 2, 2, 5))
                    FormMainMode.PEAFDiceInterface.ZOrder
                    戰鬥系統類.擲骰時血量介面頂層顯示
                    FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
                    FormMainMode.PEAFDiceInterface.DiceStart = True
            End Select
        ElseIf moveturn = 2 Then
            Select Case 擲骰表單溝通暫時變數(1)
                Case 1
                    FormMainMode.PEAFDiceInterface.DiceATKMode = 2
                    FormMainMode.PEAFDiceInterface.DiceInputMode = 2
                    FormMainMode.PEAFDiceInterface.diceusTotal = 擲骰表單溝通暫時變數(9)
                    FormMainMode.PEAFDiceInterface.dicecomTotal = 擲骰表單溝通暫時變數(10)
                    FormMainMode.PEAFDiceInterface.PersonImage = 戰鬥擲骰介面人物立繪圖路徑紀錄數(2)
                    FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, 角色人物對戰人數(2, 2), 2, 2, 5))
                    FormMainMode.PEAFDiceInterface.ZOrder
                    戰鬥系統類.擲骰時血量介面頂層顯示
                    FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
                    FormMainMode.PEAFDiceInterface.DiceStart = True
                Case 2
                    FormMainMode.PEAFDiceInterface.DiceATKMode = 1
                    FormMainMode.PEAFDiceInterface.DiceInputMode = 2
                    FormMainMode.PEAFDiceInterface.diceusTotal = 擲骰表單溝通暫時變數(9)
                    FormMainMode.PEAFDiceInterface.dicecomTotal = 擲骰表單溝通暫時變數(10)
                    FormMainMode.PEAFDiceInterface.PersonImage = 戰鬥擲骰介面人物立繪圖路徑紀錄數(1)
                    FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 2, 5))
                    FormMainMode.PEAFDiceInterface.ZOrder
                    戰鬥系統類.擲骰時血量介面頂層顯示
                    FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
                    FormMainMode.PEAFDiceInterface.DiceStart = True
            End Select
        End If
    Else
        '========================
        目前數(26) = 0
        '========================
        戰鬥系統類.骰數零執行判斷
    End If
End Sub
Sub 擲骰時血量介面頂層顯示()
    FormMainMode.PEAFbloodbackimage1.ZOrder
    FormMainMode.PEAFbloodbackimage2.ZOrder
    FormMainMode.bloodnumus1.ZOrder
    FormMainMode.bloodnumus2.ZOrder
    FormMainMode.bloodnumcom1.ZOrder
    FormMainMode.bloodnumcom2.ZOrder
End Sub
Sub 擲骰後續判斷()
    If 是否系統公骰 = True Then
        If 骰數零檢查值(1) = False And 骰數零檢查值(2) = False Then
            擲骰表單溝通暫時變數(5) = Val(FormMainMode.PEAFDiceInterface.diceusTrue)
            擲骰表單溝通暫時變數(6) = Val(FormMainMode.PEAFDiceInterface.dicecomTrue)
        End If
        FormMainMode.骰子執行完啟動.Enabled = True
    Else
        If 骰數零檢查值(1) = False And 骰數零檢查值(2) = False Then
            Vss_BattleStartDiceNum(3) = Val(FormMainMode.PEAFDiceInterface.diceusTrue)
            Vss_BattleStartDiceNum(4) = Val(FormMainMode.PEAFDiceInterface.dicecomTrue)
        End If
    End If
    '=====================================================
    If Val(擲骰表單溝通暫時變數(4)) = 1 Then
        Select Case Val(擲骰表單溝通暫時變數(1))
            Case 1
                GoTo usatkcom
            Case 2
                GoTo comatkus
        End Select
    Else
        Select Case Val(擲骰表單溝通暫時變數(1))
            Case 1
                GoTo comatkus
            Case 2
                GoTo usatkcom
        End Select
    End If
    '==========================================
    Exit Sub
usatkcom:
    If 是否系統公骰 = True Then
        擲骰表單溝通暫時變數(2) = 擲骰表單溝通暫時變數(5) - 擲骰表單溝通暫時變數(6)
        擲骰表單溝通暫時變數(3) = 2
    Else
        Vss_BattleStartDiceNum(5) = Vss_BattleStartDiceNum(3) - Vss_BattleStartDiceNum(4)
    End If
    '==========================================
    Exit Sub
comatkus:
    If 是否系統公骰 = True Then
        擲骰表單溝通暫時變數(2) = 擲骰表單溝通暫時變數(6) - 擲骰表單溝通暫時變數(5)
        擲骰表單溝通暫時變數(3) = 1
    Else
        Vss_BattleStartDiceNum(5) = Vss_BattleStartDiceNum(4) - Vss_BattleStartDiceNum(3)
    End If
End Sub
Sub 雙方HP檢查()
    Dim inp As Integer    'RND暫時變數
    Dim person(1 To 2) As Integer
    Erase 人物消失檢查暫時變數
    If livecom(角色人物對戰人數(2, 2)) <= 0 Then
        人物消失檢查暫時變數(3) = 1
        If livecom(角色待機人物紀錄數(2, 2)) > 0 Then
            person(2) = 2
            交換角色紀錄暫時變數(2) = 1
        ElseIf livecom(角色待機人物紀錄數(2, 3)) > 0 Then
            交換角色紀錄暫時變數(2) = 1
            person(2) = 2
        Else
            person(2) = 1
        End If
    End If
    If liveus(角色人物對戰人數(1, 2)) <= 0 Then
        人物消失檢查暫時變數(2) = 1
        If liveus(角色待機人物紀錄數(1, 2)) > 0 Or liveus(角色待機人物紀錄數(1, 3)) > 0 Then
            person(1) = 2
            交換角色紀錄暫時變數(1) = 1
        Else
            person(1) = 1
        End If
    End If

    If person(1) = 2 Or person(2) = 2 Then
        等待時間佇列(1).Add 21
        FormMainMode.人物消失檢查.Enabled = True
        Exit Sub
    ElseIf person(1) = 0 And person(2) = 1 Then
        戰鬥模式勝敗紀錄數 = 1
        等待時間佇列(1).Add 36
        FormMainMode.人物消失檢查.Enabled = True
    ElseIf person(1) = 1 And person(2) = 0 Then
        等待時間佇列(1).Add 36
        戰鬥模式勝敗紀錄數 = 2
        FormMainMode.人物消失檢查.Enabled = True
    ElseIf person(1) = 1 And person(2) = 1 Then
        Randomize
        inp = Int(Rnd() * 2) + 1
        Select Case inp
            Case 1
                戰鬥模式勝敗紀錄數 = 1
                等待時間佇列(1).Add 36
                FormMainMode.人物消失檢查.Enabled = True
            Case 2
                戰鬥模式勝敗紀錄數 = 2
                等待時間佇列(1).Add 36
                FormMainMode.人物消失檢查.Enabled = True
        End Select
    End If

    If FormMainMode.人物消失檢查.Enabled = False Then
        Select Case HP檢查階段數
            Case 1
                '----------以下為階段繼續實行（移動階段3）
                等待時間佇列(1).Add 4
                FormMainMode.等待時間.Enabled = True
            Case 2
                等待時間佇列(1).Add 11
                FormMainMode.等待時間.Enabled = True
            Case 3
                戰鬥系統類.階段執行判斷
            Case 4
                FormMainMode.NextTurn_階段2.Enabled = True
        End Select
    End If
End Sub
Function 雙方HP檢查_結束回合檢查() As Boolean
    Dim num(1 To 2) As Integer    '選擇人物暫時變數
    Dim i As Integer

    If BattleTurn >= Val(Formsetting.ckendturnnum.Text) And Formsetting.ckendturn.Value = 1 Then
        雙方HP檢查_結束回合檢查 = True
        '==============
        For i = 1 To 3
            If liveus(角色待機人物紀錄數(1, i)) > 0 Then
                num(1) = Val(num(1)) + Val(liveus(角色待機人物紀錄數(1, i)))
            End If
            If livecom(角色待機人物紀錄數(2, i)) > 0 Then
                num(2) = Val(num(2)) + Val(livecom(角色待機人物紀錄數(2, i)))
            End If
        Next
        '==============
        If num(1) > num(2) Then
            戰鬥模式勝敗紀錄數 = 1
            FormMainMode.trend.Enabled = True
        ElseIf num(1) < num(2) Then
            戰鬥模式勝敗紀錄數 = 2
            FormMainMode.trend.Enabled = True
        ElseIf num(1) = num(2) Then
            '無條件敗北
            戰鬥模式勝敗紀錄數 = 2
            FormMainMode.trend.Enabled = True
        End If
    Else
        雙方HP檢查_結束回合檢查 = False
    End If
End Function

Sub checkpage()
    Dim i As Integer

    For i = 1 To 目前數(11)
        If 目前數(10) = 1 Then
            FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
            pageqlead(1) = Val(pageqlead(1)) - 1
        ElseIf 目前數(10) = 2 Then
            FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
            pageqlead(2) = Val(pageqlead(2)) - 1
        End If
    Next
End Sub
Sub chkcom()
    If goicheck(2) = 0 Then
        If atkingpagetot(2, 1) > 0 And movecp = 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkcom(角色人物對戰人數(2, 2))
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
            goicheck(2) = 1
        ElseIf atkingpagetot(2, 5) > 0 And movecp > 1 Then
            攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkcom(角色人物對戰人數(2, 2))
            攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
            goicheck(2) = 1
        End If
    End If
End Sub
Sub chkdef()
    If goidefus = 0 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + defus(角色人物對戰人數(1, 2))
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + defus(角色人物對戰人數(1, 2))
        FormMainMode.顯示列1.goi1 = Val(FormMainMode.顯示列1.goi1) + defus(角色人物對戰人數(1, 2))
        goidefus = 1
    End If
End Sub
Sub chkdefcom()
    If chkcomck = 0 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + defcom(角色人物對戰人數(2, 2))
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + defcom(角色人物對戰人數(2, 2))
        FormMainMode.顯示列1.goi2 = Val(FormMainMode.顯示列1.goi2) + defcom(角色人物對戰人數(2, 2))
        chkcomck = 1
    End If
End Sub
Sub chkus1()
    If goicheck(1) = 0 Then
        If atkingpagetot(1, 1) > 0 Then
            攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkus(角色人物對戰人數(1, 2))
            攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
            goicheck(1) = 1
        End If
    End If
End Sub
Sub cleanatkingpagetot()
    Dim i As Integer, j As Integer

    For i = 1 To 2
        For j = 1 To 5
            atkingpagetot(i, j) = 0
        Next
    Next
End Sub
Sub comatk1()
    Dim a As Integer
    Dim tmpcard As clsActionCard

    For a = 1 To 戰鬥系統類.CardDeckCollection(7).Count
        Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(a)
        If tmpcard.ComMark <> 1 Then
            If tmpcard.UpperType = a1a Then
                tmpcard.ComMark = 1
            ElseIf tmpcard.LowerType = a1a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
            End If
        End If
    Next
End Sub
Sub comatk2()
    Dim j As Integer
    Dim tmpcard As clsActionCard

    For j = 1 To 戰鬥系統類.CardDeckCollection(7).Count
        Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(j)
        If tmpcard.ComMark <> 1 Then
            If tmpcard.UpperType = a5a Then
                tmpcard.ComMark = 1
            ElseIf tmpcard.LowerType = a5a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
            End If
        End If
    Next
End Sub
Sub comatk_智慧型AI引導程序_超出牌張數(ByVal turn As Integer, ByVal movecpre As Integer, ByVal choose As Integer)
    Dim werstr As String, werbo As Boolean
    Dim a As Integer, k As Integer
    Dim tmpcard As clsActionCard

    If movecpre = 1 And turn = 1 Then
        werstr = a1a
    ElseIf movecpre > 1 And turn = 1 Then
        werstr = a5a
    ElseIf turn = 2 Then
        werstr = a2a
    End If
    '=================================
    For a = 1 To 戰鬥系統類.CardDeckCollection(7).Count
        werbo = False
        Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(a)
        For k = 1 To UBound(cardAInumOvertenrecord)
            If tmpcard.CardNum = cardAInumOvertenrecord(k) Then
                werbo = True
            End If
        Next
        If tmpcard.ComMark <> 1 And werbo = False Then
            If tmpcard.UpperType = werstr Then
                tmpcard.ComMark = 1
            ElseIf tmpcard.LowerType = werstr Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
            End If
            If choose = 1 And tmpcard.ComMark = 0 Then
                tmpcard.ComMark = 1
            End If
        End If
    Next
End Sub
Sub moveatkin()
    Dim j As Integer
    Dim tmpcard As clsActionCard

    Do
        For j = 1 To 戰鬥系統類.CardDeckCollection(7).Count
            Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(j)
            If tmpcard.CardType = 2 And tmpcard.ComMark <> 1 Then
                If tmpcard.UpperType = a3a And tmpcard.LowerType = a3a Then    '移動單面事件卡優先
                    tmpcard.ComMark = 1
                    目前數(25) = 目前數(25) + tmpcard.UpperNum
                End If
                If 目前數(25) >= 2 Then Exit Do
            End If
        Next
        For j = 1 To 戰鬥系統類.CardDeckCollection(7).Count
            Set tmpcard = 戰鬥系統類.CardDeckCollection(7)(j)
            If tmpcard.ComMark <> 1 Then
                If tmpcard.UpperType = a3a Then
                    tmpcard.ComMark = 1
                    目前數(25) = 目前數(25) + 1
                ElseIf tmpcard.LowerType = a3a Then
                    Call tmpcard.Reverse
                    tmpcard.ComMark = 1
                    目前數(25) = 目前數(25) + tmpcard.UpperNum
                End If
                If 目前數(25) >= 2 Then Exit Do
            End If
        Next
        Exit Do
    Loop
End Sub
Sub movetnus()
    戰鬥系統類.廣播訊息 "你有主動權。"
    FormMainMode.move3.Picture = LoadPicture(app_path & "gif\system\atk1.gif")
    FormMainMode.move4.Picture = LoadPicture(app_path & "gif\system\def1.gif")
    FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\system\atk2.gif")
    FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\system\def2.gif")
    moveturn = 1
    FormMainMode.cnmove2.Visible = False
    擲骰表單溝通暫時變數(1) = 1
End Sub
Sub movetncom()
    戰鬥系統類.廣播訊息 "對方有主動權。"
    FormMainMode.move3.Picture = LoadPicture(app_path & "gif\system\def1.gif")
    FormMainMode.move4.Picture = LoadPicture(app_path & "gif\system\atk1.gif")
    FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\system\def2.gif")
    FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\system\atk2.gif")
    moveturn = 2
    FormMainMode.cnmove2.Visible = False
    擲骰表單溝通暫時變數(1) = 1
End Sub
Sub 人物交換_使用者_指定交換(ByVal num As Integer)
    Dim ae As Integer, n As Integer, i As Integer, ahmt As String
    Dim tmpobj As clsPersonActiveSkill
    '=======================
    ReDim VBEStageNum(0 To 3) As Integer
    VBEStageNum(0) = 41
    VBEStageNum(1) = -1    '執行效果方(1.使用者/2.電腦)
    VBEStageNum(2) = 1    '交換前人物編號
    VBEStageNum(3) = num    '交換後人物編號
    '===========================執行階段插入點(41)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 41, 1
    '============================
    FormMainMode.personusminijpg.小人物消失 = True
    Do Until FormMainMode.personusminijpg.小人物消失 = False
        DoEvents
    Loop
    '=======================
    ae = 角色人物對戰人數(1, 2)
    角色人物對戰人數(1, 2) = 角色待機人物紀錄數(1, num)
    角色待機人物紀錄數(1, 1) = 角色人物對戰人數(1, 2)
    角色待機人物紀錄數(1, num) = ae
    FormMainMode.PEAFpersoncardus(角色待機人物紀錄數(1, num)).Left = 2520 * (num - 1)
    FormMainMode.PEAFpersoncardus(角色待機人物紀錄數(1, num)).Visible = True
    FormMainMode.cardus(角色待機人物紀錄數(1, num)).Visible = False

    FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).Left = 0
    FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).Visible = False
    FormMainMode.cardus(角色人物對戰人數(1, 2)).Left = 0
    FormMainMode.cardus(角色人物對戰人數(1, 2)).Top = 6240
    FormMainMode.cardus(角色人物對戰人數(1, 2)).ZOrder
    FormMainMode.cardus(角色人物對戰人數(1, 2)).Visible = True
    '=======================
    戰鬥系統類.技能說明載入_使用者
    FormMainMode.PEAFInterface.Passive_技能一方全重設 1
    For n = 5 To 8
        If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 1) = "" Then
            FormMainMode.PEAFInterface.Passive_使用者_技能隱藏 n - 4
        Else
            FormMainMode.PEAFInterface.Passive_使用者_技能名稱 VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 1), n - 4
            FormMainMode.PEAFInterface.Passive_使用者_技能顯示 n - 4
            '=======================
            If atkingck(1, 角色人物對戰人數(1, 2), n, 1) = 1 Then
                FormMainMode.PEAFInterface.Passive_使用者_技能燈發亮 n - 4
            End If
        End If
    Next
    If 人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 1) <> "" And Val(人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 2)) = 1 Then
        FormMainMode.personusminijpg.小人物圖片 = 人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 4)
        FormMainMode.personusminijpg.小人物影子圖片 = 人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 5)
        FormMainMode.顯示列1.使用者方小人物圖片 = 人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 6)
        FormMainMode.personusminijpg.小人物影子Left = Val(人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 7))
        FormMainMode.personusminijpg.小人物影子top差 = Val(人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 8))
        戰鬥擲骰介面人物立繪圖路徑紀錄數(1) = 人物實際狀態資料庫(1, 角色人物對戰人數(1, 2), 3)
    Else
        FormMainMode.personusminijpg.小人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 1)
        FormMainMode.personusminijpg.小人物影子圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 2)
        FormMainMode.顯示列1.使用者方小人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 4)
        FormMainMode.personusminijpg.小人物影子Left = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 1, 5))
        FormMainMode.personusminijpg.小人物影子top差 = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 1, 6))
        戰鬥擲骰介面人物立繪圖路徑紀錄數(1) = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 3)
    End If
    FormMainMode.顯示列1.使用者方小人物圖片left = -(FormMainMode.顯示列1.使用者方小人物圖片width)
    '--------------------------計算新距離單位(HP血條)
    距離單位(1, 1, 1) = 5295 \ liveusmax(角色人物對戰人數(1, 2))
    FormMainMode.bloodlineout1.Width = (距離單位(1, 1, 1) * liveus(角色人物對戰人數(1, 2)))
    FormMainMode.bloodnumus1.Caption = liveus(角色人物對戰人數(1, 2))
    FormMainMode.bloodnumus2.Caption = liveusmax(角色人物對戰人數(1, 2))
    '========================
    執行動作_距離變更 movecp, False, True
    '========================
    For i = 1 To 4
        戰鬥系統類.人物技能欄燈開關 False, i
    Next
    '=============================
    FormMainMode.personusminijpg.小人物顯現 = True
    Do Until FormMainMode.personusminijpg.小人物顯現 = False
        DoEvents
    Loop

End Sub

Sub 人物交換_電腦_指定交換(ByVal num As Integer)
    Dim ae As Integer, n As Integer
    '=======================
    ReDim VBEStageNum(0 To 3) As Integer
    VBEStageNum(0) = 41
    VBEStageNum(1) = -2    '執行效果方(1.使用者/2.電腦)
    VBEStageNum(2) = 1    '交換前人物編號
    VBEStageNum(3) = num    '交換後人物編號
    '===========================執行階段插入點(41)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 41, 1
    '============================
    FormMainMode.personcomminijpg.小人物消失 = True
    Do Until FormMainMode.personcomminijpg.小人物消失 = False
        DoEvents
    Loop
    '=======================
    ae = 角色人物對戰人數(2, 2)
    角色人物對戰人數(2, 2) = 角色待機人物紀錄數(2, num)
    角色待機人物紀錄數(2, num) = ae
    角色待機人物紀錄數(2, 1) = 角色人物對戰人數(2, 2)
    FormMainMode.PEAFpersoncardcom(角色待機人物紀錄數(2, num)).Left = 2520 * (num - 1)
    FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).Left = 0
    '=======================
    戰鬥系統類.技能說明載入_電腦
    FormMainMode.PEAFInterface.Passive_技能一方全重設 2
    For n = 5 To 8
        If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 1) = "" Then
            FormMainMode.PEAFInterface.Passive_電腦_技能隱藏 n - 4
        Else
            FormMainMode.PEAFInterface.Passive_電腦_技能名稱 VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 1), n - 4
            FormMainMode.PEAFInterface.Passive_電腦_技能顯示 n - 4
            '=======================
            If atkingck(2, 角色人物對戰人數(2, 2), n, 1) = 1 Then
                FormMainMode.PEAFInterface.Passive_電腦_技能燈發亮 n - 4
            End If
        End If
    Next
    '====================
    If 人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 1) <> "" And Val(人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 2)) = 1 Then
        FormMainMode.personcomminijpg.小人物圖片 = 人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 4)
        FormMainMode.personcomminijpg.小人物影子圖片 = 人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 5)
        FormMainMode.顯示列1.電腦方小人物圖片 = 人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 6)
        FormMainMode.personcomminijpg.小人物影子Left = Val(人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 7))
        FormMainMode.personcomminijpg.小人物影子top差 = Val(人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 8))
        戰鬥擲骰介面人物立繪圖路徑紀錄數(2) = 人物實際狀態資料庫(2, 角色人物對戰人數(2, 2), 3)
    Else
        FormMainMode.personcomminijpg.小人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 1)
        FormMainMode.personcomminijpg.小人物影子圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 2)
        FormMainMode.顯示列1.電腦方小人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 4)
        FormMainMode.personcomminijpg.小人物影子Left = VBEPerson(2, 角色人物對戰人數(2, 2), 2, 1, 5)
        FormMainMode.personcomminijpg.小人物影子top差 = VBEPerson(2, 角色人物對戰人數(2, 2), 2, 1, 6)
        戰鬥擲骰介面人物立繪圖路徑紀錄數(2) = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 3)
    End If
    FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
    戰鬥系統類.介面角色小卡資訊寫入 2, 角色人物對戰人數(2, 2)
    戰鬥系統類.PersonCardShowOnMode(2, 角色人物對戰人數(2, 2)) = True
    FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).ShowOnMode = True
    FormMainMode.cardcom(角色人物對戰人數(2, 2)).ShowOnMode = True
    '--------------------------計算新距離單位(HP血條)
    距離單位(1, 2, 1) = (11340 - 6060) \ livecommax(角色人物對戰人數(2, 2))
    FormMainMode.bloodlineout2.Left = 11340 - (距離單位(1, 2, 1) * livecom(角色人物對戰人數(2, 2)))
    FormMainMode.bloodnumcom1.Caption = livecom(角色人物對戰人數(2, 2))
    FormMainMode.bloodnumcom2.Caption = livecommax(角色人物對戰人數(2, 2))
    '==============================
    執行動作_距離變更 movecp, False, True
    '=============================
    FormMainMode.personcomminijpg.小人物顯現 = True
    Do Until FormMainMode.personcomminijpg.小人物顯現 = False
        DoEvents
    Loop
    '=======================
End Sub
Sub 執行動作_交換人物角色_使用者_初始()
    Dim i As Integer, k As Integer
    Dim ne As Integer
    Dim numNow As Integer, obj As clsStatus

    For i = 2 To 3
        Formchangeperson.card(i - 1).異常狀態全重設
        Formchangeperson.card(i - 1).CardBack全重設
        Formchangeperson.card(i - 1).CardMain_角色圖片 = VBEPerson(1, 角色待機人物紀錄數(1, i), 1, 5, 5)
        Formchangeperson.card(i - 1).CardMain_角色HP = liveus(角色待機人物紀錄數(1, i))
        Formchangeperson.card(i - 1).CardMain_角色HPMAX = liveusmax(角色待機人物紀錄數(1, i))
        Formchangeperson.card(i - 1).CardMain_角色ATK = atkus(角色待機人物紀錄數(1, i))
        Formchangeperson.card(i - 1).CardMain_角色DEF = defus(角色待機人物紀錄數(1, i))
        Formchangeperson.card(i - 1).CardMain_是否為新樣式資訊 = CBool(Val(VBEPerson(1, 角色待機人物紀錄數(1, i), 1, 3, 5)) = 1)
    Next
    戰鬥系統類.技能說明載入_人物卡片背面_交換角色

    ne = 1
    For k = 2 To 3
        numNow = 1
        For Each obj In 人物異常狀態列表(1, 角色待機人物紀錄數(1, k))
            Formchangeperson.card(ne).更改異常狀態資料 numNow, obj.ImagePath, obj.Value, obj.Total, True
            numNow = numNow + 1
            If numNow > 14 Then Exit For
        Next
        ne = ne + 1
    Next

    交換角色紀錄暫時變數(1) = 0
    For i = 2 To 3
        Formchangeperson.card(i - 1).MusicPlayerObj = FormMainMode.cMusicPlayer(9)
        Formchangeperson.card(i - 1).ShowOnMode = True
    Next
    If Formsetting.chkusenewaipersonauto.Value = 1 Then
        Formchangeperson.使用者方智慧型AI_自動控制選人.Enabled = True
    End If
    Formchangeperson.Left = FormMainMode.Left + 2430
    Formchangeperson.Top = FormMainMode.Top + 1655
    Formchangeperson.Show 1
End Sub
Sub 執行動作_交換人物角色_電腦_初始()
    Select Case 交換角色紀錄暫時變數(2)
        Case 1
            交換角色紀錄暫時變數(2) = 0
            等待時間佇列(1).Add 18
            FormMainMode.等待時間.Enabled = True
        Case 0
            等待時間佇列(1).Add 19
            FormMainMode.等待時間.Enabled = True
    End Select

End Sub
Sub 執行動作_交換人物角色_電腦_交換()
    If livecom(角色待機人物紀錄數(2, 2)) > 0 Then
        人物交換_電腦_指定交換 2
    ElseIf livecom(角色待機人物紀錄數(2, 3)) > 0 Then
        人物交換_電腦_指定交換 3
    End If
    執行動作_交換人物角色_結束執行
End Sub
Sub 執行動作_交換人物角色_初始()
    If (交換角色紀錄暫時變數(1) = 1 Or 交換角色紀錄暫時變數(2) = 1) And 交換角色紀錄暫時變數(3) = 0 Then
        turnatk = 6
        階段狀態數 = 5
        戰鬥系統類.時間軸_重設
        FormMainMode.顯示列1.顯示列圖片 = App.Path & "\gif\system\linechange.png"
        FormMainMode.顯示列1.Visible = True
        FormMainMode.顯示列1.goi1顯示 = False
        FormMainMode.顯示列1.goi2顯示 = False
        戰鬥系統類.時間軸_顯示
        FormMainMode.trtimeline.Enabled = True
        小人物頭像移動方向數(1) = 2
        小人物頭像移動方向數(2) = 2
        FormMainMode.小人物頭像移動_使用者.Enabled = True
        FormMainMode.小人物頭像移動_電腦.Enabled = True
        交換角色紀錄暫時變數(3) = 1
        FormMainMode.顯示列1.移動階段選擇值 = 0
        FormMainMode.顯示列1.移動階段圖顯示 = False
    End If
    If 交換角色紀錄暫時變數(1) = 1 Then
        執行動作_交換人物角色_使用者_初始
    ElseIf 交換角色紀錄暫時變數(2) = 1 Then
        執行動作_交換人物角色_電腦_初始
    End If
End Sub
Sub 執行動作_移動階段選擇執行()
    '===========交換角色類
    If 交換角色紀錄暫時變數(1) = 1 Or 交換角色紀錄暫時變數(2) = 1 Then
        執行動作_交換人物角色_初始
    Else
        交換角色紀錄暫時變數(3) = 0
        等待時間佇列(1).Add 17
        FormMainMode.等待時間.Enabled = True
    End If
End Sub
Sub 執行動作_人物死亡交換階段選擇執行()
    If 交換角色紀錄暫時變數(1) = 1 Or 交換角色紀錄暫時變數(2) = 1 Then
        執行動作_交換人物角色_初始
    Else
        交換角色紀錄暫時變數(3) = 0
        等待時間佇列(1).Add 20
        FormMainMode.等待時間.Enabled = True
    End If
End Sub
Sub 執行動作_交換人物角色_結束執行()
    Formchangeperson.Hide
    戰鬥系統類.時間軸_停止
    Select Case 交換角色紀錄暫時變數(4)
        Case 1
            執行動作_移動階段選擇執行
        Case 2
            執行動作_人物死亡交換階段選擇執行
    End Select
End Sub
Sub 事件卡處理_初始_使用者方()
    Dim ck As Boolean
    Dim m As Integer, i As Integer, j As Integer, tmpfailed As Integer, tmpcardstr As String

    If Formsetting.comboeventcarrdus.Text = "無" Then    '=====(無)
        For i = 1 To 18
            Randomize
            m = Int(Rnd() * 3) + 1
            Select Case m
                Case 1
                    tmpcardstr = "劍1"
                Case 2
                    tmpcardstr = "槍1"
                Case 3
                    tmpcardstr = "防1"
            End Select
            戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
        Next
    ElseIf Formsetting.comboeventcarrdus.Text = "自訂" Then    '=====自訂
        If 事件卡記錄暫時數(0, 1) = 18 Or Formsetting.persontgreus.Value = 0 Then
            For i = 1 To 18
                If 一般系統類.事件卡資料庫(Formsetting.personus(i).Text, 1) = 99 Then
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                Else
                    tmpcardstr = Formsetting.personus(i).Text
                End If
                戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        ElseIf 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
            For i = 1 To 18
                If i >= 7 Or 一般系統類.事件卡資料庫(Formsetting.personus(i).Text, 1) = 99 Then
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                Else
                    tmpcardstr = Formsetting.personus(i).Text
                End If
                戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        End If
    ElseIf Formsetting.comboeventcarrdus.Text = "最大值" Then    '===============選擇最大值
        If Formsetting.persontgreus.Value = 1 Then  '===遵守規則
            For i = 1 To 18
                If i = 7 And 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then Exit For

                Select Case Formsetting.persontgus(i).Caption
                    Case 0
                        Randomize
                        m = Int(Rnd() * 8) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "劍3/槍1"
                            Case 2
                                tmpcardstr = "槍3/劍1"
                            Case 3
                                tmpcardstr = "防3/移1"
                            Case 4
                                tmpcardstr = "劍3/移1"
                            Case 5
                                tmpcardstr = "槍3/移1"
                            Case 6
                                tmpcardstr = "劍3/防1"
                            Case 7
                                tmpcardstr = "槍3/防1"
                            Case 8
                                tmpcardstr = "特2"
                        End Select
                    Case 1
                        Randomize
                        m = Int(Rnd() * 3) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "劍5/槍3"
                            Case 2
                                tmpcardstr = "劍5/移1"
                            Case 3
                                tmpcardstr = "劍8"
                        End Select
                    Case 2
                        Randomize
                        m = Int(Rnd() * 3) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "槍5/劍3"
                            Case 2
                                tmpcardstr = "槍5/移1"
                            Case 3
                                tmpcardstr = "槍8"
                        End Select
                    Case 3
                        Randomize
                        m = Int(Rnd() * 3) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "防5/移1"
                            Case 2
                                tmpcardstr = "防7"
                            Case 3
                                tmpcardstr = "HP回復3"
                        End Select
                    Case 4
                        Randomize
                        m = Int(Rnd() * 2) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "移3/特3"
                            Case 2
                                tmpcardstr = "移5"
                        End Select
                    Case 5
                        tmpcardstr = "機會5"
                    Case 6
                        tmpcardstr = "詛咒術5"
                    Case 7
                        Randomize
                        m = Int(Rnd() * 2) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "特3/防3"
                            Case 1
                                tmpcardstr = "特5"
                        End Select
                End Select
                戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
            If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
                For i = 7 To 18
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                    戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
                Next
            End If
        Else  '================================不遵守規則
            For i = 1 To 18
                Do
                    Randomize
                    m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                    '==============================
                    Select Case Formsetting.personus(i).List(m)
                        Case "劍8"
                            Exit Do
                        Case "槍8"
                            Exit Do
                        Case "防7"
                            Exit Do
                        Case "移5"
                            Exit Do
                        Case "HP回復3"
                            Exit Do
                        Case "機會5"
                            Exit Do
                        Case "詛咒術5"
                            Exit Do
                        Case "特5"
                            Exit Do
                        Case "劍5/槍3"
                            Exit Do
                        Case "槍5/劍3"
                            Exit Do
                        Case "防5/移1"
                            Exit Do
                        Case "槍5/移1"
                            Exit Do
                        Case "劍5/移1"
                            Exit Do
                        Case "移3/特3"
                            Exit Do
                        Case "特3/防3"
                            Exit Do
                    End Select
                Loop
                tmpcardstr = Formsetting.personus(i).List(m)
                戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        End If
    ElseIf Formsetting.comboeventcarrdus.Text = "隨機" Or Formsetting.comboeventcarrdus.Text = "隨機(不含聖水)" Then    '=====隨機
        If Formsetting.persontgreus.Value = 1 Then    '===遵守規則
            For i = 1 To 18
                If i = 7 And 事件卡記錄暫時數(0, 1) = 12 Then Exit For

                tmpfailed = 0
                Do
                    Randomize
                    m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                    If 一般系統類.事件卡資料庫(Formsetting.personus(i).List(m), 1) = Formsetting.persontgus(i).Caption Or _
                       (tmpfailed > 10 And 一般系統類.事件卡資料庫(Formsetting.personus(i).List(m), 1) = 0) Then
                        If Formsetting.comboeventcarrdus.Text = "隨機(不含聖水)" And Formsetting.personus(i).List(m) = "聖水" Then
                        Else
                            tmpcardstr = Formsetting.personus(i).List(m)
                            Exit Do
                        End If
                    End If
                    tmpfailed = tmpfailed + 1
                Loop
                戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
            If 事件卡記錄暫時數(0, 1) = 12 Then
                For i = 7 To 18
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                    戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
                Next
            End If
        Else    '=============================不遵守規則
            For i = 1 To 18
                Randomize
                m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                If Formsetting.comboeventcarrdus.Text = "隨機(不含聖水)" And Formsetting.personus(i).List(m) = "聖水" Then
                    i = i - 1
                Else
                    tmpcardstr = Formsetting.personus(i).List(m)
                End If
                戰鬥系統類.發行卡牌_事件卡 1, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        End If
    End If
End Sub
Sub 事件卡處理_初始_電腦方()
    Dim m As Integer, i As Integer, j As Integer, tmpfailed As Integer, tmpcardstr As String
    Dim ay() As String

    If Formsetting.comboeventcarrdcom.Text = "無" Then    '=====(無)
        For i = 1 To 18
            Randomize
            m = Int(Rnd() * 3) + 1
            Select Case m
                Case 1
                    tmpcardstr = "劍1"
                Case 2
                    tmpcardstr = "槍1"
                Case 3
                    tmpcardstr = "防1"
            End Select
            戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
        Next
    ElseIf Formsetting.comboeventcarrdcom.Text = "自訂" Then    '=====自訂
        If 事件卡記錄暫時數(0, 1) = 18 Or Formsetting.persontgrecom.Value = 0 Then
            For i = 1 To 18
                If 一般系統類.事件卡資料庫(Formsetting.personcom(i).Text, 1) = 99 Then
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                Else
                    tmpcardstr = Formsetting.personcom(i).Text
                End If
                戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        ElseIf 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 1 To 18
                If i >= 7 Or 一般系統類.事件卡資料庫(Formsetting.personcom(i).Text, 1) = 99 Then
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                Else
                    tmpcardstr = Formsetting.personcom(i).Text
                End If
                戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        End If
    ElseIf Formsetting.comboeventcarrdcom.Text = "最大值" Then    '=====選擇最大值
        If Formsetting.persontgrecom.Value = 1 Then  '===遵守規則
            For i = 1 To 18
                If i = 7 And 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then Exit For

                Select Case Formsetting.persontgcom(i).Caption
                    Case 0
                        Randomize
                        m = Int(Rnd() * 8) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "劍3/槍1"
                            Case 2
                                tmpcardstr = "槍3/劍1"
                            Case 3
                                tmpcardstr = "防3/移1"
                            Case 4
                                tmpcardstr = "劍3/移1"
                            Case 5
                                tmpcardstr = "槍3/移1"
                            Case 6
                                tmpcardstr = "劍3/防1"
                            Case 7
                                tmpcardstr = "槍3/防1"
                            Case 8
                                tmpcardstr = "特2"
                        End Select
                    Case 1
                        Randomize
                        m = Int(Rnd() * 3) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "劍5/槍3"
                            Case 2
                                tmpcardstr = "劍5/移1"
                            Case 3
                                tmpcardstr = "劍8"
                        End Select
                    Case 2
                        Randomize
                        m = Int(Rnd() * 3) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "槍5/劍3"
                            Case 2
                                tmpcardstr = "槍5/移1"
                            Case 3
                                tmpcardstr = "槍8"
                        End Select
                    Case 3
                        Randomize
                        m = Int(Rnd() * 3) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "防5/移1"
                            Case 2
                                tmpcardstr = "防7"
                            Case 3
                                tmpcardstr = "HP回復3"
                        End Select
                    Case 4
                        Randomize
                        m = Int(Rnd() * 2) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "移3/特3"
                            Case 2
                                tmpcardstr = "移5"
                        End Select
                    Case 5
                        tmpcardstr = "機會5"
                    Case 6
                        tmpcardstr = "詛咒術5"
                    Case 7
                        Randomize
                        m = Int(Rnd() * 2) + 1
                        Select Case m
                            Case 1
                                tmpcardstr = "特3/防3"
                            Case 1
                                tmpcardstr = "特5"
                        End Select
                End Select
                戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
            If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
                For i = 7 To 18
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                    戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
                Next
            End If
        Else  '================================不遵守規則
            For i = 1 To 18
                Do
                    Randomize
                    m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                    '==============================
                    Select Case Formsetting.personcom(i).List(m)
                        Case "劍8"
                            Exit Do
                        Case "槍8"
                            Exit Do
                        Case "防7"
                            Exit Do
                        Case "移5"
                            Exit Do
                        Case "HP回復3"
                            Exit Do
                        Case "機會5"
                            Exit Do
                        Case "詛咒術5"
                            Exit Do
                        Case "特5"
                            Exit Do
                        Case "劍5/槍3"
                            Exit Do
                        Case "槍5/劍3"
                            Exit Do
                        Case "防5/移1"
                            Exit Do
                        Case "槍5/移1"
                            Exit Do
                        Case "劍5/移1"
                            Exit Do
                        Case "移3/特3"
                            Exit Do
                        Case "特3/防3"
                            Exit Do
                    End Select
                Loop
                tmpcardstr = Formsetting.personcom(i).List(m)
                戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        End If
    ElseIf Formsetting.comboeventcarrdcom.Text = "隨機" Or Formsetting.comboeventcarrdcom.Text = "隨機(不含聖水)" Then    '=====隨機
        If Formsetting.persontgrecom.Value = 1 Then    '===遵守規則
            For i = 1 To 18
                If i = 7 And 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then Exit For

                tmpfailed = 0
                Do
                    Randomize
                    m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                    If 一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                       (tmpfailed > 10 And 一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 1) = 0) Then
                        If Formsetting.comboeventcarrdcom.Text = "隨機(不含聖水)" And Formsetting.personcom(i).List(m) = "聖水" Then
                        Else
                            tmpcardstr = Formsetting.personcom(i).List(m)
                            Exit Do
                        End If
                    End If
                    tmpfailed = tmpfailed + 1
                Loop
                戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
            If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
                For i = 7 To 18
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "劍1"
                        Case 2
                            tmpcardstr = "槍1"
                        Case 3
                            tmpcardstr = "防1"
                    End Select
                    戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
                Next
            End If
        Else    '=============================不遵守規則
            For i = 1 To 18
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If Formsetting.comboeventcarrdcom.Text = "隨機(不含聖水)" And Formsetting.personcom(i).List(m) = "聖水" Then
                    i = i - 1
                Else
                    tmpcardstr = Formsetting.personcom(i).List(m)
                End If
                戰鬥系統類.發行卡牌_事件卡 2, tmpcardstr, 一般系統類.事件卡資料庫(tmpcardstr, 2)
            Next
        End If
    End If
End Sub
Sub 事件卡處理_分派_使用者方()
    If 戰鬥系統類.CardDeckCollection(3).Count > 0 Then
        Dim tmpcard As clsActionCard
        Set tmpcard = 戰鬥系統類.CardDeckCollection(3)(1)

        FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
        戰鬥系統類.座標計算_使用者手牌
        牌移動暫時變數(3) = tmpcard.CardNum
        戰鬥系統類.牌順序增加_手牌_使用者
        tmpcard.Location = 1
        tmpcard.Owner = 1
        tmpcard.XYLeft = 牌移動暫時變數(1)    '指定目前Left(座標)
        tmpcard.XYTop = 牌移動暫時變數(2)    '指定目前Top(座標)
        FormMainMode.card(tmpcard.CardNum).Left = 牌移動暫時變數(1)
        FormMainMode.card(tmpcard.CardNum).Top = 牌移動暫時變數(2)
        FormMainMode.card(tmpcard.CardNum).ZOrder
        FormMainMode.card(tmpcard.CardNum).Visible = True

        戰鬥系統類.卡牌牌堆集合更換 tmpcard, 3, 5
    End If
End Sub
Sub 事件卡處理_分派_電腦方()
    If 戰鬥系統類.CardDeckCollection(4).Count > 0 Then
        Dim i As Integer
        Dim tmpcard As clsActionCard
        Set tmpcard = 戰鬥系統類.CardDeckCollection(4)(1)

        FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
        戰鬥系統類.座標計算_電腦手牌
        牌移動暫時變數(3) = tmpcard.CardNum
        戰鬥系統類.公用牌變背面
        戰鬥系統類.牌順序增加_手牌_電腦
        tmpcard.Location = 1
        tmpcard.Owner = 2
        tmpcard.XYLeft = 牌移動暫時變數(1)    '指定目前Left(座標)
        tmpcard.XYTop = 牌移動暫時變數(2)    '指定目前Top(座標)
        FormMainMode.card(tmpcard.CardNum).Left = 牌移動暫時變數(1)
        FormMainMode.card(tmpcard.CardNum).Top = 牌移動暫時變數(2)
        FormMainMode.card(tmpcard.CardNum).ZOrder
        FormMainMode.card(tmpcard.CardNum).Visible = True
        戰鬥系統類.卡牌牌堆集合更換 tmpcard, 4, 7

        For i = 1 To 3
            FormMainMode.PEAFpersoncardcom(i).ZOrder
        Next
    End If
End Sub
Sub 事件卡處理_計算張數()
    If 角色人物對戰人數(1, 1) > 1 Or 角色人物對戰人數(2, 1) > 1 Then
        事件卡記錄暫時數(0, 1) = 18
    Else
        事件卡記錄暫時數(0, 1) = 12
    End If
End Sub
Sub 執行動作_防禦階段結束時技能啟動()
    '===========================執行階段插入點(ATK-14/DEF-34)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 14, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 34, 2
    '============================
    '===========================執行階段插入點(ATK-15/DEF-35)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 15, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 35, 2
    '============================
    '===========================執行階段插入點(ATK-16/DEF-36)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 16, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 36, 2
    '============================
    HP檢查階段數 = 3
    戰鬥系統類.雙方HP檢查
End Sub
Sub 執行動作_攻擊階段結束時技能啟動()
    '===========================執行階段插入點(ATK-14/DEF-34)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 14, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 34, 2
    '============================
    '===========================執行階段插入點(ATK-15/DEF-35)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 15, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 35, 2
    '============================
    '===========================執行階段插入點(ATK-16/DEF-36)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 16, 2
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 36, 2
    '============================
    HP檢查階段數 = 3
    戰鬥系統類.雙方HP檢查
End Sub
Sub 技能說明載入_人物卡片背面_使用者(ByVal n As Integer)
    Dim i As Integer
    For i = 1 To 4
        '==============================主動技
        FormMainMode.cardus(n).CardBack_主動技_技能名稱 i, VBEPerson(1, n, 3, i, 1)
        FormMainMode.cardus(n).CardBack_主動技_階段代碼 i, Val(VBEPerson(1, n, 3, i, 8))
        FormMainMode.cardus(n).CardBack_主動技_距離代碼 i, VBEPerson(1, n, 3, i, 9)
        FormMainMode.cardus(n).CardBack_主動技_卡片代碼 i, VBEPerson(1, n, 3, i, 10)
        FormMainMode.cardus(n).CardBack_主動技_技能說明 i, VBEPerson(1, n, 3, i, 5)
        '==============================被動技
        FormMainMode.cardus(n).CardBack_被動技_技能名稱 i, VBEPerson(1, n, 3, i + 4, 1)
        FormMainMode.cardus(n).CardBack_被動技_技能說明 i, VBEPerson(1, n, 3, i + 4, 2)
    Next
End Sub
Sub 技能說明載入_人物卡片背面_電腦(ByVal n As Integer)
    Dim i As Integer
    For i = 1 To 4
        '==============================主動技
        FormMainMode.cardcom(n).CardBack_主動技_技能名稱 i, VBEPerson(2, n, 3, i, 1)
        FormMainMode.cardcom(n).CardBack_主動技_階段代碼 i, Val(VBEPerson(2, n, 3, i, 8))
        FormMainMode.cardcom(n).CardBack_主動技_距離代碼 i, VBEPerson(2, n, 3, i, 9)
        FormMainMode.cardcom(n).CardBack_主動技_卡片代碼 i, VBEPerson(2, n, 3, i, 10)
        FormMainMode.cardcom(n).CardBack_主動技_技能說明 i, VBEPerson(2, n, 3, i, 5)
        '==============================被動技
        FormMainMode.cardcom(n).CardBack_被動技_技能名稱 i, VBEPerson(2, n, 3, i + 4, 1)
        FormMainMode.cardcom(n).CardBack_被動技_技能說明 i, VBEPerson(2, n, 3, i + 4, 2)
    Next
End Sub

Sub 技能說明載入_人物卡片背面_交換角色()
    Dim n As Integer, i As Integer
    For n = 1 To 2
        For i = 1 To 4
            '==============================主動技
            Formchangeperson.card(n).CardBack_主動技_技能名稱 i, VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 1)
            Formchangeperson.card(n).CardBack_主動技_階段代碼 i, Val(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 8))
            Formchangeperson.card(n).CardBack_主動技_距離代碼 i, VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9)
            Formchangeperson.card(n).CardBack_主動技_卡片代碼 i, VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 10)
            Formchangeperson.card(n).CardBack_主動技_技能說明 i, VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 5)
            '==============================被動技
            Formchangeperson.card(n).CardBack_被動技_技能名稱 i, VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i + 4, 1)
            Formchangeperson.card(n).CardBack_被動技_技能說明 i, VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i + 4, 2)
        Next
    Next
End Sub
Sub 發行卡牌_公用牌()
    Dim i As Integer, j As Integer, tmpnewActionCardNum As Integer
    Dim tmpcard As clsActionCard, tmpindexobj As clsCollectionIndex

    For i = 1 To 公用牌各牌類型紀錄數(0, 2)
        tmpnewActionCardNum = 戰鬥系統類.遊戲實體牌物件創建發行

        Set tmpcard = New clsActionCard
        tmpcard.CardNum = tmpnewActionCardNum
        tmpcard.Location = 4
        tmpcard.CardType = 1
        For j = 1 To UBound(公用牌各牌類型紀錄數, 1)
            If Val(公用牌各牌類型紀錄數(j, 1)) < Val(公用牌各牌類型紀錄數(j, 2)) Then
                Select Case j
                    Case 1  '==移1槍1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\021.png"
                        tmpcard.ImageStr = "021"
                        tmpcard.CardOnIn = 1
                    Case 2  '==移1槍2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\019.png"
                        tmpcard.ImageStr = "019"
                        tmpcard.CardOnIn = 1
                    Case 3  '==移1槍3類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b3b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\017.png"
                        tmpcard.ImageStr = "017"
                        tmpcard.CardOnIn = 1
                    Case 4  '==移1盾1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\025.png"
                        tmpcard.ImageStr = "025"
                        tmpcard.CardOnIn = 1
                    Case 5  '==移1盾2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\024.png"
                        tmpcard.ImageStr = "024"
                        tmpcard.CardOnIn = 1
                    Case 6  '==移1盾3類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b3b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\023.png"
                        tmpcard.ImageStr = "023"
                        tmpcard.CardOnIn = 1
                    Case 7  '==移2特3類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b3b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\026.png"
                        tmpcard.ImageStr = "026"
                        tmpcard.CardOnIn = 1
                    Case 8  '==移3移3類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a3a
                        tmpcard.UpperNum = b3b
                        tmpcard.LowerType = a3a
                        tmpcard.LowerNum = b3b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\027.png"
                        tmpcard.ImageStr = "027"
                        tmpcard.CardOnIn = 1
                    Case 9  '==劍6劍6類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b6b
                        tmpcard.LowerType = a1a
                        tmpcard.LowerNum = b6b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\001.png"
                        tmpcard.ImageStr = "001"
                        tmpcard.CardOnIn = 1
                    Case 10  '==劍1槍1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\011.png"
                        tmpcard.ImageStr = "011"
                        tmpcard.CardOnIn = 1
                    Case 11  '==劍2槍1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\007.png"
                        tmpcard.ImageStr = "007"
                        tmpcard.CardOnIn = 1
                    Case 12  '==劍2槍2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\006.png"
                        tmpcard.ImageStr = "006"
                        tmpcard.CardOnIn = 1
                    Case 13  '==劍3槍3類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b3b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b3b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\004.png"
                        tmpcard.ImageStr = "004"
                        tmpcard.CardOnIn = 1
                    Case 14  '==劍5槍5類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b5b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b5b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\028.png"
                        tmpcard.ImageStr = "028"
                        tmpcard.CardOnIn = 1
                    Case 15  '==劍1盾1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\012.png"
                        tmpcard.ImageStr = "012"
                        tmpcard.CardOnIn = 1
                    Case 16  '==劍2盾1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\009.png"
                        tmpcard.ImageStr = "009"
                        tmpcard.CardOnIn = 1
                    Case 17  '==劍2盾2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\008.png"
                        tmpcard.ImageStr = "008"
                        tmpcard.CardOnIn = 1
                    Case 18  '==劍3盾3類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b3b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b3b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\005.png"
                        tmpcard.ImageStr = "005"
                        tmpcard.CardOnIn = 1
                    Case 19  '==劍1特1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b1b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\013.png"
                        tmpcard.ImageStr = "013"
                        tmpcard.CardOnIn = 1
                    Case 20  '==劍2特1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\010.png"
                        tmpcard.ImageStr = "010"
                        tmpcard.CardOnIn = 1
                    Case 21  '==劍4特1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b4b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\003.png"
                        tmpcard.ImageStr = "003"
                        tmpcard.CardOnIn = 1
                    Case 22  '==劍5特2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a1a
                        tmpcard.UpperNum = b5b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\002.png"
                        tmpcard.ImageStr = "002"
                        tmpcard.CardOnIn = 1
                    Case 23  '==槍4槍4類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a5a
                        tmpcard.UpperNum = b4b
                        tmpcard.LowerType = a5a
                        tmpcard.LowerNum = b4b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\015.png"
                        tmpcard.ImageStr = "015"
                        tmpcard.CardOnIn = 1
                    Case 24  '==槍2特1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a5a
                        tmpcard.UpperNum = b2b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\020.png"
                        tmpcard.ImageStr = "020"
                        tmpcard.CardOnIn = 1
                    Case 25  '==槍3特2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a5a
                        tmpcard.UpperNum = b3b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\018.png"
                        tmpcard.ImageStr = "018"
                        tmpcard.CardOnIn = 1
                    Case 26  '==槍4特1類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a5a
                        tmpcard.UpperNum = b4b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b1b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\016.png"
                        tmpcard.ImageStr = "016"
                        tmpcard.CardOnIn = 1
                    Case 27  '==槍5特2類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a5a
                        tmpcard.UpperNum = b5b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b2b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\014.png"
                        tmpcard.ImageStr = "014"
                        tmpcard.CardOnIn = 1
                    Case 28  '==盾5盾5類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a2a
                        tmpcard.UpperNum = b5b
                        tmpcard.LowerType = a2a
                        tmpcard.LowerNum = b5b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\022.png"
                        tmpcard.ImageStr = "022"
                        tmpcard.CardOnIn = 1
                    Case 29  '==盾3特5類
                        公用牌各牌類型紀錄數(j, 1) = Val(公用牌各牌類型紀錄數(j, 1)) + 1
                        公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
                        tmpcard.UpperType = a2a
                        tmpcard.UpperNum = b3b
                        tmpcard.LowerType = a4a
                        tmpcard.LowerNum = b5b
                        tmpcard.Owner = 0
                        FormMainMode.card(i).cardImage = app_path & "card\029.png"
                        tmpcard.ImageStr = "029"
                        tmpcard.CardOnIn = 1
                End Select
                戰鬥系統類.CardDeckCollection(1).Add tmpcard, CStr(tmpnewActionCardNum)

                Set tmpindexobj = New clsCollectionIndex
                tmpindexobj.CollectionIndex = 1
                tmpindexobj.CardNum = tmpnewActionCardNum
                tmpindexobj.Index = 戰鬥系統類.CardDeckCollection(0).Count + 1
                戰鬥系統類.CardDeckCollection(0).Add tmpindexobj, CStr(tmpnewActionCardNum)
                Exit For
            End If
        Next
    Next
End Sub
Sub 發行卡牌_事件卡(ByVal uscom As Integer, ByVal cardname As String, ByVal filename As String, Optional ByVal beforeindex As Integer)
    Dim ay() As String
    Dim tn As Integer, tmpnewActionCardNum As Integer
    Dim tmpcard As clsActionCard
    Dim tmpindexobj As clsCollectionIndex

    If cardname = "" Or filename = "" Then Exit Sub

    tmpnewActionCardNum = 戰鬥系統類.遊戲實體牌物件創建發行

    Set tmpcard = New clsActionCard
    tmpcard.CardNum = tmpnewActionCardNum
    '============
    Erase ay
    ay = Split(一般系統類.事件卡資料庫(cardname, 3), "=")
    '============
    tmpcard.UpperType = ay(0)
    tmpcard.UpperNum = ay(1)
    tmpcard.LowerType = ay(2)
    tmpcard.LowerNum = ay(3)
    tmpcard.Owner = 0
    tmpcard.Location = 0
    tmpcard.ImageStr = filename
    tmpcard.ComMark = 0
    FormMainMode.card(tmpnewActionCardNum).cardImage = app_path & "card\" & filename & ".png"
    FormMainMode.card(tmpnewActionCardNum).CardRotationType = 1
    tmpcard.CardOnIn = 1
    tmpcard.CardType = 2

    Set tmpindexobj = New clsCollectionIndex
    tmpindexobj.CardNum = tmpnewActionCardNum
    tmpindexobj.Index = 戰鬥系統類.CardDeckCollection(0).Count + 1

    Select Case uscom
        Case 1
            tmpindexobj.CollectionIndex = 3
            戰鬥系統類.CardDeckCollection(0).Add tmpindexobj, CStr(tmpnewActionCardNum)

            If beforeindex > 0 And beforeindex <= 戰鬥系統類.CardDeckCollection(3).Count Then
                戰鬥系統類.CardDeckCollection(3).Add tmpcard, CStr(tmpnewActionCardNum), beforeindex
            Else
                戰鬥系統類.CardDeckCollection(3).Add tmpcard, CStr(tmpnewActionCardNum)
            End If
        Case 2
            tmpindexobj.CollectionIndex = 4
            戰鬥系統類.CardDeckCollection(0).Add tmpindexobj, CStr(tmpnewActionCardNum)

            If beforeindex > 0 And beforeindex <= 戰鬥系統類.CardDeckCollection(4).Count Then
                戰鬥系統類.CardDeckCollection(4).Add tmpcard, CStr(tmpnewActionCardNum), beforeindex
            Else
                戰鬥系統類.CardDeckCollection(4).Add tmpcard, CStr(tmpnewActionCardNum)
            End If
    End Select
End Sub
Sub 公用牌地圖牌種類配置(ByVal name As String)
    Select Case name
        Case "萊丁貝魯格城堡"
            公用牌各牌類型紀錄數(0, 2) = 57
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 0
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 0
        Case "誘惑森林"
            公用牌各牌類型紀錄數(0, 2) = 50
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 2
            公用牌各牌類型紀錄數(16, 2) = 0
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 0
            公用牌各牌類型紀錄數(20, 2) = 0
            公用牌各牌類型紀錄數(21, 2) = 1
            公用牌各牌類型紀錄數(22, 2) = 1
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 0
        Case "垃圾之街"
            公用牌各牌類型紀錄數(0, 2) = 55
            公用牌各牌類型紀錄數(1, 2) = 2
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 0
        Case "冰封湖畔(新)"
            公用牌各牌類型紀錄數(0, 2) = 53
            公用牌各牌類型紀錄數(1, 2) = 4
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 2
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "人魂墓地"
            公用牌各牌類型紀錄數(0, 2) = 50
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 4
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 0
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 2
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 0
        Case "盡頭之村"
            公用牌各牌類型紀錄數(0, 2) = 54
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 0
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 0
            公用牌各牌類型紀錄數(26, 2) = 1
            公用牌各牌類型紀錄數(27, 2) = 0
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 0
        Case "風暴荒野"
            公用牌各牌類型紀錄數(0, 2) = 52
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 2
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 2
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 0
            公用牌各牌類型紀錄數(20, 2) = 0
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 1
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 0
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "藩骸兒的遺跡"
            公用牌各牌類型紀錄數(0, 2) = 49
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 1
            公用牌各牌類型紀錄數(3, 2) = 1
            公用牌各牌類型紀錄數(4, 2) = 3
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 2
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 1
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 2
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 1
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "魔都羅占布爾克"
            公用牌各牌類型紀錄數(0, 2) = 42
            公用牌各牌類型紀錄數(1, 2) = 0
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 2
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 0
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 0
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "瘋狂山脈"
            公用牌各牌類型紀錄數(0, 2) = 47
            公用牌各牌類型紀錄數(1, 2) = 2
            公用牌各牌類型紀錄數(2, 2) = 0
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 2
            公用牌各牌類型紀錄數(5, 2) = 0
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 2
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "魔女山谷"
            公用牌各牌類型紀錄數(0, 2) = 52
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 3
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 3
            公用牌各牌類型紀錄數(11, 2) = 1
            公用牌各牌類型紀錄數(12, 2) = 1
            公用牌各牌類型紀錄數(13, 2) = 0
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "烏波斯的黑湖"
            公用牌各牌類型紀錄數(0, 2) = 50
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 1
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 2
            公用牌各牌類型紀錄數(8, 2) = 1
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 2
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 1
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 0
            公用牌各牌類型紀錄數(21, 2) = 1
            公用牌各牌類型紀錄數(22, 2) = 1
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 0
            公用牌各牌類型紀錄數(26, 2) = 1
            公用牌各牌類型紀錄數(27, 2) = 1
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 1
        Case "白魔的圓環石陣"
            公用牌各牌類型紀錄數(0, 2) = 50
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 2
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 0
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 0
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 1
            公用牌各牌類型紀錄數(22, 2) = 1
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 1
            公用牌各牌類型紀錄數(27, 2) = 1
            公用牌各牌類型紀錄數(28, 2) = 0
            公用牌各牌類型紀錄數(29, 2) = 0
        Case Else
            公用牌各牌類型紀錄數(0, 2) = 57
            公用牌各牌類型紀錄數(1, 2) = 6
            公用牌各牌類型紀錄數(2, 2) = 2
            公用牌各牌類型紀錄數(3, 2) = 2
            公用牌各牌類型紀錄數(4, 2) = 6
            公用牌各牌類型紀錄數(5, 2) = 2
            公用牌各牌類型紀錄數(6, 2) = 1
            公用牌各牌類型紀錄數(7, 2) = 3
            公用牌各牌類型紀錄數(8, 2) = 0
            公用牌各牌類型紀錄數(9, 2) = 1
            公用牌各牌類型紀錄數(10, 2) = 4
            公用牌各牌類型紀錄數(11, 2) = 2
            公用牌各牌類型紀錄數(12, 2) = 2
            公用牌各牌類型紀錄數(13, 2) = 2
            公用牌各牌類型紀錄數(14, 2) = 0
            公用牌各牌類型紀錄數(15, 2) = 4
            公用牌各牌類型紀錄數(16, 2) = 2
            公用牌各牌類型紀錄數(17, 2) = 2
            公用牌各牌類型紀錄數(18, 2) = 2
            公用牌各牌類型紀錄數(19, 2) = 1
            公用牌各牌類型紀錄數(20, 2) = 1
            公用牌各牌類型紀錄數(21, 2) = 2
            公用牌各牌類型紀錄數(22, 2) = 2
            公用牌各牌類型紀錄數(23, 2) = 1
            公用牌各牌類型紀錄數(24, 2) = 1
            公用牌各牌類型紀錄數(25, 2) = 1
            公用牌各牌類型紀錄數(26, 2) = 2
            公用牌各牌類型紀錄數(27, 2) = 2
            公用牌各牌類型紀錄數(28, 2) = 1
            公用牌各牌類型紀錄數(29, 2) = 0
    End Select
End Sub
Sub 傷害執行_立即死亡_使用者(ByVal num As Integer)
    Dim stageInfoListObj As clsVSStageObj
    Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
    '===============================
    VBEStageNum(0) = 46
    VBEStageNum(1) = -1    '受到傷害方(1.使用者/2.電腦)
    VBEStageNum(2) = num    '受到傷害人物編號
    VBEStageNum(3) = 3    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    VBEStageNum(4) = liveus(角色待機人物紀錄數(1, num))  '受到傷害之數值(現有HP)
    stageInfoListObj.Argument = liveus(角色待機人物紀錄數(1, num))   '受到傷害之數值
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1"    '受到傷害方(1.使用者/2.電腦)
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num)    '受到傷害人物編號
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "3"    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    '===========================執行階段插入點(46)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 46, 1
    '============================
    If stageInfoListObj.CommandStr = "PersonBloodControl" Then
        If stageInfoListObj.Value = "BLOODOFF" Then
            Exit Sub
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                戰鬥系統類.傷害執行_技能直傷_使用者 Val(tmpstr(1)), num, False
                Exit Sub
            End If
        End If
    End If
    '============================
    Select Case num
        Case 1
            戰鬥系統類.廣播訊息 "您受到了" & liveus(角色人物對戰人數(1, 2)) & "點傷害。"
            FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = 0
            FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).CurrentHP = 0
            liveus(角色人物對戰人數(1, 2)) = 0
            FormMainMode.bloodnumus1.Caption = 0
            FormMainMode.bloodlineout1.Width = 0
            牌總階段數(1) = 牌總階段數(1) + 1
            戰鬥系統類.播放傷害音樂
        Case Is > 1
            liveus(角色待機人物紀錄數(1, num)) = 0
            FormMainMode.cardus(角色待機人物紀錄數(1, num)).CardMain_角色HP = 0
            FormMainMode.PEAFpersoncardus(角色待機人物紀錄數(1, num)).CurrentHP = 0
            牌總階段數(1) = 牌總階段數(1) + 1
    End Select
End Sub
Sub 傷害執行_立即死亡_電腦(ByVal num As Integer)
    Dim stageInfoListObj As clsVSStageObj
    Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
    '===============================
    VBEStageNum(0) = 46
    VBEStageNum(1) = -2    '受到傷害方(1.使用者/2.電腦)
    VBEStageNum(2) = num    '受到傷害人物編號
    VBEStageNum(3) = 3    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    VBEStageNum(4) = livecom(角色待機人物紀錄數(2, num))    '受到傷害之數值(現有HP)
    stageInfoListObj.Argument = livecom(角色待機人物紀錄數(2, num))  '受到傷害之數值
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2"    '受到傷害方(1.使用者/2.電腦)
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num)    '受到傷害人物編號
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "3"    '受到傷害之形式(1.骰傷/2.直傷/3.立即死亡)
    '===========================執行階段插入點(46)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 46, 1
    '============================
    If stageInfoListObj.CommandStr = "PersonBloodControl" Then
        If stageInfoListObj.Value = "BLOODOFF" Then
            Exit Sub
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                戰鬥系統類.傷害執行_技能直傷_電腦 Val(tmpstr(1)), num, False
                Exit Sub
            End If
        End If
    End If
    '============================
    Select Case num
        Case 1
            戰鬥系統類.廣播訊息 "對方受到了" & livecom(角色人物對戰人數(2, 2)) & "點傷害。"
            FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).CurrentHP = 0
            FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = 0
            FormMainMode.bloodnumcom1.Caption = 0
            livecom(角色人物對戰人數(2, 2)) = 0
            FormMainMode.bloodlineout2.Left = 11580
            牌總階段數(2) = 牌總階段數(2) + 1
            戰鬥系統類.播放傷害音樂
        Case Is > 1
            FormMainMode.cardcom(角色待機人物紀錄數(2, num)).CardMain_角色HP = 0
            livecom(角色待機人物紀錄數(2, num)) = 0
            FormMainMode.PEAFpersoncardcom(角色待機人物紀錄數(2, num)).CurrentHP = 0
            牌總階段數(2) = 牌總階段數(2) + 1
    End Select
End Sub
Sub 角色復活_使用者(ByVal num As Integer)
    If liveus(角色待機人物紀錄數(1, num)) > 0 Then Exit Sub
    '===============================
    Dim stageInfoListObj As New clsVSStageObj
    Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
    VBEStageNum(0) = 49
    VBEStageNum(1) = -1    '角色復活方(1.使用者/2.電腦)
    VBEStageNum(2) = num    '角色復活人物編號
    '===========================執行階段插入點(49)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, 49, 1
    '============================
    Dim tmpflag As Boolean
    tmpflag = False
    If stageInfoListObj.CommandStr = "PersonResurrect" Then
        If stageInfoListObj.Value = "OFF" Then
            tmpflag = True
        End If
    End If

    If tmpflag = False Then
        Select Case num
            Case 1
                FormMainMode.cardus(角色人物對戰人數(1, 2)).CardMain_角色HP = 1
                FormMainMode.PEAFpersoncardus(角色人物對戰人數(1, 2)).CurrentHP = 1
                liveus(角色人物對戰人數(1, 2)) = 1
                FormMainMode.bloodlineout1.Width = 距離單位(1, 1, 1)
                FormMainMode.bloodnumus1.Caption = liveus(角色人物對戰人數(1, 2))
            Case Is > 1
                liveus(角色待機人物紀錄數(1, num)) = 1
                FormMainMode.PEAFpersoncardus(角色待機人物紀錄數(1, num)).CurrentHP = 1
                FormMainMode.cardus(角色待機人物紀錄數(1, num)).CardMain_角色HP = 1
        End Select
    End If
End Sub
Sub 角色復活_電腦(ByVal num As Integer)
    '===============================
    If livecom(角色待機人物紀錄數(2, num)) > 0 Then Exit Sub
    '===============================
    Dim stageInfoListObj As New clsVSStageObj
    Set stageInfoListObj = 執行階段系統類.VBEVSStageInfoList(執行階段系統類.VBEVSStageInfoList.Count)
    VBEStageNum(0) = 49
    VBEStageNum(1) = -2    '角色復活方(1.使用者/2.電腦)
    VBEStageNum(2) = num    '角色復活人物編號
    '===========================執行階段插入點(49)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 2, 49, 1
    '============================
    Dim tmpflag As Boolean
    tmpflag = False
    If stageInfoListObj.CommandStr = "PersonResurrect" Then
        If stageInfoListObj.Value = "OFF" Then
            tmpflag = True
        End If
    End If

    If tmpflag = False Then
        Select Case num
            Case 1
                FormMainMode.PEAFpersoncardcom(角色人物對戰人數(2, 2)).CurrentHP = 1
                FormMainMode.cardcom(角色人物對戰人數(2, 2)).CardMain_角色HP = 1
                FormMainMode.bloodnumcom1.Caption = 1
                livecom(角色人物對戰人數(2, 2)) = 1
                FormMainMode.bloodlineout2.Left = 11580 - 距離單位(1, 2, 1)
            Case Is > 1
                FormMainMode.cardcom(角色待機人物紀錄數(2, num)).CardMain_角色HP = 1
                livecom(角色待機人物紀錄數(2, num)) = 1
                FormMainMode.PEAFpersoncardcom(角色待機人物紀錄數(2, num)).CurrentHP = 1
        End Select
    End If
End Sub
Sub 解析骰量變化(ByVal str As String, ByVal uscom As Integer)
    Dim cmdstr() As String
    Dim i As Integer

    cmdstr = Split(str, "=")
    If 顯示列雙方數值鎖定紀錄數(uscom) = False Then
        For i = 0 To UBound(cmdstr) - 1
            Select Case Mid(cmdstr(i), 1, 1)
                Case "+"
                    攻擊防禦骰子總數(uscom) = 攻擊防禦骰子總數(uscom) + Mid(cmdstr(i), 2, Len(cmdstr(i)))
                Case "-"
                    攻擊防禦骰子總數(uscom) = 攻擊防禦骰子總數(uscom) - Mid(cmdstr(i), 2, Len(cmdstr(i)))
                Case "*"
                    攻擊防禦骰子總數(uscom) = 攻擊防禦骰子總數(uscom) * Mid(cmdstr(i), 2, Len(cmdstr(i)))
                Case "/"
                    攻擊防禦骰子總數(uscom) = Int(攻擊防禦骰子總數(uscom) / Mid(cmdstr(i), 2, Len(cmdstr(i))) + 0.9)
                Case "\"
                    攻擊防禦骰子總數(uscom) = 攻擊防禦骰子總數(uscom) \ Mid(cmdstr(i), 2, Len(cmdstr(i)))
                Case "@"
                    攻擊防禦骰子總數(uscom) = Mid(cmdstr(i), 2, Len(cmdstr(i)))
                    顯示列雙方數值鎖定紀錄數(uscom) = True
                    Exit Sub    '==指定數值時其他變化量無效
            End Select
        Next
    End If
End Sub
Sub 遊戲對戰結束物件消滅()
    Dim i As Integer
    '==========
    For i = 1 To FormMainMode.PEAFvssc.UBound
        Unload FormMainMode.PEAFvssc(i)
    Next
    '==========
    '==========
    For i = 1 To FormMainMode.card.UBound
        Unload FormMainMode.card(i)
    Next
    '==========
    For i = 1 To FormMainMode.cardus.UBound
        Unload FormMainMode.cardus(i)
    Next
    For i = 1 To FormMainMode.cardcom.UBound
        Unload FormMainMode.cardcom(i)
    Next
    '==========
    Erase 執行階段系統類.VSSCObjectCollection
End Sub
Function 遊戲實體牌物件創建發行() As Integer
    Dim i As Integer

    戰鬥系統類.ActionCardTotNum = 戰鬥系統類.ActionCardTotNum + 1
    i = 戰鬥系統類.ActionCardTotNum

    Load FormMainMode.card(i)
    Set FormMainMode.card(i).Container = FormMainMode.PEAttackingForm
    FormMainMode.card(i).Left = 240
    FormMainMode.card(i).Top = 960
    FormMainMode.card(i).Visible = False
    FormMainMode.card(i).CardEventType = False
    FormMainMode.card(i).LocationType = 0

    遊戲實體牌物件創建發行 = 戰鬥系統類.ActionCardTotNum
End Function
Sub 廣播訊息(ByVal messagestr As String)
    FormMainMode.PEAFInterface.Message messagestr
End Sub
Sub 遊戲角色卡片物件創立()
    Dim i As Integer

    For i = 1 To 3
        Load FormMainMode.cardus(i)
        Load FormMainMode.cardcom(i)
    Next
End Sub
Sub 執行動作_系統總卡牌張數更新()
    FormMainMode.PEAFInterface.CardNum = BattleCardNum
    FormMainMode.pageul.Caption = BattleCardNum
End Sub
Sub 執行動作_電腦方各階段出牌完畢後行動(ByVal turnnum As Integer)
    Dim ckl As Integer

    Select Case turnnum
        Case 1
            FormMainMode.攻擊階段_階段2.Enabled = True
        Case 2
            FormMainMode.PEAFInterface.BnOKStartListen
            '==============
            小人物頭像移動方向數(1) = 1
            小人物頭像移動方向數(2) = 2
            FormMainMode.小人物頭像移動_使用者.Enabled = True
            FormMainMode.小人物頭像移動_電腦.Enabled = True
            '==============
            階段狀態數 = 1
            一般系統類.音效播放 6
            戰鬥系統類.時間軸_重設
            FormMainMode.trtimeline.Enabled = True
        Case 3
            turnpageonin = 1
            階段狀態數 = 1
            FormMainMode.PEAFInterface.BnOKStartListen
            If Vss_EventPlayerAllActionOffNum(1) = 1 Then
                For ckl = 1 To 戰鬥系統類.ActionCardTotNum
                    FormMainMode.card(ckl).CardEnabledType = False
                Next
                FormMainMode.PEAFInterface.BnOKEnabled False
                等待時間佇列(2).Add 47
                FormMainMode.等待時間_2.Enabled = True
            ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
                For ckl = 1 To 戰鬥系統類.ActionCardTotNum
                    FormMainMode.card(ckl).CardEnabledType = False
                Next
                FormMainMode.PEAFInterface.BnOKEnabled False
                等待時間佇列(2).Add 45
                FormMainMode.等待時間_2.Enabled = True
            End If
    End Select
End Sub
Sub 移動階段移動前執行階段呼叫(ByVal ns As Integer)
    Dim moveusTempnum As Integer, movecomTempnum As Integer, moveusSelectnum As Integer, movecomSelectnum As Integer
    If Vss_PersonMoveControlNum(1, 2) = 0 Then
        moveusTempnum = moveus + Vss_PersonMoveControlNum(1, 1)
    Else
        moveusTempnum = Vss_PersonMoveControlNum(1, 1)
    End If
    If Vss_PersonMoveControlNum(2, 2) = 0 Then
        movecomTempnum = movecom + Vss_PersonMoveControlNum(2, 1)
    Else
        movecomTempnum = Vss_PersonMoveControlNum(2, 1)
    End If
    '==================================
    If moveusTempnum < 0 Then moveusTempnum = 0
    If movecomTempnum < 0 Then movecomTempnum = 0
    '==================================
    If Vss_PersonMoveActionChangeNum(1, 1) = 1 Then
        moveusSelectnum = Vss_PersonMoveActionChangeNum(1, 2)
    Else
        moveusSelectnum = FormMainMode.顯示列1.移動階段選擇值
    End If
    If Vss_PersonMoveActionChangeNum(2, 1) = 1 Then
        movecomSelectnum = Vss_PersonMoveActionChangeNum(2, 2)
    Else
        movecomSelectnum = 電腦方移動階段選擇數
        If movecomTempnum <= 0 Then
            movecomSelectnum = 2
        End If
    End If
    '===============
    If Vss_EventPlayerAllActionOffNum(1) = 1 Then moveusSelectnum = 0
    If Vss_EventPlayerAllActionOffNum(2) = 1 Then movecomSelectnum = 0
    ReDim VBEStageNum(0 To 4) As Integer
    VBEStageNum(0) = ns
    VBEStageNum(1) = moveusTempnum    '使用者方總移動數
    VBEStageNum(2) = movecomTempnum    '電腦方總移動數
    VBEStageNum(3) = moveusSelectnum    '使用者方目前移動階段行動選擇
    VBEStageNum(4) = movecomSelectnum    '電腦方目前移動階段行動選擇
    '===========================執行階段插入點(ns)
    執行階段系統類.執行階段系統總主要程序_執行階段開始 1, ns, 1
    '============================
End Sub
Sub 介面角色小卡資訊寫入(ByVal uscom As Integer, ByVal num As Integer)
    'Dim tmpobj As New clsPersonCard

    Select Case uscom
        Case 1
            With FormMainMode.PEAFpersoncardus(num)
                .Level = uslevel(num)
                .ATK = atkus(num)
                .DEF = defus(num)
                .CurrentHP = liveus(num)
                .AllHP = liveusmax(num)
                .personName = nameus(num)
            End With
        Case 2
            With FormMainMode.PEAFpersoncardcom(num)
                .Level = comlevel(num)
                .ATK = atkcom(num)
                .DEF = defcom(num)
                .CurrentHP = livecom(num)
                .AllHP = livecommax(num)
                .personName = namecom(num)
            End With
    End Select
End Sub
Sub 卡牌牌堆集合更換(ByRef tmpcard As clsActionCard, ByVal torigc As Integer, ByVal tnewc As Integer)
    Dim tmpindexobj As clsCollectionIndex
    Set tmpindexobj = 戰鬥系統類.CardDeckCollection(0)(CStr(tmpcard.CardNum))

    戰鬥系統類.CardDeckCollection(torigc).Remove CStr(tmpcard.CardNum)
    戰鬥系統類.CardDeckCollection(tnewc).Add tmpcard, CStr(tmpcard.CardNum)
    '=========索引更新
    tmpindexobj.CollectionIndex = tnewc
End Sub
Function 卡牌牌堆集合索引_CollectionIndex(ByVal tmpindex As Variant) As Integer
    Dim tmpindexobj As clsCollectionIndex
    Set tmpindexobj = 戰鬥系統類.CardDeckCollection(0)(tmpindex)
    卡牌牌堆集合索引_CollectionIndex = tmpindexobj.CollectionIndex
End Function
Function 卡牌牌堆集合索引_CardNum(ByVal tmpindex As Variant) As Integer
    Dim tmpindexobj As clsCollectionIndex
    Set tmpindexobj = 戰鬥系統類.CardDeckCollection(0)(tmpindex)
    卡牌牌堆集合索引_CardNum = tmpindexobj.CardNum
End Function
Function 卡牌牌堆集合索引_Index(ByVal tmpindex As Variant) As Integer
    Dim tmpindexobj As clsCollectionIndex
    Set tmpindexobj = 戰鬥系統類.CardDeckCollection(0)(tmpindex)
    卡牌牌堆集合索引_Index = tmpindexobj.Index
End Function
