Attribute VB_Name = "智慧型AI系統類"
Public cardcountAInum() As String  '公用牌計算暫時基本資料(第x張,1.正面類型/2.正面數值/3.反面類型/4.反面數值/5.牌編號)
Public cardcountAInumMOV() As String  '公用牌計算暫時基本資料-移動階段續-原本(第x張,1.正面類型/2.正面數值/3.反面類型/4.反面數值/5.牌編號)
Dim cardAIn() As Integer '排列組合計算暫時變數
Dim cardAInumans As String '排列組合計算暫時變數
Public cardAInumnm() As String '排列組合計算最終數值
Public cardAInumFinal() As Integer '排列組合計算最終期望值
Public cardAInumFinal2() As Integer '排列組合計算最終期望值-排列後
Public cardAInumcase(1 To 5, 1 To 2) As Integer '公用牌計算統計資料(1.ATK-劍/2.DEF/3.MOV/4.SPE/5.ATK-槍,1.組合下最低數值/2.組合下最高數值)
Public cardAInumcaseperson() As Integer '公用牌計算統計暫時資料-個別組合
Public cardAInumuscom As Integer '手牌擁有者牌數記錄暫時變數
Public cardAITotalNUM As Integer '排列組合計算總共組合數
Public cardAInumcasepersonTER() As Integer '公用牌計算統計暫時資料-個別組合-個別卡面數值計數統計
Public cardAInumselect1 As Integer  '公用牌計算統計比序暫時變數-目前最高期望值
Public cardAInumselect4 As Integer  '公用牌計算統計比序暫時變數-目前最高個別加總期望值
Public cardAInumselect2 As String '公用牌計算統計比序暫時變數-目前最高期望值下編號串-初始
Public cardAInumselect3() As String '公用牌計算統計比序暫時變數-目前最高期望值下編號串-陣列
Public cardAInumchoose As Integer '公用牌計算最終選擇組合編號
Public cardAInumMOVmain(1 To 2, 1 To 15) As String 'AI-移動階段續-組合暫時紀錄
Public cardAInumMOVnm() As String 'AI-移動階段續-正向面-計算排列組合串暫時紀錄
Public cardAInumMOVnmtot() As String 'AI-移動階段續-正向面-總共排列組合串相關資料暫時紀錄
Public cardAInumMOVFinal(1 To 3) As String 'AI-移動階段續-正向面-最終結果紀錄數(1.最終排列組合串/2.最終排列組合編號/3.最終選定目標距離[1.近/2.遠])
Public 是否移動階段續估計判斷程序 As Boolean 'AI-移動階段續-是否為估計判斷程序標記數
Public cardAInumOvertenrecord() As Integer 'AI引導程序-超出牌張數-牌紀錄暫時變數(1~10.牌編號)
Public personatkingtfr(1 To 5) As Integer '計算個別技能-是否為Ex技(1~4.(1)有/(2)無,5.是否有封印)
Sub 智慧型AI系統計算_一階段_初始(ByVal pagenumber As Integer)
Erase cardcountAInum
Erase cardAInumnm
Erase cardAInumcase
Erase cardAInumselect3
cardAInumans = ""
cardAInumselect1 = 0
cardAInumselect4 = 0
cardAInumselect2 = ""
cardAInumchoose = 0
cardAInumuscom = pagenumber
cardAITotalNUM = 2 ^ cardAInumuscom
ReDim cardcountAInum(1 To cardAInumuscom, 1 To 5) As String
ReDim cardAInumcaseperson(1 To cardAITotalNUM, 1 To 2, 1 To 15) As Integer
ReDim cardAInumcasepersonTER(1 To cardAITotalNUM, 1 To 5, 1 To 10) As Integer
ReDim cardAInumFinal(1 To cardAITotalNUM, 1 To 4) As Integer
ReDim cardAInumFinal2(1 To cardAITotalNUM, 1 To 4) As Integer
'=========計算正反面排列組合數值
智慧型AI系統類.排列組合計算 pagenumber
End Sub
Sub 智慧型AI系統計算_一階段_取得牌面資料(ByVal 是否一般 As Boolean, ByVal uscom As Integer)
If 是否一般 = True Then
        '=========擷取目前牌面資料
        Select Case uscom
            Case 1
                戰鬥系統類.出牌順序計算_使用者_手牌
            Case 2
                戰鬥系統類.出牌順序計算_電腦_手牌
        End Select
        Dim w As Integer '暫時變數
        w = 2 * uscom '(2-使用者手牌/4-電腦手牌)
        For i = 1 To pageglead(uscom)
            cardcountAInum(i, 5) = 出牌順序統計暫時變數(w, i, 2)
            cardcountAInum(i, 1) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 1)
            cardcountAInum(i, 2) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 2)
            cardcountAInum(i, 3) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 3)
            cardcountAInum(i, 4) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 4)
        Next
End If
'======================
'智慧型AI系統類.排列組合統計數值計算_手牌總計
智慧型AI系統類.排列組合統計數值計算_部分相似法去除重複組合
智慧型AI系統類.排列組合統計數值計算_個別組合
End Sub
Sub 智慧型AI系統計算_二階段_計算期望值_初始(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
Dim wnum As Integer, whnum As Integer '暫時變數
Select Case turn
    Case 1 '===攻擊階段
         If uscom = 1 Then whnum = atkus(角色人物對戰人數(1, 2)) Else whnum = atkcom(角色人物對戰人數(2, 2))
         '==========================
         For i = 0 To (cardAITotalNUM) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If (cardcountAInum(j, 1) = a1a And movecpre = 1) Or (cardcountAInum(j, 1) = a5a And movecpre > 1) Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                      Case 1
                          If (cardcountAInum(j, 3) = a1a And movecpre = 1) Or (cardcountAInum(j, 3) = a5a And movecpre > 1) Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
             If Val(wnum) > 0 Then
                 cardAInumFinal(i + 1, 1) = Val(cardAInumFinal(i + 1, 1)) + Val(whnum)
             End If
         Next
    Case 2  '===防禦階段
         For i = 0 To (cardAITotalNUM) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If cardcountAInum(j, 1) = a2a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                      Case 1
                          If cardcountAInum(j, 3) = a2a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
         Next
    Case 3  '===移動階段
         For i = 0 To (cardAITotalNUM) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If cardcountAInum(j, 1) = a3a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                          End If
                      Case 1
                          If cardcountAInum(j, 3) = a3a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
         Next
End Select
End Sub
Sub 智慧型AI系統計算_二階段_計算期望值_個別技能(ByVal name As String, ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
'智慧型AI系統類.檢查人物技能是否有EX技 uscom, name
'If personatkingtfr(5) = 1 Then
'   Exit Sub '有封印狀態時無法發動技能
'End If
'Select Case name
'     Case "艾伯李斯特"
'           智慧型AI人物類.艾伯李斯特 turn, movecpre, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "雪莉"
'           智慧型AI人物類.雪莉 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "艾茵"
'           智慧型AI人物類.艾茵 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "古魯瓦爾多"
'           智慧型AI人物類.古魯瓦爾多 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "帕茉"
'           智慧型AI人物類.帕茉 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "史塔夏"
'           智慧型AI人物類.史塔夏 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "C.C."
'           智慧型AI人物類.CC turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "伊芙琳"
'           智慧型AI人物類.伊芙琳 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "布勞"
'           智慧型AI人物類.布勞 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "梅倫"
'           智慧型AI人物類.梅倫 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "音音夢"
'           智慧型AI人物類.音音夢 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "艾依查庫"
'           智慧型AI人物類.艾依查庫 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "阿貝爾"
'           智慧型AI人物類.阿貝爾 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "利恩"
'           智慧型AI人物類.利恩 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "夏洛特"
'           智慧型AI人物類.夏洛特 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "泰瑞爾"
'           智慧型AI人物類.泰瑞爾 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "瑪格莉特"
'           智慧型AI人物類.瑪格莉特 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "庫勒尼西"
'           智慧型AI人物類.庫勒尼西 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "蕾格烈芙"
'           智慧型AI人物類.蕾格烈芙 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "多妮妲"
'           智慧型AI人物類.多妮妲 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "傑多"
'           智慧型AI人物類.傑多 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "阿奇波爾多"
'           智慧型AI人物類.阿奇波爾多 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "露緹亞"
'           智慧型AI人物類.露緹亞 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "梅莉"
'           智慧型AI人物類.梅莉 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "貝琳達"
'           智慧型AI人物類.貝琳達 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "蕾"
'           智慧型AI人物類.蕾 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "羅莎琳"
'           智慧型AI人物類.羅莎琳 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "洛洛妮"
'           智慧型AI人物類.洛洛妮 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "克頓"
'           智慧型AI人物類.克頓 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "艾蕾可"
'           智慧型AI人物類.艾蕾可 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "尤莉卡"
'           智慧型AI人物類.尤莉卡 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'End Select
'===========================執行階段插入點(99)
智慧型AI系統類.智慧型AI系統_執行階段99_主動技能執行 uscom, turn, movecpre
'============================
End Sub
Sub 排列組合計算(ByVal qnum As Integer)
'===========
ReDim cardAIn(1 To Val(qnum))
Erase cardAInumnm
cardAInumans = ""
Dim s As Integer
For i = 1 To qnum   '重設區塊數值
    cardAIn(i) = 0
Next
s = 1
'================
Do
    For i = qnum To 1 Step -1
        cardAInumans = cardAInumans & cardAIn(i)
    Next
    '================
    cardAIn(1) = cardAIn(1) + 1
    智慧型AI系統類.排列組合計算_區塊進位 qnum '共[qnum]位數
    '================
    s = s + 1
    cardAInumans = cardAInumans & "="
Loop Until s > (2 ^ qnum)
cardAInumnm = Split(cardAInumans, "=")

End Sub
Sub 排列組合計算_區塊進位(ByVal num As Integer)
For i = 1 To num - 1
    If cardAIn(i) = 2 Then
        cardAIn(i + 1) = cardAIn(i + 1) + 1
        cardAIn(i) = 0
    End If
Next

End Sub
Sub 排列組合統計數值計算_手牌總計()
Dim we As Integer  '暫時變數
For i = 1 To cardAInumuscom
    For j = 1 To 2
        we = 2 * j
        Select Case cardcountAInum(i, j)
             Case a1a
                  If cardcountAInum(i, we) < cardAInumcase(1, 1) Or cardAInumcase(1, 1) = 0 Then
                      cardAInumcase(1, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(1, 2) Or cardAInumcase(1, 2) = 0 Then
                      cardAInumcase(1, 2) = cardcountAInum(i, we)
                  End If
             Case a2a
                  If cardcountAInum(i, we) < cardAInumcase(2, 1) Or cardAInumcase(2, 1) = 0 Then
                      cardAInumcase(2, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(2, 2) Or cardAInumcase(2, 2) = 0 Then
                      cardAInumcase(2, 2) = cardcountAInum(i, we)
                  End If
             Case a3a
                  If cardcountAInum(i, we) < cardAInumcase(3, 1) Or cardAInumcase(3, 1) = 0 Then
                      cardAInumcase(3, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(3, 2) Or cardAInumcase(3, 2) = 0 Then
                      cardAInumcase(3, 2) = cardcountAInum(i, we)
                  End If
             Case a4a
                  If cardcountAInum(i, we) < cardAInumcase(4, 1) Or cardAInumcase(4, 1) = 0 Then
                      cardAInumcase(4, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(4, 2) Or cardAInumcase(4, 2) = 0 Then
                      cardAInumcase(4, 2) = cardcountAInum(i, we)
                  End If
             Case a5a
                  If cardcountAInum(i, we) < cardAInumcase(5, 1) Or cardAInumcase(5, 1) = 0 Then
                      cardAInumcase(5, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(5, 2) Or cardAInumcase(5, 2) = 0 Then
                      cardAInumcase(5, 2) = cardcountAInum(i, we)
                  End If
        End Select
    Next
Next
End Sub
Sub 排列組合統計數值計算_個別組合()
Dim we As Integer '暫時變數
For i = 1 To cardAITotalNUM
    For j = 1 To cardAInumuscom
        Select Case Mid(cardAInumnm(i - 1), j, 1)
            Case 0
                 we = 2
                  Select Case cardcountAInum(j, 1)
                     Case a1a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 1) Or cardAInumcaseperson(i, 1, 1) = 0 Then
                              cardAInumcaseperson(i, 1, 1) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 2) Or cardAInumcaseperson(i, 1, 2) = 0 Then
                              cardAInumcaseperson(i, 1, 2) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 11) = cardAInumcaseperson(i, 1, 11) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 1, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 1, cardcountAInum(j, we))) + 1
                     Case a2a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 3) Or cardAInumcaseperson(i, 1, 3) = 0 Then
                              cardAInumcaseperson(i, 1, 3) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 4) Or cardAInumcaseperson(i, 1, 4) = 0 Then
                              cardAInumcaseperson(i, 1, 4) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 12) = cardAInumcaseperson(i, 1, 12) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 2, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 2, cardcountAInum(j, we))) + 1
                     Case a3a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 5) Or cardAInumcaseperson(i, 1, 5) = 0 Then
                              cardAInumcaseperson(i, 1, 5) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 6) Or cardAInumcaseperson(i, 1, 6) = 0 Then
                              cardAInumcaseperson(i, 1, 6) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 13) = cardAInumcaseperson(i, 1, 13) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 3, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 3, cardcountAInum(j, we))) + 1
                     Case a4a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 7) Or cardAInumcaseperson(i, 1, 7) = 0 Then
                              cardAInumcaseperson(i, 1, 7) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 8) Or cardAInumcaseperson(i, 1, 8) = 0 Then
                              cardAInumcaseperson(i, 1, 8) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 14) = cardAInumcaseperson(i, 1, 14) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 4, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 4, cardcountAInum(j, we))) + 1
                     Case a5a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 9) Or cardAInumcaseperson(i, 1, 9) = 0 Then
                              cardAInumcaseperson(i, 1, 9) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 10) Or cardAInumcaseperson(i, 1, 10) = 0 Then
                              cardAInumcaseperson(i, 1, 10) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 15) = cardAInumcaseperson(i, 1, 15) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 5, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 5, cardcountAInum(j, we))) + 1
                End Select
            Case 1
                 we = 4
                  Select Case cardcountAInum(j, 3)
                     Case a1a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 1) Or cardAInumcaseperson(i, 1, 1) = 0 Then
                              cardAInumcaseperson(i, 1, 1) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 2) Or cardAInumcaseperson(i, 1, 2) = 0 Then
                              cardAInumcaseperson(i, 1, 2) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 11) = cardAInumcaseperson(i, 1, 11) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 1, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 1, cardcountAInum(j, we))) + 1
                     Case a2a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 3) Or cardAInumcaseperson(i, 1, 3) = 0 Then
                              cardAInumcaseperson(i, 1, 3) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 4) Or cardAInumcaseperson(i, 1, 4) = 0 Then
                              cardAInumcaseperson(i, 1, 4) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 12) = cardAInumcaseperson(i, 1, 12) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 2, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 2, cardcountAInum(j, we))) + 1
                     Case a3a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 5) Or cardAInumcaseperson(i, 1, 5) = 0 Then
                              cardAInumcaseperson(i, 1, 5) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 6) Or cardAInumcaseperson(i, 1, 6) = 0 Then
                              cardAInumcaseperson(i, 1, 6) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 13) = cardAInumcaseperson(i, 1, 13) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 3, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 3, cardcountAInum(j, we))) + 1
                     Case a4a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 7) Or cardAInumcaseperson(i, 1, 7) = 0 Then
                              cardAInumcaseperson(i, 1, 7) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 8) Or cardAInumcaseperson(i, 1, 8) = 0 Then
                              cardAInumcaseperson(i, 1, 8) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 14) = cardAInumcaseperson(i, 1, 14) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 4, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 4, cardcountAInum(j, we))) + 1
                     Case a5a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 9) Or cardAInumcaseperson(i, 1, 9) = 0 Then
                              cardAInumcaseperson(i, 1, 9) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 10) Or cardAInumcaseperson(i, 1, 10) = 0 Then
                              cardAInumcaseperson(i, 1, 10) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 15) = cardAInumcaseperson(i, 1, 15) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 5, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 5, cardcountAInum(j, we))) + 1
                End Select
        End Select
    Next
Next
End Sub
Sub 排列組合統計數值計算_部分相似法去除重複組合()
Dim 卡片已重複標記() As Boolean
Dim 卡片相似標記() As Boolean
Dim 卡片排列組合複製() As String
Dim CardChooseNum As Integer, 卡片相似標記數 As Integer
Dim c1 As String, c2 As String
ReDim 卡片已重複標記(1 To cardAITotalNUM) As Boolean
ReDim 卡片排列組合複製(cardAITotalNUM - 1) As String
'ReDim 卡片相似標記(1 To cardAInumuscom) As Boolean
'=========================================
For i = 1 To cardAITotalNUM
    For j = i - 1 To 1 Step -1
        卡片相似標記數 = 0
        ReDim 卡片相似標記(1 To cardAInumuscom) As Boolean
'        If 卡片已重複標記(j) = False Then
            For k = 1 To cardAInumuscom
                c1 = Mid(cardAInumnm(i - 1), k, 1)
                For p = 1 To cardAInumuscom
                     c2 = Mid(cardAInumnm(j - 1), p, 1)
                     If c1 = "0" And c2 = "0" Then
                         If cardcountAInum(k, 1) = cardcountAInum(p, 1) And cardcountAInum(k, 2) = cardcountAInum(p, 2) And 卡片相似標記(p) = False Then
                             卡片相似標記(p) = True
                             卡片相似標記數 = 卡片相似標記數 + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     ElseIf c1 = "0" And c2 = "1" Then
                         If cardcountAInum(k, 1) = cardcountAInum(p, 3) And cardcountAInum(k, 2) = cardcountAInum(p, 4) And 卡片相似標記(p) = False Then
                             卡片相似標記(p) = True
                             卡片相似標記數 = 卡片相似標記數 + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     ElseIf c1 = "1" And c2 = "0" Then
                         If cardcountAInum(k, 3) = cardcountAInum(p, 1) And cardcountAInum(k, 4) = cardcountAInum(p, 2) And 卡片相似標記(p) = False Then
                             卡片相似標記(p) = True
                             卡片相似標記數 = 卡片相似標記數 + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     ElseIf c1 = "1" And c2 = "1" Then
                         If cardcountAInum(k, 3) = cardcountAInum(p, 3) And cardcountAInum(k, 4) = cardcountAInum(p, 4) And 卡片相似標記(p) = False Then
                             卡片相似標記(p) = True
                             卡片相似標記數 = 卡片相似標記數 + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     End If
                Next
            Next
'        End If
        If 卡片相似標記數 = cardAInumuscom Then
            卡片已重複標記(i) = True
            CardChooseNum = CardChooseNum + 1
            j = 1 'Exit For
        End If
    Next
Next
'MsgBox "部分相似重複組合數:" & CardChooseNum
'============================
If CardChooseNum > 0 Then
    '====================
    For i = 0 To cardAITotalNUM - 1
        卡片排列組合複製(i) = cardAInumnm(i)
    Next
    ReDim cardAInumnm((cardAITotalNUM - 1) - CardChooseNum) As String
    '====================
    k = 0
    For i = 0 To cardAITotalNUM - 1
        If 卡片已重複標記(i + 1) = False Then
            cardAInumnm(k) = 卡片排列組合複製(i)
            k = k + 1
        End If
    Next
    cardAITotalNUM = cardAITotalNUM - CardChooseNum
    '=====================
    ReDim cardAInumcaseperson(1 To cardAITotalNUM, 1 To 2, 1 To 15) As Integer
    ReDim cardAInumcasepersonTER(1 To cardAITotalNUM, 1 To 5, 1 To 10) As Integer
    ReDim cardAInumFinal(1 To cardAITotalNUM, 1 To 4) As Integer
    ReDim cardAInumFinal2(1 To cardAITotalNUM, 1 To 4) As Integer
    '=====================
End If
End Sub
Sub 智慧型AI系統計算_三階段_統計排列()
'=================複製內容
For k = 1 To cardAITotalNUM
    cardAInumFinal2(k, 1) = cardAInumFinal(k, 1)
    cardAInumFinal2(k, 2) = cardAInumFinal(k, 2)
Next
'=================
Dim wer As Integer, wes As Integer
For i = cardAITotalNUM To 1 Step -1
    For j = 1 To i - 1
        If Val(cardAInumFinal2(j, 1)) < Val(cardAInumFinal2(j + 1, 1)) Then
            wer = cardAInumFinal2(j + 1, 1)
            wes = cardAInumFinal2(j + 1, 2)
            cardAInumFinal2(j + 1, 1) = cardAInumFinal2(j, 1)
            cardAInumFinal2(j + 1, 2) = cardAInumFinal2(j, 2)
            cardAInumFinal2(j, 1) = wer
            cardAInumFinal2(j, 2) = wes
        End If
    Next
Next
End Sub
Sub 智慧型AI系統計算_四階段_比序_1_初始()
For i = 1 To cardAITotalNUM
    If Val(cardAInumFinal2(i, 1)) > Val(cardAInumselect1) Then
        cardAInumselect1 = cardAInumFinal2(i, 1)
    End If
Next
'====================
If cardAInumselect1 < 0 Then cardAInumselect1 = 0 '去除總期望值為負數之組合
'====================
For i = 1 To cardAITotalNUM
    If cardAInumFinal2(i, 1) = cardAInumselect1 Then
        cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
    End If
Next
'====================
If cardAInumselect2 = "" Then  '沒有任何組合符合條件
    cardAInumselect2 = "-10=-10"
End If
End Sub
Sub 智慧型AI系統計算_四階段_比序_2_超額比序判斷_1()
'cardAInumselect3 = Split(cardAInumselect2, "=")
'If UBound(cardAInumselect3) > 1 Then
'    For i = 1 To cardAITotalNUM
'        For j = 1 To cardAInumuscom
'             If cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, j) < 0 Then
'                 cardAInumFinal2(i, 3) = 1
'             End If
'             cardAInumFinal2(i, 4) = Val(cardAInumFinal2(i, 4)) + Val(cardAInumcaseperson(cardAInumFinal2(i, 2), 2, j))
'        Next
'    Next
'    '===============
'    Erase cardAInumselect3
'    cardAInumselect2 = ""
'    '======
'    For i = 1 To cardAITotalNUM
'        If cardAInumFinal2(i, 1) = cardAInumselect1 And cardAInumFinal2(i, 3) = 0 Then
'            cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
'        End If
'    Next
'    cardAInumselect3 = Split(cardAInumselect2, "=")
'End If
End Sub
Sub 智慧型AI系統計算_四階段_比序_2_超額比序判斷_2()
cardAInumselect3 = Split(cardAInumselect2, "=")
If UBound(cardAInumselect3) > 1 Then
    Dim wer As Integer
    wer = cardAInumFinal2(1, 4)  '目標選取最高牌張數，總期望值最高之組合
    '===============
    For i = 1 To cardAITotalNUM
         If Val(cardAInumFinal2(i, 4)) > Val(wer) And cardAInumFinal2(i, 1) = cardAInumselect1 Then
             wer = cardAInumFinal2(i, 4)
         End If
    Next
    '===============
    Erase cardAInumselect3
    cardAInumselect2 = ""
    cardAInumselect4 = wer
    '======
    For i = 1 To cardAITotalNUM
        If cardAInumFinal2(i, 4) = wer And cardAInumFinal2(i, 1) = cardAInumselect1 Then
            cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
        End If
    Next
    cardAInumselect3 = Split(cardAInumselect2, "=")
End If
End Sub
Sub 智慧型AI系統計算_四階段_比序_3_選擇組合()
If UBound(cardAInumselect3) > 1 Then
    Dim wtr As Integer '暫時變數
    wtr = Int(Rnd() * UBound(cardAInumselect3)) + 1
    cardAInumchoose = cardAInumselect3(wtr)
Else
    cardAInumchoose = cardAInumselect3(1)
End If
End Sub
Sub 智慧型AI系統計算_最後階段_實行選牌(ByVal choose As Integer, ByVal uscom As Integer)
Dim wer As Integer '暫時變數
If choose = 1 Then
    wer = 0
Else
    wer = 1
End If
'=================
Dim pu As Integer '暫時變數
'=====
If cardAInumchoose = -10 Then  '==沒有任何組合符合出牌條件
    Exit Sub
End If
'=======================如組合符合出牌條件的話
Select Case uscom
     Case 1 '==使用者方
            For i = 1 To UBound(cardcountAInum, 1)
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 4
                    ElseIf cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 3
                    End If
            Next
     Case 2 '==電腦方
            For i = 1 To UBound(cardcountAInum, 1)
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        cspce = pagecardnum(pu, 1)
                        cspme = pagecardnum(pu, 2)
                        pagecardnum(pu, 1) = pagecardnum(pu, 3)
                        pagecardnum(pu, 2) = pagecardnum(pu, 4)
                        pagecardnum(pu, 3) = cspce
                        pagecardnum(pu, 4) = cspme
                        If pageonin(pu) = 2 Then
                           pageonin(pu) = 1
                        Else
                           pageonin(pu) = 2
                        End If
                    End If
                    If cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 1
                    End If
            Next
End Select
End Sub
Sub 智慧型AI系統計算_暫時匯出(ByVal uscom As Integer)
'If Formsetting.checktest.Value = 1 Then
''    Open App.Path & "\test\out1.txt" For Output As #1
'    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & BattleTurn & "turn_" & 戰鬥系統類.turnatk & "_" & uscom & "_1.txt" For Output As #1
'    For i = 1 To cardAITotalNUM
'        Print #1, cardAInumnm(Val(cardAInumFinal2(i, 2)) - 1) & "=" & cardAInumFinal2(i, 1) & "/" & cardAInumFinal2(i, 4) & "#" & cardAInumFinal2(i, 2) & "@";
'        For k = 1 To cardAInumuscom
'            Print #1, cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, k) & "=";
'        Next
'        Print #1,
'    Next
'    Close
'    'MsgBox "已匯出完畢1"
'End If
End Sub
Sub 智慧型AI系統計算_引導程序_試驗1(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer)
智慧型AI系統類.智慧型AI系統計算_一階段_初始 uscom
智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turn, movecpre, uscom
智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turn, movecpre, uscom
智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
智慧型AI系統類.智慧型AI系統計算_暫時匯出 uscom
End Sub
Sub 智慧型AI系統計算_引導程序_選擇(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer)
Dim CardMaxNum As Integer
If Formsetting.chksetcomaipagenum.Value = 1 Then
    CardMaxNum = Val(Formsetting.自訂AI手牌張數.Text)
Else
    CardMaxNum = 7
End If
'========================
If Val(pageglead(uscom)) > CardMaxNum Then
    智慧型AI系統類.智慧型AI系統計算_引導程序_超出牌張數 uscom, turn, name, movecpre, choose, CardMaxNum
ElseIf Val(pageglead(uscom)) > 0 And Val(pageglead(uscom)) <= CardMaxNum Then
    智慧型AI系統類.智慧型AI系統計算_一階段_初始 pageglead(uscom)
    智慧型AI系統類.智慧型AI系統計算_一階段_取得牌面資料 True, uscom
    智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turn, movecpre, uscom
    智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turn, movecpre, uscom
    智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_1_初始
'    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_1
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_2
    智慧型AI系統類.智慧型AI系統計算_暫時匯出 uscom
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_3_選擇組合
    If turn = 3 And cardAInumchoose > 0 Then
        智慧型AI系統類.智慧型AI系統計算_引導程序_移動階段續 uscom, turn, name, movecpre, choose, pageglead(uscom)
    Else
        智慧型AI系統類.智慧型AI系統計算_最後階段_實行選牌 choose, uscom
    End If
End If
End Sub
Sub 智慧型AI系統計算_引導程序_移動階段續(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer, ByVal pagenumber As Integer)
If Val(pagenumber) > 0 Then
    Select Case 智慧型AI系統計算_移動階段續_判斷出牌資格(uscom)
        Case True
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_一階段_準備進行資料
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_二階段_進行估計排列組合串計算 pagenumber, uscom
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_三階段_進行估計期望值計算 uscom, name, choose, movecpre, pagenumber
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_四階段_統計估計期望值及判斷 uscom
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_五階段_實行選牌 choose, uscom, pagenumber
        Case False
'            智慧型AI系統類.智慧型AI系統計算_移動階段續_否定面_一階段_重設期望值_個別
            智慧型AI系統類.智慧型AI系統計算_移動階段續_否定面_二階段_選擇行動 uscom
            智慧型AI系統類.智慧型AI系統計算_最後階段_實行選牌 choose, uscom
    End Select
End If
End Sub
Function 智慧型AI系統_目前可執行之人物判斷(ByVal name As String) As Boolean
If Formsetting.chkusenewai.Value = 0 Then
    智慧型AI系統_目前可執行之人物判斷 = False
    Exit Function
End If
Select Case name
    Case "艾伯李斯特"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "雪莉"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "艾茵"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "古魯瓦爾多"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "帕茉"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "史塔夏"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "C.C."
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "伊芙琳"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "布勞"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "梅倫"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "音音夢"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "艾依查庫"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "阿貝爾"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "利恩"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "夏洛特"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "泰瑞爾"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "瑪格莉特"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "庫勒尼西"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "蕾格烈芙"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "多妮妲"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "傑多"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "阿奇波爾多"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "露緹亞"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "梅莉"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "貝琳達"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "蕾"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "羅莎琳"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "洛洛妮"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "克頓"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "艾蕾可"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "尤莉卡"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case Else
            智慧型AI系統_目前可執行之人物判斷 = False
End Select
End Function
Function 階層數(ByVal num As Integer) As Single
Dim w As Double
w = 1
If num <> 0 Then
    For i = 1 To Val(num)
        w = Val(w) * Val(i)
    Next
Else
    w = 1
End If
階層數 = w
End Function
Function 階層數_取C(ByVal c1 As Integer, ByVal c2 As Integer) As Single

階層數_取C = (智慧型AI系統類.階層數(c1) / 智慧型AI系統類.階層數(Val(c1) - Val(c2))) / 智慧型AI系統類.階層數(c2)

End Function
Sub 智慧型AI系統計算_移動階段續_取得計算之排列組合(ByVal n1 As Integer, ByVal n2 As Integer)
Dim wtstr As String, wtall As Integer, wtpnum() As String, wtn As Integer
'===================
智慧型AI系統類.排列組合計算 n1
wtall = 智慧型AI系統類.階層數_取C(n1, n2)
ReDim cardAInumMOVnm(1 To wtall) As String
'====================
For i = 1 To 2 ^ n1
    wtn = 0
    For j = 1 To n1
        If Val(Mid(cardAInumnm(i - 1), j, 1)) = 1 Then
            wtn = wtn + 1
        End If
    Next
    If wtn = n2 Then '==有n2張出牌之組合
        wtstr = wtstr & "=" & i
    End If
Next
wtpnum = Split(wtstr, "=")
'If UBound(wtpnum) = wtall Then
'    MsgBox wtstr
'    For i = 1 To UBound(wtpnum)
'        Debug.Print wtpnum(i) & "=" & cardAInumnm(wtpnum(i) - 1)
'    Next
'Else
'    MsgBox "失敗"
'End If
For i = 1 To UBound(wtpnum)
    cardAInumMOVnm(i) = cardAInumnm(wtpnum(i) - 1)
Next
End Sub
Function 智慧型AI系統計算_移動階段續_判斷出牌資格(ByVal uscom As Integer) As Boolean
Erase cardAInumMOVmain
Erase cardAInumMOVnm
Erase cardAInumMOVnmtot
Dim wtmovnum As Integer '暫時變數
Dim buffobj As clsStatus
If cardAInumchoose = -10 Then
    智慧型AI系統計算_移動階段續_判斷出牌資格 = False
    Exit Function
End If
'============紀錄目前組合
cardAInumMOVmain(1, 1) = cardAInumselect1
cardAInumMOVmain(1, 2) = cardAInumselect4
cardAInumMOVmain(1, 3) = cardAInumnm(cardAInumchoose - 1)
cardAInumMOVmain(1, 4) = cardAInumcaseperson(cardAInumchoose, 1, 13)
cardAInumMOVmain(1, 5) = cardAInumchoose
For i = 1 To cardAInumuscom
    cardAInumMOVmain(2, i) = cardAInumcaseperson(cardAInumchoose, 2, i)
Next
'==============計算有效移動數
wtmovnum = cardAInumMOVmain(1, 4)
For Each n In 人物異常狀態列表(uscom, 角色人物對戰人數(uscom, 2))
    Set buffobj = n
    If buffobj.Identifier = "BUFFN00302" Then
        wtmovnum = Val(wtmovnum) - buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00301" Then
        wtmovnum = Val(wtmovnum) + buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00801" Then
        wtmovnum = -100
    ElseIf buffobj.Identifier = "BUFFN00501" And _
        ((uscom = 1 And liveus(角色人物對戰人數(uscom, 2)) = 1) Or (uscom = 2 And livecom(角色人物對戰人數(uscom, 2)) = 1)) Then
        wtmovnum = -100
    End If
Next
'=====================
If wtmovnum >= 2 Then
    智慧型AI系統計算_移動階段續_判斷出牌資格 = True
Else
    智慧型AI系統計算_移動階段續_判斷出牌資格 = False
End If
End Function
Sub 智慧型AI系統計算_移動階段續_否定面_一階段_重設期望值_個別()
'For i = 1 To cardAInumuscom
'    If cardAInumcaseperson(cardAInumchoose, 2, i) < 10 Then
'        cardAInumcaseperson(cardAInumchoose, 2, i) = 0
'        cardAInumMOVmain(2, i) = 0
'    End If
'Next
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_一階段_準備進行資料()
Dim wercnum As Integer, werct As String, werpnum As Integer
ReDim cardcountAInumMOV(1 To cardAInumuscom, 1 To 5) As String
是否移動階段續估計判斷程序 = True
For k = 1 To cardAInumuscom
    Select Case Mid(cardAInumMOVmain(1, 3), k, 1)
         Case 0
              If cardcountAInum(k, 1) = a3a And cardAInumMOVmain(2, k) = 0 Then
                  wercnum = Val(wercnum) + 1
                  werct = werct & "=" & k
'              ElseIf cardAInumMOVmain(2, k) >= 10 Then
'                 werpnum = Val(werpnum) + 1
              End If
         Case 1
              If cardcountAInum(k, 3) = a3a And cardAInumMOVmain(2, k) = 0 Then
                  wercnum = Val(wercnum) + 1
                  werct = werct & "=" & k
'              ElseIf cardAInumMOVmain(2, k) >= 10 Then
'                 werpnum = Val(werpnum) + 1
              End If
    End Select
    For q = 1 To 5
         cardcountAInumMOV(k, q) = cardcountAInum(k, q)
    Next
Next
'===============
'If Val(werpnum) >= 1 Then werpnum = 1
'===============
ReDim cardAInumMOVnmtot(0 To (2 ^ wercnum), 1 To 8) As String
cardAInumMOVnmtot(0, 1) = werct
cardAInumMOVnmtot(0, 2) = 1
cardAInumMOVnmtot(0, 3) = wercnum
cardAInumMOVnmtot(0, 4) = 1
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_二階段_進行估計排列組合串計算(ByVal pagenumber As Integer, ByVal uscom As Integer)
Dim weru As Integer, wernum As Integer, werqr As String
Dim werstru As String
Dim werpstr() As String
Dim wermovnm As Integer, wermovynm As Integer
'============進行估計之移動牌排列組合計算
For i = 1 To Val(cardAInumMOVnmtot(0, 3))
       智慧型AI系統計算_移動階段續_取得計算之排列組合 Val(cardAInumMOVnmtot(0, 3)), i
       weru = 1
       wernum = 階層數_取C(Val(cardAInumMOVnmtot(0, 3)), i)
        For k = Val(cardAInumMOVnmtot(0, 2)) To (Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)) - 1
             cardAInumMOVnmtot(k, 1) = cardAInumMOVnm(weru)
             weru = Val(weru) + 1
        Next
        cardAInumMOVnmtot(0, 2) = Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)
Next
'=====================進行剩餘移動牌之排列組合串整合
'werpstr = Split(cardAInumMOVnmtot(1, 1), "=")
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    weru = 0
    werstru = ""
    wermovnm = 0
    wermovynm = 0
    For k = 1 To pagenumber
        Select Case Mid(cardAInumMOVmain(1, 3), k, 1)
              Case 0
                    If cardcountAInum(k, 1) = a3a And i <= (2 ^ Val(cardAInumMOVnmtot(0, 3)) - 1) And cardAInumMOVmain(2, k) = 0 Then
                        weru = Val(weru) + 1
                        If Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 1 Then
                            werstru = werstru & "1"
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
                            wermovynm = Val(wermovynm) + 1
                        ElseIf Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 0 Then
'                            werstru = werstru & "1"
'                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
'                            wermovynm = Val(wermovynm) + 1
'                        Else
                            werstru = werstru & "n"
                        End If
                    ElseIf cardAInumMOVmain(2, k) = 1 Then
                        werstru = werstru & "1"
                        If cardcountAInum(k, 1) = a3a Then
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
                            wermovynm = Val(wermovynm) + 1
                        End If
                    Else
                        werstru = werstru & "n"
                    End If
              Case 1
                    If cardcountAInum(k, 3) = a3a And i <= (2 ^ Val(cardAInumMOVnmtot(0, 3)) - 1) And cardAInumMOVmain(2, k) = 0 Then
                        weru = Val(weru) + 1
                        If Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 1 Then
                            werstru = werstru & "1"
                            wermovynm = Val(wermovynm) + 1
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
                        ElseIf Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 0 Then
'                            werstru = werstru & "1"
'                            wermovynm = Val(wermovynm) + 1
'                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
'                        Else
                            werstru = werstru & "n"
                        End If
                    ElseIf cardAInumMOVmain(2, k) = 1 Then
                        werstru = werstru & "1"
                        If cardcountAInum(k, 3) = a3a Then
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
                            wermovynm = Val(wermovynm) + 1
                        End If
                    Else
                        werstru = werstru & "n"
                    End If
        End Select
    Next
    cardAInumMOVnmtot(i, 2) = werstru
    cardAInumMOVnmtot(i, 6) = wermovnm
    cardAInumMOVnmtot(i, 7) = wermovynm
Next
'=========================測試用匯出
'If Formsetting.checktest.Value = 1 Then
''    Open App.Path & "\test\out2.txt" For Output As #1
'    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & BattleTurn & "turn_" & 戰鬥系統類.turnatk & "_" & uscom & "_2.txt" For Output As #1
'    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
'        Print #1, cardAInumMOVnmtot(i, 2)
'    Next
'    Print #1, cardAInumMOVmain(1, 5) & "=" & cardAInumMOVmain(1, 3)
'    Close
'    'MsgBox "已匯出完畢2"
'End If
'==============================
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_三階段_進行估計期望值計算(ByVal uscom As Integer, ByVal name As String, ByVal choose As Integer, ByVal movecpre As Integer, ByVal pagenumber As Integer)
Dim weru As Integer, wertp As Integer, movecpren As Integer, turnm As Integer, werucount As Boolean
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    For k = 1 To 2
         '===========將資料轉移至待運算資料
         weru = 0
         For wp = 1 To pagenumber
              If Mid(cardAInumMOVnmtot(i, 2), wp, 1) = "n" Then
                  weru = Val(weru) + 1
              End If
         Next
         If Val(weru) > 0 Then
                 智慧型AI系統類.智慧型AI系統計算_一階段_初始 weru
                 wertp = 0
                 '=======
                 For q = 1 To pagenumber
                     If Mid(cardAInumMOVnmtot(i, 2), q, 1) = "n" Then
                           wertp = Val(wertp) + 1
                           For wds = 1 To 5
                                 cardcountAInum(wertp, wds) = cardcountAInumMOV(q, wds)
                           Next
                    End If
                Next
                '========================
                If k = 1 Then movecpren = 1 Else movecpren = 3
                If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) And werucount = True Then
                    turnm = 2
                    movecpren = movecpre
                Else
                    turnm = 1
                End If
                '========================
                智慧型AI系統類.智慧型AI系統計算_一階段_取得牌面資料 False, uscom
                智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turnm, movecpren, uscom
                智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turnm, movecpren, uscom
                智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_1_初始
'                智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_1
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_2
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_3_選擇組合
        Else
                cardAInumselect1 = 0
        End If
        '=======================將重新估計後資料儲存
        If k = 1 And werucount = False Then
           movecpren = 3
        ElseIf k = 2 And werucount = False Then
           movecpren = 4
        ElseIf werucount = True Then
           movecpren = 8
        End If
        '=========
        cardAInumMOVnmtot(i, movecpren) = cardAInumselect1
        '=========
        If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) And k = 2 And werucount = False Then
           werucount = True
           k = 0
        ElseIf werucount = True Then
           k = 2
        End If
        '==========================
    Next
Next
'=========================測試用匯出
'If Formsetting.checktest.Value = 1 Then
''    Open App.Path & "\test\out3.txt" For Output As #1
'    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & BattleTurn & "turn_" & 戰鬥系統類.turnatk & "_" & uscom & "_3.txt" For Output As #1
'    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
'        Print #1, i & "=" & cardAInumMOVnmtot(i, 2) & "=";
'        For k = 3 To 4
'              Print #1, cardAInumMOVnmtot(i, k) & "#";
'        Next
'        If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) Then
'            Print #1, cardAInumMOVnmtot(i, 8);
'        End If
'        Print #1,
'    Next
'
'    Close
'    'MsgBox "已匯出完畢3"
'End If
'==============================
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_四階段_統計估計期望值及判斷(ByVal uscom As Integer)
Dim atk1max As Integer, atk2max As Integer, defmax As Integer, chemax As Integer, chestr As String
Dim wtmovnum As Integer
Dim buffobj As clsStatus
'==================篩選是否符合移動量
For Each n In 人物異常狀態列表(uscom, 角色人物對戰人數(uscom, 2))
    Set buffobj = n
    If buffobj.Identifier = "BUFFN00302" Then
        wtmovnum = Val(wtmovnum) - buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00301" Then
        wtmovnum = Val(wtmovnum) + buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00801" Then
        wtmovnum = -100
    End If
Next
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, 6)) + Val(wtmovnum) < 2 Then
         cardAInumMOVnmtot(i, 5) = "x"
     Else
         cardAInumMOVnmtot(i, 5) = "y"
     End If
Next
'===================
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, 3)) > Val(atk1max) And cardAInumMOVnmtot(i, 5) = "y" Then
         atk1max = cardAInumMOVnmtot(i, 3)
     End If
     If Val(cardAInumMOVnmtot(i, 4)) > Val(atk2max) And cardAInumMOVnmtot(i, 5) = "y" Then
         atk2max = cardAInumMOVnmtot(i, 4)
     End If
Next
defmax = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 8)
'==================
If Val(atk1max) >= Val(atk2max) And Val(atk1max) >= Val(defmax) Then
    chemax = 1
ElseIf Val(atk1max) <= Val(atk2max) And Val(atk2max) >= Val(defmax) Then
    chemax = 2
ElseIf Val(defmax) >= Val(atk1max) And Val(defmax) >= Val(atk2max) Then
    chemax = 3
Else
    chemax = 3
End If
'==================
Select Case chemax
     Case 1
           cardAInumMOVFinal(3) = 1
           cardAInumMOVFinal(2) = atk1max
           智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_確認實行_選擇最終組合 1, atk1max
     Case 2
           cardAInumMOVFinal(3) = 2
           cardAInumMOVFinal(2) = atk2max
           智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_確認實行_選擇最終組合 2, atk2max
     Case 3
           cardAInumMOVFinal(1) = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 2)
           cardAInumMOVFinal(3) = 3
           cardAInumMOVFinal(2) = defmax
End Select
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_確認實行_選擇最終組合(ByVal movche As Integer, ByVal atkmax As Integer)
Dim werstr As String, werg() As String, werg2() As String, werg3() As String
Dim werpagenum As Integer, werpgnumstr As String
Dim wermovmaxnum As Integer, wermvaxstr As String
Dim werrndnum As Integer, werche As Integer
'==========================
If movche = 1 Then werche = 3 Else werche = 4
'==========================
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, werche)) = Val(atkmax) Then
         werstr = werstr & "=" & i
     End If
Next
werg = Split(werstr, "=")
If UBound(werg) > 1 Then
        '====================================
        werpagenum = 0 '==目的取最大之出牌數
        For k = 1 To UBound(werg)
            If cardAInumMOVnmtot(werg(k), 7) > werpagenum Then
                werpagenum = cardAInumMOVnmtot(werg(k), 7)
            End If
        Next
        For k = 1 To UBound(werg)
            If cardAInumMOVnmtot(werg(k), 7) = werpagenum Then
                werpgnumstr = werpgnumstr & "=" & werg(k)
            End If
        Next
        werg2 = Split(werpgnumstr, "=")
        If UBound(werg2) > 1 Then
                '====================================
                wermovmaxnum = 0 '==目的取最大之移動數
                For k = 1 To UBound(werg2)
                    If Val(cardAInumMOVnmtot(werg(k), 6)) > Val(wermovmaxnum) Then
                        wermovmaxnum = cardAInumMOVnmtot(werg(k), 6)
                    End If
                Next
                For k = 1 To UBound(werg2)
                    If Val(cardAInumMOVnmtot(werg(k), 6)) = wermovmaxnum Then
                        wermvaxstr = wermvaxstr & "=" & werg2(k)
                    End If
                Next
                werg3 = Split(wermvaxstr, "=")
                If UBound(werg3) > 1 Then
                     Randomize
                     werrndnum = Int(Rnd() * UBound(werg3)) + 1
                     cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg3(werrndnum), 2)
                Else
                     cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg3(1), 2)
                End If
                '==========================================
        Else
                cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg2(1), 2)
        End If
        '====================================
Else
        cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg(1), 2)
End If
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_五階段_實行選牌(ByVal choose As Integer, ByVal uscom As Integer, ByVal pagenumber As Integer)
'Dim wer As Integer '暫時變數
'If choose = 1 Then
'    wer = 0
'Else
'    wer = 1
'End If
'=================
Dim pu As Integer '暫時變數
'=======================如組合符合出牌條件的話
Select Case uscom
     Case 1 '==使用者方
            For i = 1 To pagenumber
                    pu = cardcountAInumMOV(i, 5)
                    If Mid(cardAInumMOVFinal(1), i, 1) = 1 Then
'                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 And Val(cardAInumMOVmain(2, i)) >= wer Then
                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 Then
                                pagecardnum(pu, 11) = 4
'                            ElseIf Val(cardAInumMOVmain(2, i)) >= wer Then
                            Else
                                pagecardnum(pu, 11) = 3
                            End If
                    End If
            Next
            '===================選擇行動
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      目前數(33) = 3
                 Case 2
                      目前數(33) = 1
                 Case 3
                      目前數(33) = 2
            End Select
     Case 2 '==電腦方
            For i = 1 To pagenumber
                    pu = cardcountAInumMOV(i, 5)
                    If Mid(cardAInumMOVFinal(1), i, 1) = 1 Then
'                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 And Val(cardAInumMOVmain(2, i)) >= wer Then
                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 Then
                                cspce = pagecardnum(pu, 1)
                                cspme = pagecardnum(pu, 2)
                                pagecardnum(pu, 1) = pagecardnum(pu, 3)
                                pagecardnum(pu, 2) = pagecardnum(pu, 4)
                                pagecardnum(pu, 3) = cspce
                                pagecardnum(pu, 4) = cspme
                                If pageonin(pu) = 2 Then
                                   pageonin(pu) = 1
                                Else
                                   pageonin(pu) = 2
                                End If
                            End If
                            '==================
                            pagecardnum(pu, 11) = 1
                    End If
            Next
            '===================選擇行動
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      電腦方移動階段選擇數 = 3
                 Case 2
                      電腦方移動階段選擇數 = 1
                 Case 3
                      電腦方移動階段選擇數 = 2
            End Select
End Select

是否移動階段續估計判斷程序 = False
End Sub
Sub 智慧型AI系統計算_引導程序_超出牌張數(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer, ByVal CardNumMax As Integer)
If Val(pageglead(uscom)) > CardNumMax Then
    Dim CardOverCountNUM As Integer, CardNowNUM1 As Integer, CardNowNUM2 As Integer, CardNowCountNUM As Integer
    Dim w As Integer, k As Integer '暫時變數
    CardOverCountNUM = Int(Val(pageglead(uscom)) / Val(CardNumMax) + Val(0.9))
    CardNowNUM1 = 1: CardNowNUM2 = CardNumMax
    CardNowCountNUM = 0
    '==========================
    Do
        ReDim cardAInumOvertenrecord(1 To (CardNowNUM2 - CardNowNUM1 + 1)) As Integer
        智慧型AI系統類.智慧型AI系統計算_一階段_初始 (CardNowNUM2 - CardNowNUM1 + 1)
        '=========擷取目前牌面資料(前[CardNumMax]張)
            Select Case uscom
                Case 1
                    戰鬥系統類.出牌順序計算_使用者_手牌
                Case 2
                    戰鬥系統類.出牌順序計算_電腦_手牌
            End Select
            w = 2 * uscom '(2-使用者手牌/4-電腦手牌)
            k = 1
            For i = CardNowNUM1 To CardNowNUM2
                cardcountAInum(k, 5) = 出牌順序統計暫時變數(w, i, 2)
                cardAInumOvertenrecord(k) = 出牌順序統計暫時變數(w, i, 2)
                cardcountAInum(k, 1) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 1)
                cardcountAInum(k, 2) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 2)
                cardcountAInum(k, 3) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 3)
                cardcountAInum(k, 4) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 4)
                k = k + 1
            Next
         '========================
        智慧型AI系統類.智慧型AI系統計算_一階段_取得牌面資料 False, uscom
        智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turn, movecpre, uscom
        智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turn, movecpre, uscom
        智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
        智慧型AI系統類.智慧型AI系統計算_四階段_比序_1_初始
    '    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_1
        智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_2
        智慧型AI系統類.智慧型AI系統計算_暫時匯出 uscom
        智慧型AI系統類.智慧型AI系統計算_四階段_比序_3_選擇組合
        If turn = 3 And cardAInumchoose > 0 Then
            智慧型AI系統類.智慧型AI系統計算_引導程序_移動階段續 uscom, turn, name, movecpre, choose, (CardNowNUM2 - CardNowNUM1 + 1)
            Exit Do
        Else
            智慧型AI系統類.智慧型AI系統計算_最後階段_實行選牌 choose, uscom
        End If
        '==============================
        CardNowNUM1 = CardNowNUM1 + CardNumMax
        CardNowNUM2 = CardNowNUM2 + CardNumMax
        If CardNowNUM2 > Val(pageglead(uscom)) Then CardNowNUM2 = Val(pageglead(uscom))
        CardNowCountNUM = CardNowCountNUM + 1
    Loop Until CardNowCountNUM >= CardOverCountNUM
    '==========================
'    If turn <> 3 Then
'        戰鬥系統類.comatk_智慧型AI引導程序_超出牌張數 turn, movecpre, choose
'    End If
End If
End Sub
Sub 檢查人物技能是否有EX技(ByVal uscom As Integer, ByVal name As String)
'Erase personatkingtfr
'For i = 1 To 3
'     If VBEPerson(uscom, i, 1, 1, 1) = name Then
'         For k = 1 To 4
'               If Mid(VBEPerson(uscom, i, 3, k, 1), 1, 2) = "Ex" Then
'                   personatkingtfr(k) = 1
'               Else
'                   personatkingtfr(k) = 0
'               End If
'          Next
'          For k = 1 To 14
''                If (人物異常狀態資料庫(uscom, i, k, 3) = 22 And uscom = 1) Or _
''                    (人物異常狀態資料庫(uscom, i, k, 3) = 23 And uscom = 2) Then
'                 If 人物異常狀態資料庫(uscom, i, k, 3) = "BUFFN00701" Then
'                    personatkingtfr(5) = 1
'                End If
'          Next
'     End If
'Next
End Sub
Sub 智慧型AI系統_使用者出牌階段判斷反轉()
For i = 1 To 公用牌實體卡片分隔紀錄數(1)
    If Val(pagecardnum(i, 11)) = 4 And Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
        If pageonin(i) = 1 Then
           pageonin(i) = 2
        Else
           pageonin(i) = 1
        End If
        FormMainMode.card(i).CardRotationType = pageonin(i)
        FormMainMode.card_CardButtonClickin (i)
        pagecardnum(i, 11) = 3
    End If
Next
End Sub
Sub 智慧型AI系統計算_移動階段續_否定面_二階段_選擇行動(ByVal uscom As Integer)
Select Case uscom
    Case 1
        目前數(33) = 2
    Case 2
         電腦方移動階段選擇數 = 2
End Select
End Sub
Sub 智慧型AI系統_執行階段99_主動技能執行(ByVal uscom As Integer, ByVal turn As Integer, ByVal movecpre As Integer)
'=======================
執行階段系統_宣告開始或結束 1
'=======================
Dim VBEStageNumMain(1 To 1) As Integer
ReDim Vss_EventActiveAIScoreNum(1 To 1) As Integer
'=======================
For i = 1 To cardAITotalNUM
    For atkingnum = 1 To 4
        If Vss_PersonAtkingOffNum(uscom, 角色人物對戰人數(uscom, 2), atkingnum) = 0 And Val(VBEPerson(uscom, 角色人物對戰人數(uscom, 2), 3, atkingnum, 8)) = turn Then
            If 執行階段系統類.執行階段系統_驗證(atkingnum, 99, VBEPerson(uscom, 角色人物對戰人數(uscom, 2), 3, atkingnum, 11), uscom, 角色人物對戰人數(uscom, 2)) = True Then
                   智慧型AI系統類.智慧型AI系統_執行階段準備變數統合資料 uscom, VBEStageNumMain, turn, movecpre, i
                   智慧型AI系統類.智慧型AI系統_執行階段99_計算個別期望推薦值統計 uscom, atkingnum, i, turn, 角色人物對戰人數(uscom, 2)
            End If
        End If
    Next
Next
'=======================
執行階段系統_宣告開始或結束 2
'=======================
End Sub
Sub 智慧型AI系統_執行階段準備變數統合資料(ByVal uscom As Integer, ByRef VBEStageNumMain() As Integer, ByVal turnai As Integer, ByVal movecpre As Integer, ByVal cardAICaseNum As Integer)
    '===========================
    Erase VBEPersonVS 'VBE人物統一變數-VS版
    Erase atkingpagetotVS '每階段出牌種類及數值統計資料-VS版
    Erase VBEPersonBuffVSF  '異常狀態資料-VS-F版
    Erase VBEPersonBuffVSS  '異常狀態資料-VS-S版
    Erase AtkingckVSS '技能資訊一覽-S版(技能啟動碼)
    Erase AtkingckVSF '技能資訊一覽-F版(技能備註字串)
    Erase VBEAtkingVSF 'VBE>VS給予變數統一資料-F版
    Erase VBEAtkingVSS 'VBE>VS給予變數統一資料-S版
'    Erase VBEPageCardNumVS '公用牌資料-VS版
    ReDim VBEPageCardNumVS(1 To cardAInumuscom, 1 To 6) As Variant '公用牌資料-VS版
'    Erase VBEVSStageNum '執行階段系統-執行階段多用途紀錄變數-VS版
    ReDim VBEVSStageNum(1 To UBound(VBEStageNumMain)) As Variant '執行階段系統-執行階段多用途紀錄變數-VS版
    Erase VBEActualStatusVS '人物實際狀態資料-VS版
    '===========================
    Dim q As Integer, w As Integer, rr As Integer, cs1 As Variant, cs2 As Variant, tempc As Integer, buffobj As clsStatus
    tempc = 1
    For i = 1 To 2
        For j = 1 To 3
            If 人物異常狀態列表(i, j).Count > tempc Then
                tempc = 人物異常狀態列表(i, j).Count
            End If
        Next
    Next
    ReDim VBEPersonBuffVSF(1 To 2, 1 To 3, 1 To tempc, 1 To 2) As Variant '異常狀態資料-VS-F版
    ReDim VBEPersonBuffVSS(1 To 2, 1 To 3, 1 To tempc) As Variant '異常狀態資料-VS-S版
    '===========================
    Select Case uscom
         Case 1
             '(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11)
             For i = 1 To 2
                 For j = 1 To 3
                     For k = 1 To 4
                         For m = 1 To 30
                             For p = 1 To 11
                                 VBEPersonVS(i, j, k, m, p) = VBEPerson(i, 角色待機人物紀錄數(i, j), k, m, p)
                             Next
                         Next
                      Next
                 Next
            Next
            '======================
            For i = 1 To cardAInumuscom
                For j = 1 To 6
                    If j = 1 Or j = 3 Then
                       Select Case cardcountAInum(i, j)
                           Case "ATK-劍"
                               VBEPageCardNumVS(i, j) = 1
                           Case "DEF"
                               VBEPageCardNumVS(i, j) = 2
                           Case "MOV"
                               VBEPageCardNumVS(i, j) = 3
                           Case "SPE"
                               VBEPageCardNumVS(i, j) = 4
                           Case "ATK-槍"
                               VBEPageCardNumVS(i, j) = 5
                           Case "DRAW"
                               VBEPageCardNumVS(i, j) = 6
                           Case "BRK"
                               VBEPageCardNumVS(i, j) = 7
                           Case "HPL"
                               VBEPageCardNumVS(i, j) = 8
                           Case Else
                               VBEPageCardNumVS(i, j) = 0
                       End Select
                    ElseIf j >= 5 Then
                       VBEPageCardNumVS(i, j) = 1
                    Else
                        VBEPageCardNumVS(i, j) = Val(cardcountAInum(i, j))
                    End If
                Next
                '==================
                If Mid(cardAInumnm(cardAICaseNum - 1), i, 1) = 1 Then
                    cs1 = VBEPageCardNumVS(i, 1)
                    cs2 = VBEPageCardNumVS(i, 2)
                    VBEPageCardNumVS(i, 1) = VBEPageCardNumVS(i, 3)
                    VBEPageCardNumVS(i, 2) = VBEPageCardNumVS(i, 4)
                    VBEPageCardNumVS(i, 3) = cs1
                    VBEPageCardNumVS(i, 4) = cs2
                End If
                '==================
            Next
            '======================
            '(1 To 2, 1 To 5)
            For j = 1 To 5
                atkingpagetotVS(1, j) = cardAInumcaseperson(cardAICaseNum, 1, 10 + j)
            Next
            For j = 1 To 5
                atkingpagetotVS(2, j) = atkingpagetot(2, j)
            Next
            '======================
            '(1 To 2, 1 To 3, 1 To 14, 1 To 3)
            For i = 1 To 2
                For rr = 1 To 3
                    For j = 1 To 人物異常狀態列表(i, 角色待機人物紀錄數(i, rr)).Count
                        Set buffobj = 人物異常狀態列表(i, 角色待機人物紀錄數(i, rr))(j)
                        VBEPersonBuffVSF(i, rr, j, 1) = buffobj.Value
                        VBEPersonBuffVSF(i, rr, j, 2) = buffobj.Total
                        VBEPersonBuffVSS(i, rr, j) = buffobj.Identifier
                    Next
                Next
            Next
            '======================
            '(1 to 2,1 to 3,1 to 2)
            For i = 1 To 2
                For rr = 1 To 3
                    VBEActualStatusVS(i, rr, 1) = 人物實際狀態資料庫(i, 角色待機人物紀錄數(i, rr), 1)
                    VBEActualStatusVS(i, rr, 2) = 人物實際狀態資料庫(i, 角色待機人物紀錄數(i, rr), 9)
                Next
            Next
            '======================
            '(1 to 8,1 to 3)
            For i = 1 To 8
                For j = 1 To 3
                    AtkingckVSS(i, j) = atkingck(uscom, 角色人物對戰人數(uscom, 2), i, j)
                Next
                AtkingckVSF(i, 1) = Vss_AtkingInformationRecordStr(uscom, 角色人物對戰人數(uscom, 2), i)
            Next
            '======================
            For i = 1 To 3
                VBEAtkingVSF(1, i, 1) = liveus(角色待機人物紀錄數(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(1, i, 2) = liveusmax(角色待機人物紀錄數(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(2, i, 1) = livecom(角色待機人物紀錄數(2, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(2, i, 2) = livecommax(角色待機人物紀錄數(2, i))
            Next
            '========================
            VBEAtkingVSS(0) = 1
            VBEAtkingVSS(1) = cardAInumuscom
            VBEAtkingVSS(2) = 0
            VBEAtkingVSS(3) = pageqlead(2)
            VBEAtkingVSS(4) = pageglead(2)
            VBEAtkingVSS(6) = movecpre
            If 是否移動階段續估計判斷程序 = False Then
                VBEAtkingVSS(5) = 擲骰表單溝通暫時變數(2)
                VBEAtkingVSS(7) = Val(攻擊防禦骰子總數(1))
                VBEAtkingVSS(8) = Val(攻擊防禦骰子總數(2))
                VBEAtkingVSS(14) = 擲骰表單溝通暫時變數(5)
                VBEAtkingVSS(15) = 擲骰表單溝通暫時變數(6)
                VBEAtkingVSS(16) = moveturn
            Else
                VBEAtkingVSS(5) = 0
                VBEAtkingVSS(7) = 0
                VBEAtkingVSS(8) = 0
                VBEAtkingVSS(14) = 0
                VBEAtkingVSS(15) = 0
                VBEAtkingVSS(16) = 1
            End If
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            Select Case turnai
                Case 1
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 2
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
             End Select
             VBEAtkingVSS(17) = 角色人物對戰人數(1, 1)
             VBEAtkingVSS(18) = 角色人物對戰人數(2, 1)
             VBEAtkingVSS(19) = 牌總階段數(1)
             VBEAtkingVSS(20) = 牌總階段數(2)
             '=========================
             For i = 1 To UBound(VBEStageNumMain)
                 If VBEStageNumMain(i) = -1 Or VBEStageNumMain(i) = -2 Then
                     VBEVSStageNum(i) = Abs(VBEStageNumMain(i))
                 Else
                     VBEVSStageNum(i) = VBEStageNumMain(i)
                 End If
             Next
         Case 2 '===============================================================
             '(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11)
             For i = 1 To 2
                 If i = 1 Then q = 2 Else q = 1
                 For j = 1 To 3
                     For k = 1 To 4
                         For m = 1 To 30
                             For p = 1 To 11
                                 VBEPersonVS(i, j, k, m, p) = VBEPerson(q, 角色待機人物紀錄數(q, j), k, m, p)
                             Next
                         Next
                      Next
                 Next
            Next
            '======================
            For i = 1 To cardAInumuscom
                For j = 1 To 6
                     If j = 1 Or j = 3 Then
                       Select Case cardcountAInum(i, j)
                           Case "ATK-劍"
                               VBEPageCardNumVS(i, j) = 1
                           Case "DEF"
                               VBEPageCardNumVS(i, j) = 2
                           Case "MOV"
                               VBEPageCardNumVS(i, j) = 3
                           Case "SPE"
                               VBEPageCardNumVS(i, j) = 4
                           Case "ATK-槍"
                               VBEPageCardNumVS(i, j) = 5
                           Case "DRAW"
                               VBEPageCardNumVS(i, j) = 6
                           Case "BRK"
                               VBEPageCardNumVS(i, j) = 7
                           Case "HPL"
                               VBEPageCardNumVS(i, j) = 8
                           Case Else
                               VBEPageCardNumVS(i, j) = 0
                       End Select
                    ElseIf j >= 5 Then
                        VBEPageCardNumVS(i, j) = 1
                    Else
                       VBEPageCardNumVS(i, j) = Val(cardcountAInum(i, j))
                    End If
                Next
                '==================
                If Mid(cardAInumnm(cardAICaseNum - 1), i, 1) = 1 Then
                    cs1 = VBEPageCardNumVS(i, 1)
                    cs2 = VBEPageCardNumVS(i, 2)
                    VBEPageCardNumVS(i, 1) = VBEPageCardNumVS(i, 3)
                    VBEPageCardNumVS(i, 2) = VBEPageCardNumVS(i, 4)
                    VBEPageCardNumVS(i, 3) = cs1
                    VBEPageCardNumVS(i, 4) = cs2
                End If
                '==================
            Next
            '======================
            '(1 To 2, 1 To 5)
            For j = 1 To 5
                atkingpagetotVS(1, j) = cardAInumcaseperson(cardAICaseNum, 1, 10 + j)
            Next
            For j = 1 To 5
                atkingpagetotVS(2, j) = atkingpagetot(1, j)
            Next
            '======================
            '(1 To 2, 1 To 3, 1 To 14, 1 To 3)
            For i = 1 To 2
                If i = 1 Then q = 2 Else q = 1
                For rr = 1 To 3
                    For j = 1 To 人物異常狀態列表(q, 角色待機人物紀錄數(q, rr)).Count
                        Set buffobj = 人物異常狀態列表(q, 角色待機人物紀錄數(q, rr))(j)
                        VBEPersonBuffVSF(i, rr, j, 1) = buffobj.Value
                        VBEPersonBuffVSF(i, rr, j, 2) = buffobj.Total
                        VBEPersonBuffVSS(i, rr, j) = buffobj.Identifier
                    Next
                Next
            Next
            '======================
            '(1 to 2,1 to 3,1 to 2)
            For i = 1 To 2
                If i = 1 Then w = 2 Else w = 1
                For rr = 1 To 3
                    VBEActualStatusVS(i, rr, 1) = 人物實際狀態資料庫(w, 角色待機人物紀錄數(w, rr), 1)
                    VBEActualStatusVS(i, rr, 2) = 人物實際狀態資料庫(w, 角色待機人物紀錄數(w, rr), 9)
                Next
            Next
            '======================
            '(1 to 8,1 to 3)
            For i = 1 To 8
                For j = 1 To 3
                    AtkingckVSS(i, j) = atkingck(uscom, 角色人物對戰人數(uscom, 2), i, j)
                Next
                AtkingckVSF(i, 1) = Vss_AtkingInformationRecordStr(uscom, 角色人物對戰人數(uscom, 2), i)
            Next
            '======================
            For i = 1 To 3
                VBEAtkingVSF(2, i, 1) = liveus(角色待機人物紀錄數(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(2, i, 2) = liveusmax(角色待機人物紀錄數(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(1, i, 1) = livecom(角色待機人物紀錄數(2, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(1, i, 2) = livecommax(角色待機人物紀錄數(2, i))
            Next
            '========================
            VBEAtkingVSS(0) = 1
            VBEAtkingVSS(1) = cardAInumuscom
            VBEAtkingVSS(2) = 0
            VBEAtkingVSS(3) = pageqlead(1)
            VBEAtkingVSS(4) = pageglead(1)
            VBEAtkingVSS(6) = movecpre
            If 是否移動階段續估計判斷程序 = False Then
                VBEAtkingVSS(5) = 擲骰表單溝通暫時變數(2)
                VBEAtkingVSS(7) = Val(攻擊防禦骰子總數(2))
                VBEAtkingVSS(8) = Val(攻擊防禦骰子總數(1))
                VBEAtkingVSS(14) = 擲骰表單溝通暫時變數(6)
                VBEAtkingVSS(15) = 擲骰表單溝通暫時變數(5)
                If moveturn = 2 Then VBEAtkingVSS(16) = 1 Else VBEAtkingVSS(16) = 2
            Else
                VBEAtkingVSS(5) = 0
                VBEAtkingVSS(7) = 0
                VBEAtkingVSS(8) = 0
                VBEAtkingVSS(14) = 0
                VBEAtkingVSS(15) = 0
                VBEAtkingVSS(16) = 1
            End If
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            Select Case turnai
                Case 1
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 2
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
             End Select
             VBEAtkingVSS(17) = 角色人物對戰人數(2, 1)
             VBEAtkingVSS(18) = 角色人物對戰人數(1, 1)
             VBEAtkingVSS(19) = 牌總階段數(2)
             VBEAtkingVSS(20) = 牌總階段數(1)
             '=========================
             For i = 1 To UBound(VBEStageNumMain)
                 If VBEStageNumMain(i) = -1 Then
                     VBEVSStageNum(i) = 2
                 ElseIf VBEStageNumMain(i) = -2 Then
                     VBEVSStageNum(i) = 1
                 Else
                     VBEVSStageNum(i) = VBEStageNumMain(i)
                 End If
             Next
   End Select
End Sub
Sub 智慧型AI系統_執行階段99_計算個別期望推薦值統計(ByVal uscom As Integer, ByVal atkingnum As Integer, ByVal cardAICaseNum As Integer, ByVal turn As Integer, ByVal personnum As Integer)
Dim vsstr As String, vsstr2() As String, vsstr3() As String, vsstr4() As String, vstest As String, uscomt As Integer
'============擷取執行階段99之評分資料
vsstr = 執行階段系統類.執行階段系統_執行腳本_人物主動技能類(atkingnum, 99, uscom, personnum)
vsstr2 = Split(vsstr, "=")
For i = 0 To UBound(vsstr2)
    If vsstr2(i) <> "" Then
        vsstr3 = Split(vsstr2(i), "#")
        If vsstr3(0) = "EventActiveAIScore" Then
            vsstr4 = Split(vsstr3(1), ",")
            vstest = vsstr2(i)
            '===================================
            ReDim Vss_EventActiveAIScoreNum(1 To UBound(vsstr4) + 1) As Integer
            For k = 0 To UBound(vsstr4)
                Vss_EventActiveAIScoreNum(k + 1) = vsstr4(k)
            Next
            '===================================
            Exit For
        End If
    End If
Next
'==================================================================
If Vss_EventActiveAIScoreNum(1) = 1 Then
    If Vss_EventActiveAIScoreNum(2) = 1 Then
        cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + 10
    ElseIf Vss_EventActiveAIScoreNum(2) = 2 Then
        '============擷取執行階段45之總骰數變化量資料
        vsstr = 執行階段系統類.執行階段系統_執行腳本_人物主動技能類(atkingnum, 45, uscom, personnum)
        vsstr2 = Split(vsstr, "=")
        For i = 0 To UBound(vsstr2)
            If vsstr2(i) <> "" Then
                vsstr3 = Split(vsstr2(i), "#")
                If vsstr3(0) = "EventTotalDiceChange" Then
                    vsstr4 = Split(vsstr3(1), ",")
                    '===================================
                    If Val(vsstr4(0)) = 1 Then  '為自身所變化之量
                        Select Case Val(vsstr4(1))
                            Case 1
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + Val(vsstr4(2))
                            Case 2
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) - Val(vsstr4(2))
                            Case 3
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) * Val(vsstr4(2))
                            Case Is <= 5
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) \ Val(vsstr4(2))
                            Case 6
                                If turn = 1 Then
                                    cardAInumFinal(cardAICaseNum, 1) = Val(vsstr4(2))
                                ElseIf turn = 2 Then
                                    cardAInumFinal(cardAICaseNum, 1) = Val(vsstr4(2)) - VBEPerson(uscom, 角色人物對戰人數(uscom, 2), 1, 3, 3)
                                End If
                        End Select
                    ElseIf Val(vsstr4(0)) = 2 Then  '為對方所變化之量
                        Select Case Val(vsstr4(1))
                            Case 1
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) - Val(vsstr4(2))
                            Case 2
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + Val(vsstr4(2))
                            Case 3
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) \ Val(vsstr4(2))
                            Case Is <= 5
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) * Val(vsstr4(2))
                            Case 6
                                If uscom = 1 Then uscomt = 2 Else uscomt = 1
                                If turn = 1 Then
                                    cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + (VBEPerson(uscomt, 角色人物對戰人數(uscomt, 2), 1, 3, 3) - Val(vsstr4(2))) * 2 + 5
                                ElseIf turn = 2 Then
                                    cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + (VBEPerson(uscomt, 角色人物對戰人數(uscomt, 2), 1, 3, 2) - Val(vsstr4(2))) * 2 + 5
                                End If
                        End Select
                    End If
                    '===================================
'                    Exit For
                End If
            End If
        Next
    End If
    '=====================================
    If Vss_EventActiveAIScoreNum(2) = 1 Or Vss_EventActiveAIScoreNum(2) = 2 Then
        For i = 3 To UBound(Vss_EventActiveAIScoreNum)
            If Vss_EventActiveAIScoreNum(i) > 0 And Vss_EventActiveAIScoreNum(i) <= cardAInumuscom Then
                cardAInumcaseperson(cardAICaseNum, 2, Vss_EventActiveAIScoreNum(i)) = 1
'                 MsgBox vstest & Chr(10) & cardcountAInum(Vss_EventActiveAIScoreNum(i), 1) & "," & cardcountAInum(Vss_EventActiveAIScoreNum(i), 2) & ",  " & cardcountAInum(Vss_EventActiveAIScoreNum(i), 3) & "," & cardcountAInum(Vss_EventActiveAIScoreNum(i), 4) & Chr(10) & "uscom:" & uscom & "  ,atkingnum:" & atkingnum
            End If
        Next
    End If
    '=====================================
ElseIf Vss_EventActiveAIScoreNum(1) = 3 Then
    cardAInumFinal(cardAICaseNum, 1) = -100
End If
'=================
ReDim Vss_EventActiveAIScoreNum(1 To 1) As Integer
End Sub
