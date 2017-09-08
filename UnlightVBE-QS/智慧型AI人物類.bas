Attribute VB_Name = "智慧型AI人物類"
Public 夏洛特_階段處理記錄數(1 To 3) As Integer '智慧型AI-夏洛特-戰略判斷紀錄數(1.當前階段實行/2.目標結束之回合數/3.幸福的理由是否發動)
Sub 艾伯李斯特(ByVal turn As Integer, ByVal movecpre As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer '暫時變數
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================精密射擊
                 If Pn1 = 0 Then
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 10) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 10) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================雷擊
                 If Pn2 = 0 Then
                         If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 14) >= 2 Then
                                     cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                     '=====================
                                     sq = 1
                                     wnm = 0
                                     Do '==先從特卡卡面號低處開始
                                        If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                              For k = 1 To cardAInumuscom
                                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                        Case 0
                                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                                 wnm = wnm + sq
                                                             End If
                                                        Case 1
                                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                                 wnm = wnm + sq
                                                             End If
                                                   End Select
                                                   If wnm >= 2 Then Exit Do
                                              Next
                                        End If
                                        sq = sq + 1
                                    Loop Until sq > 10
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
               '===========茨林
               If Pn3 = 0 Then
                    If movecpre = 1 Then
                            If cardAInumcaseperson(i, 1, 14) >= 2 And cardAInumcaseperson(i, 1, 12) >= 2 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '=====================
                                  sq = 1
                                  wnm = 0
                                  Do '==先從特卡卡面號低處開始
                                     If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                              wnm = wnm + sq
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                              wnm = wnm + sq
                                                          End If
                                                End Select
                                                If wnm >= 2 Then Exit Do
                                           Next
                                     End If
                                     sq = sq + 1
                                 Loop Until sq > 10
                            End If
                     End If
               End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
                '=================智略
                If Pn4 = 0 Then
                        Dim werp As Integer
                        werp = 0
                        For k = 1 To cardAInumuscom
                              If cardAInumcaseperson(i, 2, k) > 0 Then
                                  werp = Val(werp) + 1
                              End If
                        Next
                        If Val(werp) >= 3 Then
        '                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                            werp = 0
                            For k = 1 To cardAInumuscom
                                If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                    cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                    werp = Val(werp) + 1
                                End If
                            Next
                            If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                        ElseIf Val(werp) < 3 Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 3 Then
                                werp = 0
                                '==============1.針對已預定出牌的部分作加成
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.對數值為1的牌作加成
                                If werp < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.對數值為2的牌作加成
                                If Val(werp) < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                            End If
                        End If
                End If
         Next
End Select

End Sub
Sub 雪莉(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, livewer As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
'=============================
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================自殺傾向
                 If Pn1 = 0 Then
                        If cardAInumcaseperson(i, 1, 14) >= 1 Then
                            If cardAInumcaseperson(i, 1, 14) < livewer Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + Val(cardAInumcaseperson(i, 1, 14)) * 5
                            '======================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 2)) * 5
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 4)) * 5
                                                   End If
                                         End Select
                                    Next
                            ElseIf cardAInumcaseperson(i, 1, 14) >= livewer Then '==組合下特卡數超過自身血量時
                                    cardAInumFinal(i, 1) = -10000
                            End If
                        End If
                 End If
                 '================飛刃雨
                 If Pn4 = 0 Then
                         If movecpre = 3 Then
                               If cardAInumcasepersonTER(i, 3, 1) >= 1 Then
                                     '=====================
                                    For k = 1 To cardAInumuscom
                                         Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0 '===防止計算到事件卡
                                                       If cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a3a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 2
                                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 2
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a3a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 2
                                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 2
                                                       End If
                                             End Select
                                    Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
                '===========異質者
                If Pn2 = 0 Then
                        If cardAInumcasepersonTER(i, 4, 3) >= 1 And _
                            執行動作_檢查是否有指定異常狀態(uscom, "BUFFN01001") = False Then
                              If (攻擊防禦骰子總數(1) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And 是否移動階段續估計判斷程序 = False And uscom = 2) Or _
                                 (攻擊防禦骰子總數(2) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And 是否移動階段續估計判斷程序 = False And uscom = 1) Or _
                                 (是否移動階段續估計判斷程序 = True And Val(livewer) <= 3) Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10000
                              '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                         End Select
                                    Next
                              End If
                        End If
                End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
                '=====================巨大黑犬
                If Pn3 = 0 Then
                        If movecpre < 3 Then
                            Dim werp As Integer
                            werp = 0
                            If cardAInumcaseperson(i, 1, 11) >= 3 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '=====================
                                    For p = Val(cardAInumcaseperson(i, 1, 2)) To 1 Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 2))
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 4))
                                                       End If
                                             End Select
                                        Next
                                    Next
                            End If
                        End If
                End If
        Next
End Select
End Sub
Sub 艾茵(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, livewer As Integer, livewermax As Integer '暫時變數
If uscom = 1 Then
    livewer = liveus(角色人物對戰人數(1, 2))
    livewermax = liveusmax(角色人物對戰人數(1, 2))
Else
    livewer = livecom(角色人物對戰人數(2, 2))
    livewermax = livecommax(角色人物對戰人數(2, 2))
End If
'================
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================十三隻眼
                 If Pn4 = 0 Then
                        If movecpre < 3 Then
                               If cardAInumcasepersonTER(i, 1, 3) >= 1 And cardAInumcasepersonTER(i, 5, 3) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If (cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 3) Or (cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 3) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 30
                                                  Else
                                                      cardAInumcaseperson(i, 2, k) = 0
                                                  End If
                                             Case 1
                                                  If (cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 3) Or (cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 3) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 30
                                                  Else
                                                      cardAInumcaseperson(i, 2, k) = 0
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
           Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
                '====================兩個身體
                If Pn2 = 0 Then
                        If cardAInumcaseperson(i, 1, 13) >= 1 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                              '=====================
                              sq = 1
                              wnm = 0
                              Do '==先從移卡卡面號低處開始
                                 If cardAInumcasepersonTER(i, 3, sq) >= 1 Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + (11 - sq)
                                                          wnm = wnm + sq
                                                          If cardcountAInum(k, 3) <> a4a Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          End If
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + (11 - sq)
                                                          wnm = wnm + sq
                                                          If cardcountAInum(k, 1) <> a4a Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          End If
                                                      End If
                                            End Select
                                            If wnm >= 1 Then Exit Do
                                       Next
                                 End If
                                 sq = sq + 1
                             Loop Until sq > 10
                        End If
                End If
                '=====================九個靈魂
                If Pn3 = 0 Then
                        If movecpre > 1 Then
                            If cardAInumcaseperson(i, 1, 12) >= 5 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  If livewer = livewermax Then
                                      cardAInumFinal(i, 1) = cardAInumFinal(i, 1) - ((cardAInumcaseperson(i, 1, 14) - cardAInumcaseperson(i, 1, 7)) * 2)
                                  End If
                                  '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 5 * Val(cardcountAInum(k, 2))
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 5 * Val(cardcountAInum(k, 4))
                                                   End If
                                         End Select
                                    Next
                            End If
                        End If
                End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
                '=====================一顆心
                If Pn1 = 0 Then
                        If movecpre = 2 Then
                            Dim werp As Integer
                            werp = 0
                            If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '=====================
                                    For p = Val(cardAInumcaseperson(i, 1, 8)) To 1 Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 2))
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 4))
                                                       End If
                                             End Select
                                        Next
                                    Next
                            End If
                        End If
                End If
        Next
End Select

End Sub
Sub 古魯瓦爾多(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, weryu(1 To 3) As Integer '暫時變數
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================猛擊
                 If Pn1 = 0 Then
                        If movecpre = 1 Then
                               If cardAInumcasepersonTER(i, 1, 1) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
                '===========血之恩賜
                If Pn3 = 0 Then
                        If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                              '=====================
                              sq = 1
                              wnm = 0
                              Do '==先從特卡卡面號低處開始
                                 If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                          wnm = wnm + sq
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                          wnm = wnm + sq
                                                      End If
                                            End Select
                                            If wnm >= 2 Then Exit Do
                                       Next
                                 End If
                                 sq = sq + 1
                             Loop Until sq > 10
                        End If
                End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
                '=====================必殺架勢
                If Pn2 = 0 Then
                        werp = 0
                        If cardAInumcaseperson(i, 1, 14) >= 2 And cardAInumcaseperson(i, 1, 13) = 0 Then
                              For k = 1 To 14
                                    If 人物異常狀態資料庫(2, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN00801" Then
                                        werp = 1
                                    End If
                                    If 人物異常狀態資料庫(2, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN00302" Then
                                        werp = 1
                                    End If
                              Next
                              If werp = 1 Then
                                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                    werp = 0
                                    '=====================
                                      For p = Val(cardAInumcaseperson(i, 1, 8)) To 1 Step -1
                                          For k = 1 To cardAInumuscom
                                              Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                    Case 0
                                                         If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 2 Then
                                                             cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                             werp = Val(werp) + Val(cardcountAInum(k, 2))
                                                         End If
                                                    Case 1
                                                         If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 2 Then
                                                             cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                             werp = Val(werp) + Val(cardcountAInum(k, 4))
                                                         End If
                                               End Select
                                          Next
                                      Next
                                End If
                        End If
                End If
                '=====================精神力吸收
                If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcasepersonTER(i, 1, 1) >= 1 And cardAInumcasepersonTER(i, 5, 1) >= 1 And cardAInumcasepersonTER(i, 4, 1) >= 1 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                              '=====================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If (cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 1 And Val(weryu(3)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 1) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 1) = a5a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 1) = a4a Then
                                                       weryu(3) = weryu(3) + 1
                                                   End If
                                               End If
                                          Case 1
                                               If (cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 1 And Val(weryu(3)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 3) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 3) = a5a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 3) = a4a Then
                                                       weryu(3) = weryu(3) + 1
                                                   End If
                                               End If
                                     End Select
                                Next
                        End If
                End If
        Next
End Select

End Sub
Sub 帕茉(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, weryu(1 To 3) As Integer, livewer As Integer, livewermax As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
If uscom = 1 Then livewermax = liveusmax(角色人物對戰人數(1, 2)) Else livewermax = livecommax(角色人物對戰人數(2, 2))
'=============================
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================靜謐之背
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 2 Then
                                   For k = 1 To 14
                                         If 人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN01401" Then
                                             werp = Val(人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 2))
                                         End If
                                   Next
                                   If werp > 0 Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (Val(werp) - 6) * 3
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 2))
                                                          End If
                                                          If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + Val(cardcountAInum(k, 2))
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 4))
                                                          End If
                                                          If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + Val(cardcountAInum(k, 4))
                                                          End If
                                                End Select
                                           Next
                                   End If
                               End If
                        End If
                 End If
                 '=================慈悲的藍眼
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 6 Then
                                   For k = 1 To 14
                                         If 人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN01401" Then
                                             werp = Val(人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 2))
                                         End If
                                   Next
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (9 - Val(werp)) * 3
                                   If livewer = livewermax And werp = 9 Then
                                       cardAInumFinal(i, 1) = 0
                                   End If
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 2))
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 4))
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================戰慄的狼牙
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 6 Then
                                   For k = 1 To 14
                                         If 人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN01401" Then
                                             werp = Val(人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 2))
                                         End If
                                   Next
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (Val(werp) - 8) * 3 + (3 - Val(livewer)) * 5
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 2))
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 4))
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
                '=====================憤怒之爪
                If Pn1 = 0 Then
                        werp = 0
                        If cardAInumcaseperson(i, 1, 14) >= 1 Then
                            For k = 1 To 14
                                  If 人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN01401" Then
                                      werp = Val(人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 2))
                                  End If
                            Next
                            If werp < 9 Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                '=====================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   Exit For
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   Exit For
                                               End If
                                     End Select
                                Next
                            End If
                        End If
                End If
        Next
End Select

End Sub
Sub 史塔夏(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================愚者之手
                 If Pn2 = 0 Then
                        werp = 0
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 6 Then
'                                   If atking_史塔夏_殺戮模式狀態數(2) = 1 Then
                                   If 人物實際狀態資料庫(uscom, 角色待機人物紀錄數(uscom, 1), 1) = "UCASN00101" Then
                                       For k = 1 To 3
                                             Select Case uscom
                                                   Case 1
                                                        If liveus(角色待機人物紀錄數(uscom, 1)) > 0 Then
                                                            werp = Val(werp) + 1
                                                        End If
                                                   Case 2
                                                        If livecom(角色待機人物紀錄數(uscom, 1)) > 0 Then
                                                            werp = Val(werp) + 1
                                                        End If
                                            End Select
                                        Next
                                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (livewer - werp) * 4
                                   Else
                                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   End If
                                   '======================
                                   werp = 0
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And werp < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      werp = Val(werp) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And werp < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      werp = Val(werp) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================命運的鐵門
                 If Pn4 = 0 Then
                         Erase weryu
                         werp = 0
                         If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 9 Then
                                     For k = 1 To 3
                                          weryu(k) = 999   '目的取最低HP量
                                     Next
                                     Select Case uscom
                                          Case 1
                                               For k = 2 To 3
                                                     If Val(liveus(角色待機人物紀錄數(1, k))) < Val(weryu(1)) And Val(liveus(角色待機人物紀錄數(1, k))) > 0 Then
                                                         weryu(1) = liveus(角色待機人物紀錄數(1, k))
                                                    End If
                                               Next
                                               For k = 1 To 3
                                                     If Val(livecom(角色待機人物紀錄數(2, k))) < Val(weryu(2)) And Val(livecom(角色待機人物紀錄數(2, k))) > 0 Then
                                                         weryu(2) = livecom(角色待機人物紀錄數(2, k))
                                                    End If
                                               Next
                                          Case 2
                                               For k = 2 To 3
                                                     If Val(livecom(角色待機人物紀錄數(2, k))) < Val(weryu(1)) And Val(livecom(角色待機人物紀錄數(2, k))) > 0 Then
                                                         weryu(1) = livecom(角色待機人物紀錄數(2, k))
                                                    End If
                                               Next
                                               For k = 1 To 3
                                                     If Val(liveus(角色待機人物紀錄數(1, k))) < Val(weryu(2)) And Val(liveus(角色待機人物紀錄數(1, k))) > 0 Then
                                                         weryu(2) = liveus(角色待機人物紀錄數(1, k))
                                                    End If
                                               Next
                                     End Select
                                     If Val(weryu(2)) < Val(weryu(1)) Then
                                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                         '=====================
                                            For k = 1 To cardAInumuscom
                                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                      Case 0
                                                           If cardcountAInum(k, 1) = a1a Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           End If
                                                      Case 1
                                                           If cardcountAInum(k, 3) = a1a Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           End If
                                                 End Select
                                            Next
                                     End If
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
               '=================時間種子
               If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 12) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   werp = 0
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
                '=================殺戮器官
                If Pn1 = 0 Then
                        werp = 0
                        For k = 1 To cardAInumuscom
                              If cardAInumcaseperson(i, 2, k) > 0 Then
                                  werp = Val(werp) + 1
                              End If
                        Next
                        If Val(werp) >= 3 Then
        '                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                            werp = 0
                            For k = 1 To cardAInumuscom
                                If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                    cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                    werp = Val(werp) + 1
                                End If
                            Next
                            If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                        ElseIf Val(werp) < 3 Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 3 Then
                                werp = 0
                                '==============1.針對已預定出牌的部分作加成
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.對數值為1的牌作加成
                                If werp < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.對數值為2的牌作加成
                                If Val(werp) < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                        End If
                End If
         Next
End Select

End Sub
Sub CC(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================白銀戰機
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 15) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================高頻電磁手術刀
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 14) >= 6 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================原子之心
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================滅菌空間
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
            End If
         Next
End Select

End Sub
Sub 伊芙琳(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================紅蓮車輪
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 15) >= 2 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   For k = 1 To 14
                                         If 人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 3) = "BUFFN01501" Then
                                             werp = Val(人物異常狀態資料庫(uscom, 角色人物對戰人數(uscom, 2), k, 2))
                                         End If
                                   Next
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (9 - werp) * 5
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================慟哭之歌
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre > 1 Then
                           If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================怠惰的墓表
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre < 3 Then
                           If cardAInumcaseperson(i, 1, 14) >= 2 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '=====================
                                  sq = 1
                                  wnm = 0
                                  Do '==先從特卡卡面號低處開始
                                     If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                End Select
                                                If wnm >= 2 Then Exit Do
                                           Next
                                     End If
                                     sq = sq + 1
                                 Loop Until sq > 10
                           End If
                     End If
                End If
              '=====================赤紅石榴
                If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                                If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And _
                                    cardAInumcaseperson(i, 1, 13) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 Then
                                      cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                                      '=====================
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If (cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 1) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 1) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 1) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 1) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 1) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                                  Case 1
                                                       If (cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 3) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 3) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 3) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 3) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 3) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                             End Select
                                        Next
                                End If
                        End If
                End If
         Next
End Select

End Sub
Sub 布勞(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================時間爆彈
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                 End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================時間追獵
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre < 3 Then
                           If cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================發條機構
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre > 1 Then
                           If cardAInumcaseperson(i, 1, 14) >= 2 Then
                               '=====如下2回合後屬質數回合才做總期望值加成
                               If BattleTurn + 2 = 3 Or BattleTurn + 2 = 5 Or BattleTurn + 2 = 7 Or _
                                  BattleTurn + 2 = 11 Or BattleTurn + 2 = 13 Or BattleTurn + 2 = 17 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               End If
                               '=====================
                                  sq = 1
                                  wnm = 0
                                  Do '==先從特卡卡面號低處開始
                                     If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                End Select
                                                If wnm >= 2 Then Exit Do
                                           Next
                                     End If
                                     sq = sq + 1
                                 Loop Until sq > 10
                           End If
                     End If
                End If
              '=====================夜幕時分
              If Pn4 = 0 Then
                        werp = 0
                        wnm = 0
                        Erase weryu
                        For k = 1 To cardAInumuscom
                              If cardAInumcaseperson(i, 2, k) > 0 Then
                                  werp = Val(werp) + 1
                              End If
                        Next
                        For k = 1 To 3
                            If VBEPerson(uscom, 角色待機人物紀錄數(uscom, k), 1, 2, 1) = "R" Then
                                 weryu(1) = Val(weryu(1)) + 1
                                 Select Case uscom
                                     Case 1
                                          If (Val(liveus(角色待機人物紀錄數(1, k))) + 3) <= Val(liveusmax(角色待機人物紀錄數(1, k))) Then
                                              weryu(2) = Val(weryu(2)) + 1
                                          End If
                                     Case 2
                                          If (Val(livecom(角色待機人物紀錄數(2, k))) + 3) <= Val(livecommax(角色待機人物紀錄數(2, k))) Then
                                              weryu(2) = Val(weryu(2)) + 1
                                          End If
                                End Select
                            End If
                        Next
                        If Val(werp) >= 3 Then
        '                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                            werp = 0
                            For k = 1 To cardAInumuscom
                                If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                    cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                    werp = Val(werp) + 1
                                End If
                            Next
                            If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                        ElseIf Val(werp) < 3 And Val(weryu(1)) >= 1 Then
                            If Val(weryu(2)) = 0 Then
                                wnm = 2
                            ElseIf Val(weryu(2)) > 0 Then
                                wnm = 3
                            End If
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= wnm And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= wnm And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 3 Then
                                werp = 0
                                '==============1.針對已預定出牌的部分作加成
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.對數值為1的牌作加成
                                If werp < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.對數值為2的牌作加成
                                If Val(werp) < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============4.對數值為3的牌作加成
                                If Val(werp) < 3 And Val(weryu(2)) > 0 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                        End If
                End If
         Next
End Select

End Sub
Sub 梅倫(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=====================Lowball
                If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And _
                            cardAInumcaseperson(i, 1, 13) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                              '=====================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If (cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a3a And Val(weryu(3)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a4a And Val(weryu(4)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a5a And Val(weryu(5)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 1) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 1) = a2a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 1) = a3a Then
                                                       weryu(3) = weryu(3) + 1
                                                   ElseIf cardcountAInum(k, 1) = a4a Then
                                                       weryu(4) = weryu(4) + 1
                                                   ElseIf cardcountAInum(k, 1) = a5a Then
                                                       weryu(5) = weryu(5) + 1
                                                   End If
                                               End If
                                          Case 1
                                               If (cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a3a And Val(weryu(3)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a4a And Val(weryu(4)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a5a And Val(weryu(5)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 3) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 3) = a2a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 3) = a3a Then
                                                       weryu(3) = weryu(3) + 1
                                                   ElseIf cardcountAInum(k, 3) = a4a Then
                                                       weryu(4) = weryu(4) + 1
                                                   ElseIf cardcountAInum(k, 3) = a5a Then
                                                       weryu(5) = weryu(5) + 1
                                                   End If
                                               End If
                                     End Select
                                Next
                        End If
                 End If
                 '=================Gamble
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If Len(cardAInumnm(i - 1)) >= 3 Then
                                   For k = 1 To cardAInumuscom
                                         If cardAInumcaseperson(i, 2, k) > 0 Then
                                             werp = Val(werp) + 1
                                         End If
                                   Next
                                   If Val(werp) >= 3 Then
                                       werp = 0
                                       For k = 1 To cardAInumuscom
                                           If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                               cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                               werp = Val(werp) + 1
                                           End If
                                       Next
                                       If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   Else
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a5a) _
                                                           And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a5a) _
                                                           And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                      End If
                                            End Select
                                       Next
                                       If Val(werp) >= 3 Then
                                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                               '======================
                                               werp = 0
                                               For k = 1 To cardAInumuscom
                                                     If cardAInumcaseperson(i, 2, k) > 0 Then
                                                         werp = Val(werp) + 1
                                                     End If
                                               Next
                                               For k = 1 To cardAInumuscom
                                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                         Case 0
                                                              If cardAInumcaseperson(i, 2, k) = 0 And Val(werp) < 3 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                         Case 1
                                                              If cardAInumcaseperson(i, 2, k) = 0 And Val(werp) < 3 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                    End Select
                                               Next
                                       End If
                                   End If
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================High hand
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
             End If
              '=================Jackpot
             If Pn2 = 0 Then
                     werp = 0
                     Erase weryu
                     If movecpre = 2 Then
                            If cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                '======================
                                werp = 0
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                              If Val(cardcountAInum(k, 2)) < 6 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                              End If
                                          Case 1
                                              If Val(cardcountAInum(k, 4)) < 6 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                              End If
                                     End Select
                                Next
                            End If
                    End If
            End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub
Sub 音音夢(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewer41 As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then
    livewer = liveus(角色人物對戰人數(1, 2))
    livewer41 = liveus41(角色人物對戰人數(1, 2))
Else
    livewer = livecom(角色人物對戰人數(2, 2))
    livewer41 = livecom41(角色人物對戰人數(1, 2))
End If
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
              Erase weryu
              weryu(1) = 999 '目的取最低HP數
            '=================愉快抽血
            If Pn3 = 0 Then
                    If cardAInumcaseperson(i, 1, 14) >= 2 Then
                        Select Case uscom
                             Case 1
                                   For k = 1 To 3
                                         If Val(liveus(角色待機人物紀錄數(1, k))) < Val(weryu(1)) And Val(liveus(角色待機人物紀錄數(1, k))) > 0 Then
                                             weryu(1) = liveus(角色待機人物紀錄數(1, k))
                                             weryu(2) = liveus41(角色待機人物紀錄數(1, k))
                                             weryu(3) = k
                                        End If
                                   Next
                             Case 2
                                   For k = 1 To 3
                                         If Val(livecom(角色待機人物紀錄數(2, k))) < Val(weryu(1)) And Val(livecom(角色待機人物紀錄數(2, k))) > 0 Then
                                             weryu(1) = livecom(角色待機人物紀錄數(2, k))
                                             weryu(2) = livecom41(角色待機人物紀錄數(2, k))
                                             weryu(3) = k
                                        End If
                                   Next
                        End Select
                        If cardAInumcaseperson(i, 1, 14) < Val(weryu(1)) Or _
                           (Val(weryu(1)) < Val(weryu(2)) And cardAInumcaseperson(i, 1, 14) >= 10 And weryu(3) <> 1) Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + Val(cardAInumcaseperson(i, 1, 14)) * 5
                        '======================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 2)) * 5
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 4)) * 5
                                               End If
                                     End Select
                                Next
                        ElseIf cardAInumcaseperson(i, 1, 14) >= Val(weryu(1)) Then '==組合下特卡數超過自身/待機成員最低血量時
                                cardAInumFinal(i, 1) = -10000
                        End If
                    End If
            End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================溫柔注射
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    weryu(1) = 999 '目的取最低HP數
                     If movecpre < 3 Then
                           If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 12) >= 2 Then
                               Select Case uscom
                                    Case 1
                                          For k = 2 To 3
                                                If Val(liveus(角色待機人物紀錄數(1, k))) < Val(weryu(1)) And Val(liveus(角色待機人物紀錄數(1, k))) > 0 Then
                                                    weryu(1) = liveus(角色待機人物紀錄數(1, k))
                                                    weryu(2) = k
                                               End If
                                          Next
                                    Case 2
                                          For k = 2 To 3
                                                If Val(livecom(角色待機人物紀錄數(2, k))) < Val(weryu(1)) And Val(livecom(角色待機人物紀錄數(2, k))) > 0 Then
                                                    weryu(1) = livecom(角色待機人物紀錄數(2, k))
                                                    weryu(2) = k
                                               End If
                                          Next
                               End Select
                               If Val(weryu(2)) <> 0 And livewer < Val(weryu(1)) Then
                               Else   '====除了待機成員血量高於自身血量外
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                       '======================
                                       werp = 0
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                      If cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                                      If cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               End If
                           End If
                     End If
             End If
              '=================秘密苦藥
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre > 1 Then
                           If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 Then
                               Select Case uscom
                                    Case 1
                                           weryu(1) = liveus(角色待機人物紀錄數(1, 2))
                                           weryu(2) = liveus(角色待機人物紀錄數(1, 3))
                                           weryu(3) = liveus41(角色待機人物紀錄數(1, 2))
                                           weryu(4) = liveus41(角色待機人物紀錄數(1, 3))
                                    Case 2
                                           weryu(1) = livecom(角色待機人物紀錄數(2, 2))
                                           weryu(2) = livecom(角色待機人物紀錄數(2, 3))
                                           weryu(3) = livecom41(角色待機人物紀錄數(2, 2))
                                           weryu(4) = livecom41(角色待機人物紀錄數(2, 3))
                               End Select
                               If (Val(weryu(1)) < Val(weryu(3)) And Val(weryu(1)) > 0) And _
                                  (Val(weryu(2)) < Val(weryu(4)) And Val(weryu(2)) > 0) And _
                                   Val(livewer) < Val(livewer41) Then
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                       '======================
                                       werp = 0
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                      If cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                      End If
                                                      If cardcountAInum(k, 1) = a4a And Val(weryu(3)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                                      If cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                      End If
                                                      If cardcountAInum(k, 3) = a4a And Val(weryu(3)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               Else
                                       cardAInumFinal(i, 1) = -100
                               End If
                           End If
                     End If
            End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================美味牛奶
             If Pn1 = 0 Then
                     werp = 0
                     Erase weryu
                     For k = 1 To cardAInumuscom
                          If cardAInumcaseperson(i, 2, k) > 0 Then
                              werp = Val(werp) + 1
                          End If
                     Next
                     Select Case uscom
                         Case 1
                                weryu(1) = liveus(角色待機人物紀錄數(1, 2))
                                weryu(2) = liveus(角色待機人物紀錄數(1, 3))
                         Case 2
                                weryu(1) = livecom(角色待機人物紀錄數(2, 2))
                                weryu(2) = livecom(角色待機人物紀錄數(2, 3))
                     End Select
                      If Val(werp) >= 2 And Val(weryu(1)) <> 1 And Val(weryu(2)) <> 1 Then
                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                          werp = 0
                          '======================
                            For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                         Case 1
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                    End Select
                               Next
                      ElseIf Val(werp) < 2 And Val(weryu(1)) <> 1 And Val(weryu(2)) <> 1 Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 2 Then
                                werp = 0
                                '==============1.針對已預定出牌的部分作加成
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.對數值為1的牌作加成
                                If werp < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.對數值為2的牌作加成
                                If Val(werp) < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 2 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                      Else
                            cardAInumFinal(i, 1) = -10000
                      End If
                End If
         Next
End Select

End Sub
Sub 艾依查庫(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewermax As Integer, weryu(1 To 3) As Integer  '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
If uscom = 1 Then livewermax = liveusmax(角色人物對戰人數(1, 2)) Else livewermax = livecommax(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================連射
                 If Pn1 = 0 Then
                        If movecpre > 1 Then
                               If cardAInumcasepersonTER(i, 5, 1) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================神速之劍
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 11) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + Int(Val(cardAInumcaseperson(i, 1, 11)) / 2 + 0.9)
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) > 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) > 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                                   '=============
                                   If cardAInumcasepersonTER(i, 1, 1) Mod 2 <> 0 Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      End If
                                            End Select
                                       Next
                                   ElseIf cardAInumcasepersonTER(i, 1, 1) > 0 Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 3) <> a2a _
                                                          And cardcountAInum(k, 2) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + 1
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 1) <> a2a _
                                                          And cardcountAInum(k, 4) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + 1
                                                      End If
                                            End Select
                                       Next
                                       If Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                           For k = 1 To cardAInumuscom
                                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                         Case 0
                                                              If cardcountAInum(k, 1) = a1a _
                                                                  And cardcountAInum(k, 2) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                         Case 1
                                                              If cardcountAInum(k, 3) = a1a _
                                                                  And cardcountAInum(k, 4) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                    End Select
                                               Next
                                        End If
                                   End If
                               End If
                        End If
                End If
                 '=================憤怒一擊
                 If Pn3 = 0 Then
                        werp = 0
                        wnm = 0
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                   wnm = (livewermax - livewer) * 2
                                   If wnm > 16 Then wnm = 16
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + wnm
                                   '======================
                                   For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              werp = Val(werp) + p
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              werp = Val(werp) + p
                                                          End If
                                                End Select
                                           Next
                                     Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================不屈之心
            If Pn4 = 0 Then
                    If cardAInumcasepersonTER(i, 2, 1) >= 2 Then
                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                        '======================
                        For k = 1 To cardAInumuscom
                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                  Case 0
                                       If cardcountAInum(k, 1) = a2a And cardcountAInum(k, 2) = 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       End If
                                  Case 1
                                       If cardcountAInum(k, 3) = a2a And cardcountAInum(k, 4) = 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       End If
                             End Select
                        Next
                    End If
            End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub

Sub 阿貝爾(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================霸王閃擊
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================閃電旋風刺
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 2 Then
                               If cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a3a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a3a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================幻影劍舞
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcasepersonTER(i, 1, 1) >= 1 And cardAInumcasepersonTER(i, 1, 2) >= 1 And cardAInumcasepersonTER(i, 1, 3) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================抽刀斷水計
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre = 1 Then
                           If cardAInumcaseperson(i, 1, 14) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                                 For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + p
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + p
                                                      End If
                                            End Select
                                       Next
                                 Next
                           End If
                    End If
            End If
         Next
End Select

End Sub
Sub 利恩(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                '=================劫影攻擊
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                         If cardAInumcaseperson(i, 1, 14) >= 1 Then
                             cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                             '======================
                               For k = 1 To cardAInumuscom
                                      Select Case Mid(cardAInumnm(i - 1), k, 1)
                                            Case 0
                                                 If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                            Case 1
                                                 If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                       End Select
                                  Next
                         End If
                 End If
                 '=================毒牙
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 14) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================背刺
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   Select Case uscom
                                        Case 1
                                              For k = 1 To 14
                                                     If 人物異常狀態資料庫(2, 角色人物對戰人數(2, 2), k, 3) = "BUFFN00801" Then
                                                         weryu(1) = Val(人物異常狀態資料庫(2, 角色人物對戰人數(2, 2), k, 2))
                                                     End If
                                               Next
                                        Case 2
                                               For k = 1 To 14
                                                     If 人物異常狀態資料庫(1, 角色人物對戰人數(1, 2), k, 3) = "BUFFN00801" Then
                                                         weryu(1) = Val(人物異常狀態資料庫(1, 角色人物對戰人數(1, 2), k, 2))
                                                     End If
                                               Next
                                   End Select
                                   If (weryu(1) >= 2 And 是否移動階段續估計判斷程序 = True) Or (weryu(1) >= 1 And 是否移動階段續估計判斷程序 = False) Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                          End If
                                                          If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                          End If
                                                          If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                          End If
                                                End Select
                                           Next
                                   End If
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================反擊的狼煙
            If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
             End If
        Next
    Case 3 '==移動階段類

End Select

End Sub
Sub 夏洛特(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewer41 As Integer, livewermax As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then
    livewer = liveus(角色人物對戰人數(1, 2))
    livewer41 = liveus41(角色人物對戰人數(1, 2))
    livewermax = liveusmax(角色人物對戰人數(1, 2))
Else
    livewer = livecom(角色人物對戰人數(2, 2))
    livewer41 = livecom41(角色人物對戰人數(1, 2))
    livewermax = livecommax(角色人物對戰人數(2, 2))
End If
If 夏洛特_階段處理記錄數(2) = BattleTurn Then
    夏洛特_階段處理記錄數(1) = 0
End If
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================冬之夢
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   If 夏洛特_階段處理記錄數(1) = 0 Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                          End If
                                                          If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                          End If
                                                          If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                          End If
                                                End Select
                                           Next
                                   End If
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================大聖堂
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    werp = Val(werp) + p
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    werp = Val(werp) + p
                                                End If
                                      End Select
                                 Next
                           Next
                     End If
            End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================夜未央
              If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   If 夏洛特_階段處理記錄數(1) <> 3 Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                           If livewer <= livewer41 Then
                                               夏洛特_階段處理記錄數(1) = 2
                                               夏洛特_階段處理記錄數(2) = BattleTurn + 2
                                           End If
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a2a And weryu(1) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                          End If
                                                          If cardcountAInum(k, 1) = a3a And weryu(2) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a2a And weryu(1) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                          End If
                                                          If cardcountAInum(k, 3) = a3a And weryu(2) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                          End If
                                                End Select
                                           Next
                                     End If
                               End If
                        End If
                 End If
                 '=================幸福的理由
                 If Pn4 = 0 Then
                         werp = 0
                         Erase weryu
                         If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                If 夏洛特_階段處理記錄數(1) = 0 And 夏洛特_階段處理記錄數(3) = 0 Then
                                        Select Case uscom
                                             Case 1
                                                  For k = 2 To 3
                                                       If liveus(角色待機人物紀錄數(1, k)) <= 0 Then
                                                           werp = Val(werp) + 1
                                                       End If
                                                  Next
                                             Case 2
                                                  For k = 2 To 3
                                                       If livecom(角色待機人物紀錄數(2, k)) <= 0 Then
                                                           werp = Val(werp) + 1
                                                       End If
                                                  Next
                                        End Select
                                        If werp = 2 And livewer + 1 >= livewermax Then
                                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                                夏洛特_階段處理記錄數(1) = 3
                                                夏洛特_階段處理記錄數(2) = BattleTurn + 2
                                                夏洛特_階段處理記錄數(3) = 1
                                                '======================
                                                For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                                      For k = 1 To cardAInumuscom
                                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                                Case 0
                                                                     If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                                         cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                         werp = Val(werp) + p
                                                                     End If
                                                                Case 1
                                                                     If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                                         cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                         werp = Val(werp) + p
                                                                     End If
                                                           End Select
                                                      Next
                                                Next
                                        End If
                                End If
                        End If
                End If
         Next
End Select

End Sub
Sub 泰瑞爾(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================Rud-913
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a3a And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a3a And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================Chr-799
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================Wil-846
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================Von-541
            If Pn2 = 0 Then
                werp = 0
                Erase weryu
                If cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 Then
                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                    '======================
                    For k = 1 To cardAInumuscom
                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                              Case 0
                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                   End If
                                   If cardcountAInum(k, 1) = a2a And weryu(2) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                   End If
                              Case 1
                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                   End If
                                   If cardcountAInum(k, 3) = a2a And weryu(2) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                   End If
                         End Select
                    Next
                End If
            End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub
Sub 瑪格莉特(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewermax As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then
    livewer = liveus(角色人物對戰人數(1, 2))
    livewermax = liveusmax(角色人物對戰人數(1, 2))
Else
    livewer = livecom(角色人物對戰人數(2, 2))
    livewermax = livecommax(角色人物對戰人數(2, 2))
End If
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================月光
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================地獄獵心獸
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (Val(cardAInumcaseperson(i, 1, 11)) + Val(cardAInumcaseperson(i, 1, 15)))
                               weryu(2) = ((Val(cardAInumcaseperson(i, 1, 11)) + Val(cardAInumcaseperson(i, 1, 15))) \ 5) * 5
                               '======================
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a1a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a5a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a1a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a5a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                 End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================恍惚
            If Pn2 = 0 Then
                 werp = 0
                 Erase weryu
                 If movecpre = 1 Then
                        If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                            cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            '======================
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 5)) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                           End If
                                      Case 1
                                           If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 5)) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                           End If
                                 End Select
                            Next
                        End If
                 End If
            End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================末日幻影
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 13) = 0 Then
                           cardAInumFinal(i, 1) = 1
                           '======================
                             For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 1)) And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                               If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 9)) And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 1)) And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                               If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 9)) And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                             Next
                     End If
              End If
         Next
End Select

End Sub
Sub 庫勒尼西(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================深淵
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 14) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10 + Int(Val(cardAInumcaseperson(i, 1, 14)) / 2 + 0.9)
                               '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                           End If
                 End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================瘋狂眼窩
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre = 1 Then
                           If cardAInumcaseperson(i, 1, 14) >= 1 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                    End If
             End If
              '=================黑暗漩渦
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 3)) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 5)) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 3)) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 5)) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
              End If
              '=================沙漠中的海市蜃樓
              If Pn1 = 0 Then
                    If movecpre = 3 Then
                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                    End If
              End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub
Sub 蕾格烈芙(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================C.T.L
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================B.P.A
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================L.A.R
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
              End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=====================S.S.S
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                    If cardAInumcasepersonTER(i, 4, 1) >= 1 And cardAInumcasepersonTER(i, 4, 2) >= 1 And cardAInumcasepersonTER(i, 4, 3) >= 1 Then
                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                          '=====================
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 1 And Val(weryu(1)) < 1) Or _
                                               (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 2 And Val(weryu(2)) < 1) Or _
                                               (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 And Val(weryu(3)) < 1) Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(cardcountAInum(k, 2)) = weryu(cardcountAInum(k, 2)) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 1 And Val(weryu(1)) < 1) Or _
                                               (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 2 And Val(weryu(2)) < 1) Or _
                                               (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 And Val(weryu(3)) < 1) Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(cardcountAInum(k, 4)) = weryu(cardcountAInum(k, 4)) + 1
                                           End If
                                 End Select
                            Next
                    End If
             End If
         Next
End Select

End Sub
Sub 多妮妲(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================殘虐傾向
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 14) >= 2 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               Next
                           End If
                End If
             '=================律死擊
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre = 1 Then
                           If cardAInumcaseperson(i, 1, 11) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                                 For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                         For k = 1 To cardAInumuscom
                                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                   Case 0
                                                        If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                        End If
                                                   Case 1
                                                        If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                        End If
                                              End Select
                                         Next
                                 Next
                           End If
                     End If
             End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
                '===========異質者
                If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcasepersonTER(i, 4, 3) >= 1 Then
                              If (攻擊防禦骰子總數(uscom - 1) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And 是否移動階段續估計判斷程序 = False And uscom = 2) Or _
                                 (攻擊防禦骰子總數(uscom + 1) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And 是否移動階段續估計判斷程序 = False And uscom = 1) Or _
                                 (是否移動階段續估計判斷程序 = True And Val(livewer) <= 3) Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10000
                              '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                         End Select
                                    Next
                              End If
                        End If
                End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================超級女主角
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                         Select Case uscom
                              Case 1
                                   If 執行動作_檢查是否有指定異常狀態(1, 7) = True Then
                                       werp = 1
                                   End If
                                   If 執行動作_檢查是否有指定異常狀態(1, 8) = True Then
                                       werp = 1
                                   End If
                                   If 執行動作_檢查是否有指定異常狀態(1, 9) = True Then
                                       werp = 1
                                   End If
                              Case 2
                                   If 執行動作_檢查是否有指定異常狀態(2, 1) = True Then
                                       werp = 1
                                   End If
                                   If 執行動作_檢查是否有指定異常狀態(2, 2) = True Then
                                       werp = 1
                                   End If
                                   If 執行動作_檢查是否有指定異常狀態(2, 3) = True Then
                                       werp = 1
                                   End If
                         End Select
                         If werp = 0 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                               '======================
                                 For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                   End If
                                                   If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                   End If
                                                   If cardcountAInum(k, 1) = a4a And weryu(3) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                   End If
                                                   If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                   End If
                                                   If cardcountAInum(k, 3) = a4a And weryu(3) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                                   End If
                                         End Select
                                    Next
                           End If
                     End If
                End If
         Next
End Select

End Sub
Sub 傑多(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================因果之幻
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcaseperson(i, 1, 13) >= 1 Then
                            cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            '======================
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                           End If
                                      Case 1
                                           If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                           End If
                                 End Select
                            Next
                        End If
                 End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================因果之輪
              If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
              End If
              '=================因果之刻
              If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 4 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 4 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 4 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
              End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================因果之線
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
             End If
         Next
End Select

End Sub
Sub 阿奇波爾多(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================致命槍擊
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcasepersonTER(i, 5, 1) >= 1 And cardAInumcasepersonTER(i, 5, 2) >= 1 And cardAInumcasepersonTER(i, 5, 3) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================劫影攻擊
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                         If cardAInumcaseperson(i, 1, 14) >= 1 Then
                             cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                             '======================
                               For k = 1 To cardAInumuscom
                                      Select Case Mid(cardAInumnm(i - 1), k, 1)
                                            Case 0
                                                 If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                            Case 1
                                                 If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                       End Select
                                  Next
                         End If
                 End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================防護射擊
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 15) >= 1 And movecpre > 1 Then
                         If ((攻擊防禦骰子總數(uscom) >= 30 Or cardAInumcaseperson(i, 1, 9) = 1) And 是否移動階段續估計判斷程序 = False) Or _
                             是否移動階段續估計判斷程序 = True Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                '======================
                                If 攻擊防禦骰子總數(uscom) >= 30 Then
                                    werp = Int((攻擊防禦骰子總數(uscom) - 30) / 2 + 0.9)
                                Else
                                    werp = 0
                                End If
                                If werp > 0 Then
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a5a And Val(weryu(1)) < werp Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a5a And Val(weryu(1)) < werp Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                Else
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                End If
                         End If
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================大地崩壞
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 3 And movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
            End If
         Next
End Select

End Sub
Sub 露緹亞(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================腐朽之靈
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================渦騎劍閃
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10 + cardAInumcaseperson(i, 1, 11) * 5
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 2) * 5
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 4) * 5
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================朦朧之暗
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 And _
                         movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================暗影之翼
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 And _
                         movecpre < 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 1) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 1) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub
Sub 梅莉(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================夢幻魔杖
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================夢境搖籃
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                                   If livewer <= 2 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 95
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================徬徨夢羽
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 12) >= 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================綿羊幻夢
             If Pn3 = 0 And movecpre < 3 Then
                     werp = 0
                     Erase weryu
                     For k = 1 To cardAInumuscom
                          If cardAInumcaseperson(i, 2, k) > 0 Then
                              werp = Val(werp) + 1
                          End If
                     Next
                      If Val(werp) >= 2 And Val(livewer) >= 5 Then
                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                          werp = 0
                          '======================
                            For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                         Case 1
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                    End Select
                               Next
                      ElseIf Val(werp) < 2 And (cardAInumcaseperson(i, 1, 13) < 2 Or Val(livewer) >= 5) Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 2 Then
                                werp = 0
                                '==============1.針對已預定出牌的部分作加成
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.對數值為1的牌作加成
                                If werp < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.對數值為2的牌作加成
                                If Val(werp) < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 2 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                      Else
                            cardAInumFinal(i, 1) = -10
                      End If
                End If
         Next
End Select

End Sub
Sub 貝琳達(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================裂地冰牙
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                '=====================溶魂之雨
                If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                                If cardAInumcaseperson(i, 1, 11) >= 1 And _
                                    cardAInumcaseperson(i, 1, 13) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 Then
                                      cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                      If cardAInumuscom >= 10 Then
                                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                      End If
                                      '=====================
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If (cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 1) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 1) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 1) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 1) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 1) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                                  Case 1
                                                       If (cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 3) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 3) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 3) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 3) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 3) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                             End Select
                                        Next
                                End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================水晶幻鏡
             If Pn2 = 0 And movecpre < 3 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================雪光
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         Select Case uscom
                              Case 1
                                   If livecom(角色人物對戰人數(2, 2)) = livecommax(角色人物對戰人數(2, 2)) Then
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   End If
                              Case 2
                                   If liveus(角色人物對戰人數(1, 2)) = liveusmax(角色人物對戰人數(1, 2)) Then
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   End If
                         End Select
                         '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               Next
                     End If
            End If
         Next
End Select

End Sub
Sub 蕾(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================輪旋曲-琉璃色的微風
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 4 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                '=================Ex輪旋曲-琉璃色的微風
                 If Pn1 = 1 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 5 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================終曲-無盡輪迴的終結
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 4 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                    For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                            For k = 1 To cardAInumuscom
                                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                      Case 0
                                                           If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 4 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                           End If
                                                      Case 1
                                                           If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 4 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                           End If
                                                 End Select
                                            Next
                                    Next
                               End If
                        End If
                End If
                '=================Ex終曲-無盡輪迴的終結
                 If Pn4 = 1 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 6 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                    For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                            For k = 1 To cardAInumuscom
                                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                      Case 0
                                                           If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 6 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                           End If
                                                      Case 1
                                                           If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 6 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                           End If
                                                 End Select
                                            Next
                                    Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================協奏曲-加百烈的守護
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================Ex協奏曲-加百烈的守護
             If Pn2 = 1 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 3 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 3 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '===========安魂曲-死神的鎮魂歌/Ex安魂曲-死神的鎮魂歌
                If Pn3 = 0 Or Pn3 = 1 Then
                        If cardAInumcasepersonTER(i, 4, 3) >= 1 Then
                              If (攻擊防禦骰子總數(uscom) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And 是否移動階段續估計判斷程序 = False) Or _
                                 (是否移動階段續估計判斷程序 = True And Val(livewer) <= 3) Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                              '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       Exit For
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       Exit For
                                                   End If
                                         End Select
                                    Next
                              End If
                        End If
                End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub
Sub 羅莎琳(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================染血之刃
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                '=================Ex染血之刃
                 If Pn2 = 1 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 13) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For p = Val(cardAInumcaseperson(i, 1, 6)) To Val(cardAInumcaseperson(i, 1, 5)) Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                    Next
                               End If
                        End If
                End If
                 '=================黑霧的纏繞
                 If Pn4 = 0 Then
                       werp = 0
                       Erase weryu
                       If cardAInumcaseperson(i, 1, 14) >= 2 Then
                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                           '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                   End If
                                         End Select
                                    Next
                            Next
                       End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================黑霧幻影
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 2 And movecpre = 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                     End If
             End If
             '=================Ex黑霧幻影
             If Pn1 = 1 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 And movecpre = 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                     End If
             End If
             '=================咀咒的刻印
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 5 And cardAInumcaseperson(i, 1, 14) >= 1 And movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                           End If
                                      Case 1
                                           If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                           End If
                                 End Select
                            Next
                     End If
             End If
        Next
    Case 3 '==移動階段類
        
End Select

End Sub
Sub 洛洛妮(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewermax As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then
    livewer = liveus(角色人物對戰人數(1, 2))
    livewermax = liveusmax(角色人物對戰人數(1, 2))
Else
    livewer = livecom(角色人物對戰人數(2, 2))
    livewermax = livecommax(角色人物對戰人數(2, 2))
End If
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================砲擊壓制
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                         For k = 1 To cardAInumuscom
                                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                   Case 0
                                                        If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                        End If
                                                   Case 1
                                                        If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                        End If
                                              End Select
                                         Next
                                    Next
                               End If
                        End If
                End If
                 '=================貪婪之刃與嗜血之槍
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 5 And cardAInumcaseperson(i, 1, 15) >= 5 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================風暴感知
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================逆轉戰局的槍響
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 3 And movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (livewermax - livewer) * 2
                         '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                End If
                                      End Select
                                 Next
                            Next
                     End If
            End If
         Next
End Select

End Sub
Sub 克頓(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================隱蔽射擊
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 2 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================竊取資料
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                End If
                                      End Select
                                 Next
                            Next
                     End If
             End If
             '=================逃亡計畫
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    If cardAInumcasepersonTER(i, 2, 1) >= 1 And cardAInumcasepersonTER(i, 4, 1) >= 1 Then
                        Select Case uscom
                            Case 1
                                   weryu(1) = liveus(角色待機人物紀錄數(1, 2))
                                   weryu(2) = liveus(角色待機人物紀錄數(1, 3))
                            Case 2
                                   weryu(1) = livecom(角色待機人物紀錄數(2, 2))
                                   weryu(2) = livecom(角色待機人物紀錄數(2, 3))
                        End Select
                        If 攻擊防禦骰子總數(uscom) >= 30 And weryu(1) > 3 And weryu(2) > 3 And 是否移動階段續估計判斷程序 = False Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                '======================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a2a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 1 And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a2a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 1 And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Else
                                cardAInumFinal(i, 1) = 0
                        End If
                    End If
            End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================惡意情報
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                    If cardAInumcasepersonTER(i, 1, 3) >= 1 And cardAInumcasepersonTER(i, 5, 3) >= 1 And cardAInumcaseperson(i, 1, 14) >= 2 And movecpre > 1 Then
                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                        '======================
                        For k = 1 To cardAInumuscom
                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                  Case 0
                                       If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 3 And weryu(1) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(1) = Val(weryu(1)) + 1
                                       End If
                                       If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 3 And weryu(2) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(2) = Val(weryu(2)) + 1
                                       End If
                                  Case 1
                                       If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 3 And weryu(1) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(1) = Val(weryu(1)) + 1
                                       End If
                                       If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 3 And weryu(2) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(2) = Val(weryu(2)) + 1
                                       End If
                             End Select
                        Next
                        For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                             For k = 1 To cardAInumuscom
                                 Select Case Mid(cardAInumnm(i - 1), k, 1)
                                       Case 0
                                            If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(3)) < 2 Then
                                                cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                            End If
                                       Case 1
                                            If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(3)) < 2 Then
                                                cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                            End If
                                  End Select
                             Next
                        Next
                    End If
            End If
         Next
End Select

End Sub
Sub 艾蕾可(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================聖王威光
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcaseperson(i, 1, 14) >= 3 Then
                            cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                            '======================
                            For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                End If
                                      End Select
                                 Next
                            Next
                        End If
                End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================王座之炎
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 5 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 5 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 5 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================白百合
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    For k = 1 To cardAInumuscom
                          If cardAInumcaseperson(i, 2, k) > 0 Then
                              werp = Val(werp) + 1
                          End If
                    Next
                     If Val(werp) >= 2 And movecpre < 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                     End If
             End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================救濟天使
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 5 Then
                         Select Case uscom
                                Case 1
                                     For k = 2 To 3
                                            For p = 1 To 14
                                               If 人物異常狀態資料庫(1, 角色待機人物紀錄數(1, k), p, 3) = "BUFFS00101" Then
                                                   werp = 1
                                               End If
                                            Next
                                     Next
                                Case 2
                                     For k = 2 To 3
                                            For p = 1 To 14
                                                If 人物異常狀態資料庫(2, 角色待機人物紀錄數(2, k), p, 3) = "BUFFS00101" Then
                                                    werp = 1
                                                End If
                                            Next
                                    Next
                         End Select
                         Select Case uscom
                             Case 1
                                    weryu(1) = liveus(角色待機人物紀錄數(1, 2))
                                    weryu(2) = liveus(角色待機人物紀錄數(1, 3))
                                    weryu(3) = liveus41(角色待機人物紀錄數(1, 2))
                                    weryu(4) = liveus41(角色待機人物紀錄數(1, 3))
                             Case 2
                                    weryu(1) = livecom(角色待機人物紀錄數(2, 2))
                                    weryu(2) = livecom(角色待機人物紀錄數(2, 3))
                                    weryu(3) = livecom41(角色待機人物紀錄數(2, 2))
                                    weryu(4) = livecom41(角色待機人物紀錄數(2, 3))
                        End Select
                         If (weryu(1) <= weryu(3) And weryu(1) > 0 And weryu(2) <= weryu(4) And weryu(2) > 0) Or _
                             ((執行動作_檢查是否有指定異常狀態(uscom, 37) = False And uscom = 1 And weryu(1) = 0 And weryu(2) = 0) Or _
                             (執行動作_檢查是否有指定異常狀態(uscom, 38) = False And uscom = 2 And weryu(1) = 0 And weryu(2) = 0)) Or _
                             (werp = 0 And (weryu(1) > 0 Or weryu(2) > 0)) Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                '======================
                                  For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 5 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 5 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                   Next
                         End If
                     End If
            End If
         Next
End Select

End Sub
Sub 尤莉卡(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '暫時變數
If uscom = 1 Then livewer = liveus(角色人物對戰人數(1, 2)) Else livewer = livecom(角色人物對戰人數(2, 2))
Select Case turn
    Case 1 '==攻擊階段類
          For i = 1 To 2 ^ cardAInumuscom
                 '=================奸佞的鐵鎚
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
        Next
    Case 2 '==防禦階段類
        For i = 1 To 2 ^ cardAInumuscom
            '=================不善的信仰
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================曲惡的安寧
                If Pn3 = 0 Then
                       werp = 0
                       Erase weryu
                       If movecpre = 3 Then
                              If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '======================
                                  For k = 1 To cardAInumuscom
                                      Select Case Mid(cardAInumnm(i - 1), k, 1)
                                            Case 0
                                                 If cardcountAInum(k, 1) = a2a And weryu(1) < 3 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                 End If
                                                 If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                 End If
                                            Case 1
                                                 If cardcountAInum(k, 3) = a2a And weryu(1) < 3 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                 End If
                                                 If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                 End If
                                       End Select
                                  Next
                              End If
                       End If
                End If
        Next
    Case 3 '==移動階段類
        For i = 1 To 2 ^ cardAInumuscom
             '=================超載
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
            End If
         Next
End Select

End Sub

