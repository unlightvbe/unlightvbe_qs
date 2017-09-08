Attribute VB_Name = "AI人物"
Sub 全人物通用(ByVal n As Integer)
Dim ay As Integer
Select Case n
    Case 1
        '================異常狀態-MOV減-有效移動值判斷處理
        For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
            If 人物異常狀態資料庫(2, i, 3) = 6 Then
                ay = ay + 人物異常狀態資料庫(2, i, 1)
            End If
            If 人物異常狀態資料庫(2, i, 3) = 17 Then
                ay = 99
                Exit For
            End If
        Next
        If 目前數(25) <= Val(ay) Then
            For i = 1 To 106
               If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 1 Then
                  pagecardnum(i, 11) = 0
               End If
            Next
        End If
    Case 2
        '================在最後棄牌階段判斷處理
        For j = 1 To 106
           If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
             pagecardnum(j, 11) = 1
             If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "雪莉" _
                Or FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾妲" _
                Or FormMainMode.compi1(角色人物對戰人數(2, 2)) = "音音夢" Then
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
             ElseIf FormMainMode.compi1(角色人物對戰人數(2, 2)) = "瑪格莉特" _
                 Or FormMainMode.compi1(角色人物對戰人數(2, 2)) = "C.C." _
                 Or FormMainMode.compi1(角色人物對戰人數(2, 2)) = "帕茉" _
                 Or FormMainMode.compi1(角色人物對戰人數(2, 2)) = "露緹亞" _
                 Or FormMainMode.compi1(角色人物對戰人數(2, 2)) = "傑多" Then
                       If pagecardnum(j, 1) = a4a Then
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
             Else
                  If pagecardnum(j, 3) = a4a Then
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
             If pagecardnum(j, 3) = a3a And pagecardnum(j, 4) = 1 Then '轉移動牌(試驗)
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
End Select
End Sub
Sub 艾依查庫(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾依查庫" Then
    Select Case n
        Case 1
            '===========依據距離出牌
            Select Case movecp
                Case 1
                    For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                        End If
                      Next
                Case Is > 1
                      For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a1a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a1a And Val(pagecardnum(j, 2)) >= 2 Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                         End If
                      Next
            End Select
        Case 2
        
    End Select
End If
End Sub
Sub 艾伯李斯特(ByVal n As Integer)
Dim aw As Integer
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾伯李斯特" Then
    Select Case n
        Case 1
            If movecp = 1 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 2 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
            End If
        Case 2
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) = 1 Then
                       aw = Val(aw) + 1
                   End If
            Next
            If aw = 2 Then
                For j = 1 To 106
                       If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                               pagecardnum(j, 11) = 1
                               If pagecardnum(j, 3) = a3a Then
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
                                Exit For
                           End If
                       End If
                Next
            ElseIf aw < 2 Then
                aw = 0
                For j = 1 To 106
                    If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                        If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                            aw = Val(aw) + 1
                        End If
                    End If
                Next
                If Val(aw) >= 3 Then
                    aw = 0
                    For j = 1 To 106
                       If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                               pagecardnum(j, 11) = 1
                               aw = Val(aw) + 1
                               If pagecardnum(j, 3) = a3a Then
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
                       End If
                       If Val(aw) >= 3 Then Exit For
                    Next
                End If
            End If
    End Select
End If
End Sub
Sub 史塔夏(ByVal n As Integer)
Dim aw(1 To 2) As Integer
Dim ae As Integer
Dim num(1 To 2, 1 To 2) As Integer '選擇人物暫時變數
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "史塔夏" Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If movecp = 3 And Val(aw(2)) >= 9 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a1a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a1a Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                        End If
                 Next
            End If
        Case 2
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
               aw(1) = 0
           Else
               aw(1) = 1
           End If
            '===============================
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If aw(1) = 1 And Val(aw(2)) >= 9 Then
                For j = 1 To 106
                       If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a3a And pagecardnum(j, 3) <> a1a Then
                                pagecardnum(j, 11) = 1
                             ElseIf pagecardnum(j, 3) = a3a And pagecardnum(j, 1) <> a1a Then
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
                                pagecardnum(j, 11) = 1
                             End If
                       End If
                Next
            Else
               For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) = 1 Then
                       ae = Val(ae) + 1
                   End If
                Next
                If ae = 2 Then
                    For j = 1 To 106
                           If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                               If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                                   pagecardnum(j, 11) = 1
                                   If pagecardnum(j, 3) = a3a Then
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
                                    Exit For
                               End If
                           End If
                    Next
                End If
            End If
    End Select
End If
End Sub
Sub CC(ByVal n As Integer)
Dim aw As Integer
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "C.C." Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If movecp = 1 And Val(aw) >= 6 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                        End If
                 Next
            End If
        Case 2
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                        If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) = 1 And Val(pagecardnum(j, 4)) <= 3 Then
                            pagecardnum(j, 11) = 1
                            Exit For
                        ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) = 1 And Val(pagecardnum(j, 2)) <= 3 Then
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
                             pagecardnum(j, 11) = 1
                             Exit For
                        End If
                   End If
            Next
    End Select
End If
End Sub
Sub 梅倫(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "梅倫" Then
    Select Case n
        Case 1
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 2 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub 利恩(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "利恩" Then
    Select Case n
        Case 1
            If movecp = 1 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 3 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 3 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
            End If
    End Select
End If
End Sub
Sub 夏洛特(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "夏洛特" Then
    Select Case n
        Case 1
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 2 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub 庫勒尼西(ByVal n As Integer)
Dim aw As Integer
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "庫勒尼西" Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 Then
                       If pagecardnum(j, 1) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If Val(aw) >= 3 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                        End If
                 Next
            End If
    End Select
End If
End Sub
Sub 蕾格烈芙(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "蕾格烈芙" Then
    Select Case n
        Case 1
            If movecp = 1 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 3 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 3 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
              Else
                 For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
             End If
    End Select
End If
End Sub
Sub 多妮妲(ByVal n As Integer)
Dim aw As Integer
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "多妮妲" Then
    Select Case n
        Case 1
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) <> 3 Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) <> 3 Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub 阿奇波爾多(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "阿奇波爾多" Then
    Select Case n
        Case 1
            If movecp > 1 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 3 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 3 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
              End If
          Case 2
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 1 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 1 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub 瑪格莉特(ByVal n As Integer)
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "瑪格莉特" Then
    Select Case n
        Case 1
             If FormMainMode.comaiatk(1).Caption = "月光" And movecp < 3 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 1 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 1 Then
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
                                 pagecardnum(j, 11) = 1
                                 Exit For
                            End If
                        End If
                 Next
             End If
    End Select
End If
End Sub
Sub 艾蕾可(ByVal n As Integer)
Dim aw As Integer
If FormMainMode.compi1(角色人物對戰人數(2, 2)) = "艾蕾可" Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a1a Then
                           aw = Val(aw) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a1a Then
                           aw = Val(aw) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If movecp > 1 And Val(aw) >= 5 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a1a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a1a Then
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
                                 pagecardnum(j, 11) = 1
                            End If
                        End If
                 Next
            End If
    End Select
End If
End Sub

