Attribute VB_Name = "異常狀態"
Public 異常狀態_混沌紀錄數(1 To 4) As Integer '異常狀態-混沌-骰量紀錄暫時變數(1.紀錄數值(原始)/2.紀錄數值(變更後)/3.數值紀錄是否啟動/4.攻擊防禦模式階段數)
Public 異常狀態_AI_混沌紀錄數(1 To 4) As Integer '異常狀態-AI-混沌-骰量紀錄暫時變數(1.紀錄數值(原始)/2.紀錄數值(變更後)/3.數值紀錄是否啟動/4.攻擊防禦模式階段數)
Sub ATK加_使用者()
Select Case 異常狀態檢查數(7, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 7 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 人物異常狀態資料庫(1, i, 1)
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 人物異常狀態資料庫(1, i, 1)
       End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 7 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(7, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(7, 1) = 1
        End If
      End If
     Next
   Case 3
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 7 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 人物異常狀態資料庫(1, i, 1)
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 人物異常狀態資料庫(1, i, 1)
        異常狀態檢查數(7, 1) = 1
        Exit For
       End If
     Next
End Select
End Sub
Sub ATK加_電腦()
Select Case 異常狀態檢查數(1, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 1 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 人物異常狀態資料庫(2, i, 1)
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + 人物異常狀態資料庫(2, i, 1)
       End If
     Next
'     formmainmode.trgoi2.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 1 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(1, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(1, 1) = 1
        End If
      End If
     Next
   Case 3
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 1 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 人物異常狀態資料庫(2, i, 1)
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - 人物異常狀態資料庫(2, i, 1)
        異常狀態檢查數(1, 1) = 1
        Exit For
       End If
     Next
End Select
End Sub
Sub ATK減_使用者()
Select Case 異常狀態檢查數(10, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 10 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 人物異常狀態資料庫(1, i, 1)
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 人物異常狀態資料庫(1, i, 1)
       End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 10 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(10, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(10, 1) = 1
        End If
      End If
     Next
   Case 3
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 10 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 人物異常狀態資料庫(1, i, 1)
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 人物異常狀態資料庫(1, i, 1)
        異常狀態檢查數(10, 1) = 1
        Exit For
       End If
     Next
End Select
End Sub
Sub ATK減_電腦()
Select Case 異常狀態檢查數(4, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 4 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 人物異常狀態資料庫(2, i, 1)
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - 人物異常狀態資料庫(2, i, 1)
        異常狀態檢查數(4, 1) = 2
       End If
     Next
'     formmainmode.trgoi2.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 4 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(4, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(4, 1) = 1
        End If
      End If
     Next
    Case 3
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(1, i, 3) = 4 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 人物異常狀態資料庫(2, i, 1)
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + 人物異常狀態資料庫(2, i, 1)
        異常狀態檢查數(4, 1) = 1
        Exit For
       End If
     Next
End Select
End Sub
Sub DEF加_使用者()
Select Case 異常狀態檢查數(8, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 8 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 人物異常狀態資料庫(1, i, 1)
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 人物異常狀態資料庫(1, i, 1)
        異常狀態檢查數(8, 1) = 2
       End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 8 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(8, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(8, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub DEF減_使用者()
Select Case 異常狀態檢查數(11, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 11 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 人物異常狀態資料庫(1, i, 1)
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 人物異常狀態資料庫(1, i, 1)
        異常狀態檢查數(11, 1) = 2
       End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 11 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(11, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(11, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub DEF加_電腦()
Select Case 異常狀態檢查數(2, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 2 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 人物異常狀態資料庫(2, i, 1)
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + 人物異常狀態資料庫(2, i, 1)
        異常狀態檢查數(2, 1) = 2
       End If
     Next
     FormMainMode.trgoi2.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 2 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(2, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(2, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub DEF減_電腦()
Select Case 異常狀態檢查數(5, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 5 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 人物異常狀態資料庫(2, i, 1)
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - 人物異常狀態資料庫(2, i, 1)
        異常狀態檢查數(5, 1) = 2
       End If
     Next
     FormMainMode.trgoi2.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 5 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(5, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(5, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub MOV加_使用者()
Select Case 異常狀態檢查數(9, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 9 Then
            moveus = moveus + 人物異常狀態資料庫(1, i, 1)
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 9 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(9, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(9, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub MOV減_使用者()
Select Case 異常狀態檢查數(12, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 12 Then
           If moveus > 0 Then
               moveus = moveus - 人物異常狀態資料庫(1, i, 1)
               If moveus < 0 Then moveus = 0
           End If
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 12 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(12, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(12, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub MOV加_電腦()
Select Case 異常狀態檢查數(3, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 3 Then
           movecom = movecom + 人物異常狀態資料庫(2, i, 1)
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 3 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(3, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(3, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub MOV減_電腦()
Select Case 異常狀態檢查數(6, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 6 Then
           If movecom > 0 Then
               movecom = movecom - 人物異常狀態資料庫(2, i, 1)
               If movecom < 0 Then
                   movecom = 0
                   movecheckcom = 0
                End If
           End If
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 6 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(6, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(6, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub 不死_使用者()
Select Case 異常狀態檢查數(14, 1)
    Case 1
        For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
          If 人物異常狀態資料庫(1, i, 3) = 14 Then
             擲骰表單溝通暫時變數(2) = 0
             擲骰後骰傷害數 = 0
          End If
        Next
    Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 14 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(14, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 不死_電腦()
Select Case 異常狀態檢查數(18, 1)
    Case 1
        For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
          If 人物異常狀態資料庫(2, i, 3) = 18 Then
             擲骰表單溝通暫時變數(2) = 0
             擲骰後骰傷害數 = 0
          End If
        Next
    Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 18 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(18, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 中毒_使用者()
Select Case 異常狀態檢查數(20, 1)
    Case 1
        For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
          If 人物異常狀態資料庫(1, i, 3) = 20 Then
            人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
            傷害執行_使用者 (1)
            If 人物異常狀態資料庫(1, i, 2) = 0 Then
              '===繼承下一狀態資料
               戰鬥系統類.異常狀態繼承_使用者
               異常狀態檢查數(21, 2) = 0
           Else
               FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
           End If
         End If
        Next
End Select
End Sub
Sub 中毒_電腦()
Select Case 異常狀態檢查數(21, 1)
    Case 1
        For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
          If 人物異常狀態資料庫(2, i, 3) = 21 Then
            人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
            傷害執行_電腦 (1)
            If 人物異常狀態資料庫(2, i, 2) = 0 Then
              '===繼承下一狀態資料
               戰鬥系統類.異常狀態繼承_電腦
               異常狀態檢查數(21, 2) = 0
           Else
               FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
           End If
         End If
        Next
    Case 2
        movecom = 0
        movecheckcom = 0
End Select
End Sub
Sub 自壞_使用者()
Select Case 異常狀態檢查數(15, 1)
    Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 15 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
            戰鬥系統類.傷害執行_立即死亡_使用者 1
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(15, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 自壞_電腦()
Select Case 異常狀態檢查數(19, 1)
    Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 19 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
            戰鬥系統類.傷害執行_立即死亡_電腦 1
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(19, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 封印_使用者()
Select Case 異常狀態檢查數(22, 1)
    Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 22 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(22, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 封印_電腦()
Select Case 異常狀態檢查數(23, 1)
    Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 23 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(23, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 能力低下_使用者()
Select Case 異常狀態檢查數(24, 1)
  Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 24 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 人物異常狀態資料庫(1, i, 2) * 1
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 人物異常狀態資料庫(1, i, 2) * 1
        Exit For
       End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 24 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 人物異常狀態資料庫(1, i, 2) * 1
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 人物異常狀態資料庫(1, i, 2) * 1
        Exit For
       End If
     Next
End Select
End Sub
Sub 能力低下_電腦()
Select Case 異常狀態檢查數(25, 1)
  Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 25 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 人物異常狀態資料庫(2, i, 2) * 1
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - 人物異常狀態資料庫(2, i, 2) * 1
        Exit For
       End If
     Next
'     formmainmode.trgoi2.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 25 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 人物異常狀態資料庫(2, i, 2) * 1
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + 人物異常狀態資料庫(2, i, 2) * 1
        Exit For
       End If
     Next
End Select
End Sub
Sub 麻痺_使用者()
Select Case 異常狀態檢查數(16, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 16 Then
        moveus = 0
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 16 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(16, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(16, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub 麻痺_電腦()
Select Case 異常狀態檢查數(17, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 17 Then
        movecom = 0
        movecheckcom = 0
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 17 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(17, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(17, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub 聖痕_使用者()
Select Case 異常狀態檢查數(13, 1)
  Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 13 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 人物異常狀態資料庫(1, i, 2) * 1
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 人物異常狀態資料庫(1, i, 2) * 1
        Exit For
       End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 13 Then
        攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 人物異常狀態資料庫(1, i, 2) * 1
        攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 人物異常狀態資料庫(1, i, 2) * 1
        Exit For
       End If
     Next
End Select
End Sub
Sub 聖痕_電腦()
Select Case 異常狀態檢查數(26, 1)
  Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 26 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 人物異常狀態資料庫(2, i, 2) * 1
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + 人物異常狀態資料庫(2, i, 2) * 1
        Exit For
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 26 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + 人物異常狀態資料庫(2, i, 2) * 1
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + 人物異常狀態資料庫(2, i, 2) * 1
        Exit For
       End If
     Next
     FormMainMode.trgoi2.Enabled = True
   Case 3
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 26 Then
        攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - 人物異常狀態資料庫(2, i, 2) * 1
        攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - 人物異常狀態資料庫(2, i, 2) * 1
        Exit For
       End If
     Next
End Select
End Sub
Sub 恐怖_使用者()
Select Case 異常狀態檢查數(29, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 29 Then
            擲骰後骰傷害數 = 擲骰後骰傷害數 \ 2
            擲骰表單溝通暫時變數(2) = 擲骰後骰傷害數
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 29 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(29, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(29, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub 狂戰士_使用者()
Select Case 異常狀態檢查數(27, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 27 Then
            擲骰後骰傷害數 = 擲骰後骰傷害數 * 2
            擲骰表單溝通暫時變數(2) = 擲骰後骰傷害數
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 27 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(27, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
            異常狀態檢查數(27, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub 狂戰士_電腦()
Select Case 異常狀態檢查數(28, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 28 Then
            擲骰後骰傷害數 = 擲骰後骰傷害數 * 2
            擲骰表單溝通暫時變數(2) = 擲骰後骰傷害數
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 28 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(28, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(28, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub 恐怖_電腦()
Select Case 異常狀態檢查數(30, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 30 Then
            擲骰後骰傷害數 = 擲骰後骰傷害數 \ 2
            擲骰表單溝通暫時變數(2) = 擲骰後骰傷害數
       End If
     Next
   Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 30 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(30, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
            異常狀態檢查數(30, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub 混沌_使用者()
Select Case 異常狀態檢查數(31, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 31 Then
            If Val(異常狀態_混沌紀錄數(3)) = 0 Then
                異常狀態_混沌紀錄數(1) = 攻擊防禦骰子總數(1)
                異常狀態_混沌紀錄數(2) = 攻擊防禦骰子總數(1) * 2
                異常狀態_混沌紀錄數(3) = 1
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) * 2
            ElseIf Val(異常狀態_混沌紀錄數(3)) = 1 Then
                異常狀態_混沌紀錄數(1) = 異常狀態_混沌紀錄數(1) + (攻擊防禦骰子總數(1) - 異常狀態_混沌紀錄數(2))
                攻擊防禦骰子總數(1) = 異常狀態_混沌紀錄數(1) * 2
                異常狀態_混沌紀錄數(2) = 異常狀態_混沌紀錄數(1) * 2
            End If
       End If
     Next
   Case 2
        For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
          If 人物異常狀態資料庫(1, i, 3) = 31 Then
            人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
            If 人物異常狀態資料庫(1, i, 2) = 0 Then
              '===繼承下一狀態資料
               戰鬥系統類.異常狀態繼承_使用者
               異常狀態檢查數(31, 2) = 0
           Else
               FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
               異常狀態檢查數(31, 1) = 1
           End If
         End If
        Next
   Case 3
        Erase 異常狀態_混沌紀錄數
End Select
End Sub
Sub 混沌_電腦()
Select Case 異常狀態檢查數(32, 1)
   Case 1
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 32 Then
            If Val(異常狀態_AI_混沌紀錄數(3)) = 0 Then
                異常狀態_AI_混沌紀錄數(1) = 攻擊防禦骰子總數(2)
                異常狀態_AI_混沌紀錄數(2) = 攻擊防禦骰子總數(2) * 2
                異常狀態_AI_混沌紀錄數(3) = 1
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) * 2
            ElseIf Val(異常狀態_AI_混沌紀錄數(3)) = 1 Then
                異常狀態_AI_混沌紀錄數(1) = 異常狀態_AI_混沌紀錄數(1) + (攻擊防禦骰子總數(2) - 異常狀態_AI_混沌紀錄數(2))
                攻擊防禦骰子總數(2) = 異常狀態_AI_混沌紀錄數(1) * 2
                異常狀態_AI_混沌紀錄數(2) = 異常狀態_AI_混沌紀錄數(1) * 2
            End If
       End If
     Next
   Case 2
        For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
          If 人物異常狀態資料庫(2, i, 3) = 32 Then
            人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
            If 人物異常狀態資料庫(2, i, 2) = 0 Then
              '===繼承下一狀態資料
               戰鬥系統類.異常狀態繼承_電腦
               異常狀態檢查數(32, 2) = 0
           Else
               FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
               異常狀態檢查數(32, 1) = 1
           End If
         End If
        Next
   Case 3
        Erase 異常狀態_AI_混沌紀錄數
End Select
End Sub

Sub 咒縛_使用者(ByVal moveend As Integer)
Dim dge As Integer
Select Case 異常狀態檢查數(33, 1)
    Case 1
        If movecp > 0 Then
            For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
              If 人物異常狀態資料庫(1, i, 3) = 33 Then
                 dge = Abs(moveend - movecp)
                 傷害執行_技能直傷_使用者 dge, 1
              End If
            Next
        End If
    Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
       If 人物異常狀態資料庫(1, i, 3) = 33 Then
         人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_使用者
            異常狀態檢查數(33, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 咒縛_電腦(ByVal moveend As Integer)
Dim dge As Integer
Select Case 異常狀態檢查數(34, 1)
    Case 1
        If movecp > 0 Then
            For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
              If 人物異常狀態資料庫(2, i, 3) = 34 Then
                 dge = Abs(moveend - movecp)
                 傷害執行_技能直傷_電腦 dge, 1
              End If
            Next
        End If
    Case 2
     For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
       If 人物異常狀態資料庫(2, i, 3) = 34 Then
         人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
         If 人物異常狀態資料庫(2, i, 2) = 0 Then
           '===繼承下一狀態資料
            戰鬥系統類.異常狀態繼承_電腦
            異常狀態檢查數(34, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub 庇護_使用者(ByVal num As Integer, ByRef tot As Integer)
Select Case 異常狀態檢查數(35, 1)
  Case 1
     For i = 14 * (角色待機人物紀錄數(1, num) - 1) + 1 To 14 * 角色待機人物紀錄數(1, num)
       If 人物異常狀態資料庫(1, i, 3) = 35 Then
            If tot > 0 Then
                tot = 0
                '================
                If num = 1 Then
                    FormMainMode.messageus.AddItem "庇護效果發動!    當次受到的傷害無效化"
                Else
                    FormMainMode.messageus.AddItem "庇護效果發動!    待機成員當次受到的傷害無效化"
                End If
                戰鬥系統類.自動捲軸捲動
                '================
                人物異常狀態資料庫(1, i, 2) = 0
                 If 人物異常狀態資料庫(1, i, 2) = 0 Then
                   '===繼承下一狀態資料
                    戰鬥系統類.異常狀態繼承_使用者
                    異常狀態檢查數(35, 2) = 0
                End If
            End If
            Exit For
       End If
     Next
End Select
End Sub
Sub 庇護_電腦(ByVal num As Integer, ByRef tot As Integer)
Select Case 異常狀態檢查數(36, 1)
  Case 1
     For i = 14 * (角色待機人物紀錄數(2, num) - 1) + 1 To 14 * 角色待機人物紀錄數(2, num)
       If 人物異常狀態資料庫(2, i, 3) = 36 Then
            If tot > 0 Then
                tot = 0
                '================
                If num = 1 Then
                    FormMainMode.messageus.AddItem "庇護效果發動!    當次對手受到的傷害無效化"
                Else
                    FormMainMode.messageus.AddItem "庇護效果發動!    當次對手待機成員受到的傷害無效化"
                End If
                戰鬥系統類.自動捲軸捲動
                '================
                人物異常狀態資料庫(2, i, 2) = 0
                 If 人物異常狀態資料庫(2, i, 2) = 0 Then
                   '===繼承下一狀態資料
                    戰鬥系統類.異常狀態繼承_電腦
                    異常狀態檢查數(36, 2) = 0
                End If
            End If
            Exit For
       End If
     Next
End Select
End Sub
Sub 再生_使用者()
Select Case 異常狀態檢查數(37, 1)
    Case 1
        For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
          If 人物異常狀態資料庫(1, i, 3) = 37 Then
            人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i, 2) - 1
            戰鬥系統類.回復執行_使用者 1, 1
            If 人物異常狀態資料庫(1, i, 2) = 0 Then
              '===繼承下一狀態資料
               戰鬥系統類.異常狀態繼承_使用者
               異常狀態檢查數(37, 2) = 0
           Else
               FormMainMode.personusspe(i).person_turn = 人物異常狀態資料庫(1, i, 2)
           End If
         End If
        Next
End Select
End Sub
Sub 再生_電腦()
Select Case 異常狀態檢查數(38, 1)
    Case 1
        For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
          If 人物異常狀態資料庫(2, i, 3) = 38 Then
            人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i, 2) - 1
            戰鬥系統類.回復執行_電腦 1, 1
            If 人物異常狀態資料庫(2, i, 2) = 0 Then
              '===繼承下一狀態資料
               戰鬥系統類.異常狀態繼承_電腦
               異常狀態檢查數(38, 2) = 0
           Else
               FormMainMode.personcomspe(i).person_turn = 人物異常狀態資料庫(2, i, 2)
           End If
         End If
        Next
End Select
End Sub
Sub 臨界_使用者()
Select Case 異常狀態檢查數(39, 1)
  Case 1
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
        If 人物異常狀態資料庫(1, i, 3) = 39 Then
             If 人物異常狀態資料庫(1, i, 2) < 3 Then
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 人物異常狀態資料庫(1, i, 2) * 1
                攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 人物異常狀態資料庫(1, i, 2) * 1
                Exit For
             ElseIf 人物異常狀態資料庫(1, i, 2) >= 3 Then
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + 5
                攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + 5
                Exit For
             End If
        End If
     Next
     FormMainMode.trgoi1.Enabled = True
   Case 2
     For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
        If 人物異常狀態資料庫(1, i, 3) = 39 Then
             If 人物異常狀態資料庫(1, i, 2) < 3 Then
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 人物異常狀態資料庫(1, i, 2) * 1
                攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 人物異常狀態資料庫(1, i, 2) * 1
                Exit For
             ElseIf 人物異常狀態資料庫(1, i, 2) >= 3 Then
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) - 5
                攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) - 5
                Exit For
             End If
        End If
     Next
End Select
End Sub
