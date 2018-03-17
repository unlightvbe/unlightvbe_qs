Attribute VB_Name = "執行階段系統類"
Public VBEPersonVS(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11)  As Variant  'VBE人物統一變數-VS版
Public atkingpagetotVS(1 To 2, 1 To 5) As Variant  '每階段出牌種類及數值統計資料-VS版
Public VBEPersonBuffVSF(1 To 2, 1 To 3, 1 To 14, 1 To 2) As Variant '異常狀態資料-VS-F版
Public VBEPersonBuffVSS(1 To 2, 1 To 3, 1 To 14) As Variant '異常狀態資料-VS-S版
Public AtkingckVSS(1 To 8, 1 To 3) As Variant  '技能資訊一覽-S版(技能啟動碼)
Public AtkingckVSF(1 To 8, 1 To 1) As Variant '技能資訊一覽-F版(備註字串)
Public VBEAtkingVSF(1 To 2, 1 To 3, 1 To 2) As Variant 'VBE>VS給予變數統一資料-F版
Public VBEAtkingVSS(1 To 15) As Variant 'VBE>VS給予變數統一資料-S版
'Public VBEPageCardNumVS(1 To 106, 1 To 6) As Variant '公用牌資料-VS版
Public VBEPageCardNumVS() As Variant '公用牌資料-VS版
Public VBEVSBuffNum(1 To 2) As Variant '異常狀態專用-異常狀態之2個數值-VS版
Public VBEStageNum() As Integer '執行階段系統-執行階段多用途紀錄暫時變數(0.執行階段號/1~任意內容)
Public VBEVSStageNum() As Variant '執行階段系統-執行階段多用途紀錄變數-VS版
'Public VBEVSBuffMainCommadNum(1 To 2) As Integer '執行階段系統-異常狀態類相關暫時紀錄變數(1.是否為場上人物紀錄數/2.角色待機編號紀錄數)
Public VBEStageRemoveBuffAllNum() As Boolean '執行階段系統-執行階段73-異常狀態控制全部清除-異常狀態是否異議標記暫時變數
Public VBEActualStatusVS(1 To 2, 1 To 3, 1 To 2) As Variant '人物實際狀態資料-VS版
Sub 執行階段系統總主要程序_人物主動技能(ByVal uscom As Integer, ByVal atkingnum As Integer, ByVal ns As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    If 執行階段系統類.執行階段系統_驗證(atkingnum, ns, 執行階段系統資料(uscom, atkingnum, 1), uscom, 角色人物對戰人數(uscom, 2)) = True Then
           執行階段系統類.執行階段系統_準備變數統合資料 uscom, VBEStageNumMain
           vbecommadnum(6, vbecommadtotplayNow) = 1
           vbecommadnum(7, vbecommadtotplayNow) = 角色人物對戰人數(uscom, 2)
           執行指令集.執行指令集總程序執行 執行階段系統_執行腳本_人物主動技能類(atkingnum, ns, uscom), 執行階段系統資料(uscom, atkingnum, 3), uscom, atkingnum, ns, vbecommadtotplayNow
    End If
End Sub
Sub 執行階段系統總主要程序_人物被動技能(ByVal uscom As Integer, ByVal personnum As Integer, ByVal atkingnum As Integer, ByVal ns As Integer, ByVal PassivePersonType As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim buffvssnum As Integer '暫時變數
    If 執行階段系統類.執行階段系統_驗證(atkingnum, ns, VBEPerson(uscom, personnum, 3, atkingnum, 3), uscom, personnum) = True Then
           執行階段系統類.執行階段系統_準備變數統合資料 uscom, VBEStageNumMain
           vbecommadnum(6, vbecommadtotplayNow) = PassivePersonType
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           passivevssnum = (uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24
           執行指令集.執行指令集總程序執行 執行階段系統_執行腳本_人物被動技能類(atkingnum, ns, uscom, personnum, PassivePersonType), passivevssnum, uscom, atkingnum, ns, vbecommadtotplayNow
    End If
End Sub
Sub 執行階段73_指令_異常狀態控制_全部清除(ByVal uscom As Integer, ByVal num As Integer)
Dim vbecommadnumSecond As Integer '本層執行階段編號數
'=======================
vbecommadnumSecond = 執行階段系統_宣告開始或結束(1)
'=======================
Dim VBEStageNumMainSec(1 To 1) As Integer
ReDim VBEStageRemoveBuffAllNum(1 To 14) As Boolean
'=======================
For p = 1 To 14
    Vss_EventRemoveBuffActionOffNum = 0
    If num = 1 Then
         '===場上人物之異常狀態
         If Val(人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), p, 2)) > 0 And 人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), p, 3) <> "" Then
             執行階段系統總主要程序_異常狀態 uscom, 角色待機人物紀錄數(uscom, num), p, 73, 1, VBEStageNumMainSec, vbecommadnumSecond
         End If
     Else
         '===待機人物之異常狀態
         If Val(人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), p, 2)) > 0 And 人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), p, 3) <> "" Then
             執行階段系統總主要程序_異常狀態 uscom, 角色待機人物紀錄數(uscom, num), p, 73, 2, VBEStageNumMainSec, vbecommadnumSecond
         End If
     End If
     If Vss_EventRemoveBuffActionOffNum = 1 Then
         VBEStageRemoveBuffAllNum(i) = True
     End If
Next
'=======================
執行階段系統_宣告開始或結束 2
'=======================
End Sub
Sub 執行階段73_指令_異常狀態控制_特定清除(ByVal uscom As Integer, ByVal num As Integer, ByVal buffnum As Integer)
Dim vbecommadnumSecond As Integer '本層執行階段編號數
'=======================
vbecommadnumSecond = 執行階段系統_宣告開始或結束(1)
'=======================
Dim VBEStageNumMainSec(1 To 1) As Integer
'=======================
Vss_EventRemoveBuffActionOffNum = 0
If num = 1 Then
    '===場上人物之異常狀態
    If Val(人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), buffnum, 2)) > 0 And 人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), buffnum, 3) <> "" Then
        執行階段系統總主要程序_異常狀態 uscom, 角色待機人物紀錄數(uscom, num), buffnum, 73, 1, VBEStageNumMainSec, vbecommadnumSecond
    End If
Else
    '===待機人物之異常狀態
    If Val(人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), buffnum, 2)) > 0 And 人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, num), buffnum, 3) <> "" Then
        執行階段系統總主要程序_異常狀態 uscom, 角色待機人物紀錄數(uscom, num), buffnum, 73, 2, VBEStageNumMainSec, vbecommadnumSecond
    End If
End If
'=======================
執行階段系統_宣告開始或結束 2
'=======================
End Sub

Sub 執行階段系統總主要程序_執行階段開始(ByVal uscomFirst As Integer, ByVal ns As Integer, ByVal nstype As Integer)
    Dim vbecommadtotplayNow As Integer '本層執行階段編號數
    '==nstype(1.全執行(驗證)/2.只執行一次(驗證)/3.全執行(不驗證)/4.只執行一次(不驗證)
    '=======================
    vbecommadtotplayNow = 執行階段系統_宣告開始或結束(1)
    '=======================
    Dim VBEStageNumMain() As Integer
    If UBound(VBEStageNum) = 0 Then
        ReDim VBEStageNumMain(1 To 1) As Integer
    Else
        ReDim VBEStageNumMain(0 To UBound(VBEStageNum)) As Integer
        For i = 0 To UBound(VBEStageNum)
           VBEStageNumMain(i) = VBEStageNum(i)
        Next
    End If
'    Dim VBEStageNumMain(1 To UBound(VBEStageNum)) As Integer
'    For i = 1 To 5
'       VBEStageNumMain(i) = VBEStageNum(i)
'    Next
'    Erase VBEStageNum
    '=======================
    Dim uscom As Integer
    For k = 1 To 2
        If k = 1 Then
            If uscomFirst = 1 Then uscom = 1 Else uscom = 2
        Else
            If uscomFirst = 1 Then uscom = 2 Else uscom = 1
        End If
        '==================
        
        '==================人物實際狀態
        For w = 1 To 3
            If w = 1 Then
                 '===場上人物之人物實際狀態
                 If 人物實際狀態資料庫(uscom, 角色待機人物紀錄數(uscom, w), 1) <> "" Then
                     執行階段系統總主要程序_人物實際狀態 uscom, 角色待機人物紀錄數(uscom, w), ns, 1, VBEStageNumMain, vbecommadtotplayNow
                 End If
             Else
                 '===待機人物之人物實際狀態
                 If 人物實際狀態資料庫(uscom, 角色待機人物紀錄數(uscom, w), 1) <> "" Then
                     執行階段系統總主要程序_人物實際狀態 uscom, 角色待機人物紀錄數(uscom, w), ns, 2, VBEStageNumMain, vbecommadtotplayNow
                 End If
             End If
        Next
        '==================異常狀態
        For w = 1 To 3
           For p = 1 To 14
               If w = 1 Then
                    '===場上人物之異常狀態
                    If Val(人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, w), p, 2)) > 0 And 人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, w), p, 3) <> "" Then
                        執行階段系統總主要程序_異常狀態 uscom, 角色待機人物紀錄數(uscom, w), p, ns, 1, VBEStageNumMain, vbecommadtotplayNow
                    End If
                Else
                    '===待機人物之異常狀態
                    If Val(人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, w), p, 2)) > 0 And 人物異常狀態資料庫(uscom, 角色待機人物紀錄數(uscom, w), p, 3) <> "" Then
                        執行階段系統總主要程序_異常狀態 uscom, 角色待機人物紀錄數(uscom, w), p, ns, 2, VBEStageNumMain, vbecommadtotplayNow
                    End If
                End If
           Next
        Next
        戰鬥系統類.異常狀態繼承_使用者
        戰鬥系統類.異常狀態繼承_電腦
        '==================被動技能
        For w = 1 To 3
            If w = 1 Then
                 '===場上人物之人物被動技能
                 For atkingnum = 5 To 8
                    If atkingck(uscom, 角色待機人物紀錄數(uscom, w), atkingnum, 1) = 1 Or Vss_PersonAtkingOffNum(uscom, 角色待機人物紀錄數(uscom, w), atkingnum) = 0 Then
                        執行階段系統總主要程序_人物被動技能 uscom, 角色待機人物紀錄數(uscom, w), atkingnum, ns, 1, VBEStageNumMain, vbecommadtotplayNow
                    End If
                Next
             Else
                 '===待機人物之人物被動技能
                 For atkingnum = 5 To 8
                    If atkingck(uscom, 角色待機人物紀錄數(uscom, w), atkingnum, 1) = 1 Or Vss_PersonAtkingOffNum(uscom, 角色待機人物紀錄數(uscom, w), atkingnum) = 0 Then
                        執行階段系統總主要程序_人物被動技能 uscom, 角色待機人物紀錄數(uscom, w), atkingnum, ns, 2, VBEStageNumMain, vbecommadtotplayNow
                    End If
                Next
             End If
        Next
        '==================主動技能
        For atkingnum = 1 To 4
            If (nstype <= 2 And atkingck(uscom, 角色人物對戰人數(uscom, 2), atkingnum, 1) = 1) Or _
                (nstype > 2 And Vss_PersonAtkingOffNum(uscom, 角色人物對戰人數(uscom, 2), atkingnum) = 0) Then
                執行階段系統總主要程序_人物主動技能 uscom, atkingnum, ns, VBEStageNumMain, vbecommadtotplayNow
            End If
        Next
        '=====================
        If nstype = 2 Or nstype = 4 Then Exit For
    Next
    '=================
    ReDim VBEStageNum(0) As Integer
    執行階段系統_宣告開始或結束 2
    '=================
End Sub
Function 執行階段系統_宣告開始或結束(ByVal n As Integer) As Integer
    Select Case n
        Case 1 '==開始
            vbecommadtotplay = vbecommadtotplay + 1
            ReDim Preserve vbecommadnum(1 To 7, vbecommadtotplay)
            ReDim Preserve vbecommadstr(1 To 3, vbecommadtotplay)
        Case 2 '==結束
            vbecommadtotplay = vbecommadtotplay - 1
            ReDim Preserve vbecommadnum(1 To 7, vbecommadtotplay)
            ReDim Preserve vbecommadstr(1 To 3, vbecommadtotplay)
    End Select
    執行階段系統_宣告開始或結束 = vbecommadtotplay
End Function
Function 執行階段系統_驗證(ByVal atkingnum As Integer, ByVal ns As Integer, ByVal akstr As String, ByVal uscom As Integer, ByVal personnum As Integer) As Boolean
    If Formsetting.checktest.Value = 0 Then On Error GoTo vscheckerr
    Dim vsstr1 As String, vsstr2 As String, vsstr3() As String
    Dim textlinea As String, str As String
    '==========================
    If (uscom = 1 And liveus(personnum) <= 0 And 角色人物對戰人數(uscom, 2) <> personnum) Or _
       (uscom = 2 And livecom(personnum) <= 0 And 角色人物對戰人數(uscom, 2) <> personnum) Then
       執行階段系統_驗證 = False
       Exit Function
    End If
    '==========================
    Select Case atkingnum
        Case Is <= 4  '==主動技能
            If VBEVSSAtkingStr(uscom, 角色人物對戰人數(uscom, 2), atkingnum, 1) <> "" Then
                vsstr1 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * 角色人物對戰人數(uscom, 2) - 4) + atkingnum).Run("main", 1)
                vsstr2 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * 角色人物對戰人數(uscom, 2) - 4) + atkingnum).Run("main", 2)
                '==================
                vsstr3 = Split(vsstr2, "#")
                For k = 0 To UBound(vsstr3)
                    If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                        執行階段系統_驗證 = True
                        Exit Function
                    End If
                Next
                執行階段系統_驗證 = False
            Else
                執行階段系統_驗證 = False
            End If
        Case Is <= 8  '==被動技能
            If VBEVSSAtkingStr(uscom, personnum, atkingnum, 1) <> "" Then
                vsstr1 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("main", 1)
                vsstr2 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("main", 2)
                '==================
                vsstr3 = Split(vsstr2, "#")
                For k = 0 To UBound(vsstr3)
                    If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                        執行階段系統_驗證 = True
                        Exit Function
                    End If
                Next
                執行階段系統_驗證 = False
            Else
                執行階段系統_驗證 = False
            End If
        Case 9 '==異常狀態
             For p = 1 To UBound(VBEVSSBuffStr1)
                If VBEVSSBuffStr1(p) = akstr Then
                    vsstr1 = FormMainMode.PEAFvssc(p + 54).Run("main", 1)
                    vsstr2 = FormMainMode.PEAFvssc(p + 54).Run("main", 2)
                    '==================
                    vsstr3 = Split(vsstr2, "#")
                    For k = 0 To UBound(vsstr3)
                        If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                            執行階段系統_驗證 = True
                            Exit Function
                        End If
                    Next
                End If
             Next
             執行階段系統_驗證 = False
        Case 10 '==人物實際狀態
             For p = 1 To UBound(VBEVSSActualStatusStr1)
                If VBEVSSActualStatusStr1(p) = akstr Then
                    vsstr1 = FormMainMode.PEAFvssc((uscom - 1) * 3 + personnum + 48).Run("main", 1)
                    vsstr2 = FormMainMode.PEAFvssc((uscom - 1) * 3 + personnum + 48).Run("main", 2)
                    '==================
                    vsstr3 = Split(vsstr2, "#")
                    For k = 0 To UBound(vsstr3)
                        If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                            執行階段系統_驗證 = True
                            Exit Function
                        End If
                    Next
                End If
             Next
             執行階段系統_驗證 = False
    End Select
Exit Function
    '==============================
vscheckerr:
    執行階段系統_錯誤訊息通知 2, "1[" & uscom & "-" & ns & "-" & akstr & "]"
End Function
Function 執行階段系統_檔案讀入(ByVal atkingnum As Integer, ByVal name As String, ByVal uscom As Integer) As Boolean
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsloaderror
   Select Case atkingnum
        Case Is <= 4
            Dim textlinea As String, str As String
'            Open app_path & "character\" & name & "\" & VBEVSSAtkingStr(uscom, atkingnum, 2) For Input As #1 '正式用
'            Open App.Path & "\test\input1.txt" For Input As #1
            
            Do Until EOF(1)
               Line Input #1, textlinea
               str = str & textlinea & vbCrLf
            Loop
            Close
            If str <> "" Then
                FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * 角色人物對戰人數(uscom, 2) - 4) + atkingnum).AddCode str
                執行階段系統_檔案讀入 = True
            Else
                執行階段系統_檔案讀入 = False
            End If
        Case Else
    
    End Select
'=====================================
Exit Function
vsloaderror:
執行階段系統_檔案讀入 = False
'=====================================
End Function
Function 執行階段系統資料(ByVal uscom As Integer, ByVal atkingnum As Integer, ByVal numr As Integer) As String
     Select Case numr
          Case 1 '==技能唯一識別碼
              執行階段系統資料 = VBEPerson(uscom, 角色人物對戰人數(uscom, 2), 3, atkingnum, 11)
          Case 2 '==當前讀取階段人名
              執行階段系統資料 = VBEPerson(uscom, 角色人物對戰人數(uscom, 2), 1, 1, 1)
          Case 3 '==當前執行腳本物件號
               執行階段系統資料 = (uscom - 1) * 12 + (4 * 角色人物對戰人數(uscom, 2) - 4) + atkingnum
     End Select
End Function
Function 執行階段系統_執行腳本_人物主動技能類(ByVal atkingnum As Integer, ByVal ns As Integer, ByVal uscom As Integer) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
   執行階段系統_執行腳本_人物主動技能類 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * 角色人物對戰人數(uscom, 2) - 4) + atkingnum).Run("atking", ns, VBEPersonVS, VBEPageCardNumVS, atkingpagetotVS, VBEPersonBuffVSF, VBEPersonBuffVSS, AtkingckVSS, AtkingckVSF, VBEAtkingVSF, VBEAtkingVSS, VBEActualStatusVS, VBEVSStageNum)
'=====================================
Exit Function
'===========
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
執行階段系統類.執行階段系統_初始_腳本讀入程序
GoTo VssAdminReTry
'===========
vsgoerror:
執行階段系統_錯誤訊息通知 2, "2[1-" & atkingnum & "]"
'=====================================

End Function
Function 執行階段系統_執行腳本_人物被動技能類(ByVal atkingnum As Integer, ByVal ns As Integer, ByVal uscom As Integer, ByVal personnum As Integer, ByVal PassivePersonType As Integer) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
   Dim PassivePersonTypeVSS As Variant
   PassivePersonTypeVSS = PassivePersonType
   執行階段系統_執行腳本_人物被動技能類 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("passive", ns, VBEPersonVS, VBEPageCardNumVS, atkingpagetotVS, VBEPersonBuffVSF, VBEPersonBuffVSS, AtkingckVSS, AtkingckVSF, VBEAtkingVSF, VBEAtkingVSS, VBEActualStatusVS, PassivePersonTypeVSS, VBEVSStageNum)
'=====================================
Exit Function
'===========
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
執行階段系統類.執行階段系統_初始_腳本讀入程序
GoTo VssAdminReTry
'===========
vsgoerror:
執行階段系統_錯誤訊息通知 2, "2[2-" & atkingnum & "]"
'=====================================

End Function

Function 執行階段系統_執行腳本_異常狀態類(ByVal vssnum As Integer, ByVal ns As Integer, ByVal BuffPersonType As Integer, ByVal akstr As String) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
    Dim BuffPersonTypeVSS As Variant
    BuffPersonTypeVSS = BuffPersonType
    執行階段系統_執行腳本_異常狀態類 = FormMainMode.PEAFvssc(vssnum).Run("buff", ns, atkingpagetotVS, VBEAtkingVSF, VBEAtkingVSS, VBEVSBuffNum, BuffPersonTypeVSS, VBEVSStageNum)
'=====================================
Exit Function
'===========
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
執行階段系統類.執行階段系統_初始_腳本讀入程序
GoTo VssAdminReTry
'===========
vsgoerror:
執行階段系統_錯誤訊息通知 2, "2[3-" & akstr & "]"
'=====================================

End Function
Function 執行階段系統_執行腳本_人物實際狀態類(ByVal vssnum As Integer, ByVal ns As Integer, ByVal ActualStatusPersonType As Integer, ByVal akstr As String) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
    Dim ActualStatusPersonTypeVSS As Variant
    ActualStatusPersonTypeVSS = ActualStatusPersonType
    執行階段系統_執行腳本_人物實際狀態類 = FormMainMode.PEAFvssc(vssnum).Run("ActualStatus", ns, VBEPersonVS, VBEPageCardNumVS, atkingpagetotVS, VBEPersonBuffVSF, VBEPersonBuffVSS, VBEAtkingVSF, VBEAtkingVSS, ActualStatusPersonTypeVSS, VBEVSStageNum)
'=====================================
Exit Function
'===========
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
執行階段系統類.執行階段系統_初始_腳本讀入程序
GoTo VssAdminReTry
'===========
vsgoerror:
執行階段系統_錯誤訊息通知 2, "2[4-" & akstr & "]"
'=====================================

End Function
Sub 執行階段系統_準備變數統合資料(ByVal uscom As Integer, ByRef VBEStageNumMain() As Integer)
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
    ReDim VBEPageCardNumVS(1 To 公用牌實體卡片分隔紀錄數(1), 1 To 6) As Variant '公用牌資料-VS版
'    Erase VBEVSStageNum '執行階段系統-執行階段多用途紀錄變數-VS版
    ReDim VBEVSStageNum(1 To UBound(VBEStageNumMain)) As Variant '執行階段系統-執行階段多用途紀錄變數-VS版
    Erase VBEActualStatusVS '人物實際狀態資料-VS版
    '===========================
    Dim q As Integer, w As Integer, rr As Integer
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
            For i = 1 To 公用牌實體卡片分隔紀錄數(1)
                For j = 1 To 6
                    If j = 1 Or j = 3 Then
                       Select Case pagecardnum(i, j)
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
                    Else
                       VBEPageCardNumVS(i, j) = Val(pagecardnum(i, j))
                    End If
                Next
            Next
            '======================
            '(1 To 2, 1 To 5)
            For i = 1 To 2
                For j = 1 To 5
                    atkingpagetotVS(i, j) = atkingpagetot(i, j)
                Next
            Next
            '======================
            '(1 To 2, 1 To 3, 1 To 14, 1 To 3)
            For i = 1 To 2
                For rr = 1 To 3
                    For j = 1 To 14
                        For k = 1 To 3
                            If k = 3 Then
                                VBEPersonBuffVSS(i, rr, j) = 人物異常狀態資料庫(i, 角色待機人物紀錄數(i, rr), j, k)
                            Else
                                VBEPersonBuffVSF(i, rr, j, k) = Val(人物異常狀態資料庫(i, 角色待機人物紀錄數(i, rr), j, k))
                            End If
                        Next
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
            VBEAtkingVSS(1) = pageqlead(1)
            VBEAtkingVSS(2) = pageglead(1)
            VBEAtkingVSS(3) = pageqlead(2)
            VBEAtkingVSS(4) = pageglead(2)
            VBEAtkingVSS(5) = 擲骰表單溝通暫時變數(2)
            VBEAtkingVSS(6) = movecp
            VBEAtkingVSS(7) = Val(攻擊防禦骰子總數(1))
            VBEAtkingVSS(8) = Val(攻擊防禦骰子總數(2))
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            VBEAtkingVSS(14) = 擲骰表單溝通暫時變數(5)
            VBEAtkingVSS(15) = 擲骰表單溝通暫時變數(6)
            Select Case turnatk
                Case 1
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 2
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                Case 4, 6
                    VBEAtkingVSS(12) = 1
                    VBEAtkingVSS(13) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
             End Select
             '=========================
             If LBound(VBEStageNumMain) = 0 Then
                    Select Case VBEStageNumMain(0)
                        Case 71  '執行階段71(普通-移動前)
                            For i = 1 To UBound(VBEStageNumMain)
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                            Next
                        Case 41, 46, 48 '執行階段41/46/48(角色交換/傷害/回復)
                            For i = 1 To UBound(VBEStageNumMain)
                                    If VBEStageNumMain(i) = -1 Or VBEStageNumMain(i) = -2 Then
                                        VBEVSStageNum(i) = Abs(VBEStageNumMain(i))
                                    Else
                                        VBEVSStageNum(i) = VBEStageNumMain(i)
                                    End If
                            Next
                        Case Else
                            For i = 1 To UBound(VBEStageNumMain)
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                            Next
                    End Select
             Else
                    VBEVSStageNum(1) = 0 '無資料
             End If
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
            For i = 1 To 公用牌實體卡片分隔紀錄數(1)
                For j = 1 To 6
                     If j = 1 Or j = 3 Then
                       Select Case pagecardnum(i, j)
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
                    ElseIf j = 5 Then
                       If Val(pagecardnum(i, j)) = 2 Then
                           VBEPageCardNumVS(i, j) = 1
                        ElseIf Val(pagecardnum(i, j)) = 1 Then
                           VBEPageCardNumVS(i, j) = 2
                        Else
                           VBEPageCardNumVS(i, j) = 0
                        End If
                    Else
                       VBEPageCardNumVS(i, j) = Val(pagecardnum(i, j))
                    End If
                Next
            Next
            '======================
            '(1 To 2, 1 To 5)
            For i = 1 To 2
                If i = 1 Then q = 2 Else q = 1
                For j = 1 To 5
                    atkingpagetotVS(i, j) = atkingpagetot(q, j)
                Next
            Next
            '======================
            '(1 To 2, 1 To 3, 1 To 14, 1 To 3)
            For i = 1 To 2
                If i = 1 Then q = 2 Else q = 1
                For rr = 1 To 3
                    For j = 1 To 14
                        For k = 1 To 3
                            If k = 3 Then
                                VBEPersonBuffVSS(i, rr, j) = 人物異常狀態資料庫(q, 角色待機人物紀錄數(q, rr), j, k)
                            Else
                                VBEPersonBuffVSF(i, rr, j, k) = Val(人物異常狀態資料庫(q, 角色待機人物紀錄數(q, rr), j, k))
                            End If
                        Next
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
            VBEAtkingVSS(1) = pageqlead(2)
            VBEAtkingVSS(2) = pageglead(2)
            VBEAtkingVSS(3) = pageqlead(1)
            VBEAtkingVSS(4) = pageglead(1)
            VBEAtkingVSS(5) = 擲骰表單溝通暫時變數(2)
            VBEAtkingVSS(6) = movecp
            VBEAtkingVSS(7) = Val(攻擊防禦骰子總數(2))
            VBEAtkingVSS(8) = Val(攻擊防禦骰子總數(1))
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            VBEAtkingVSS(14) = 擲骰表單溝通暫時變數(6)
            VBEAtkingVSS(15) = 擲骰表單溝通暫時變數(5)
            Select Case turnatk
                Case 1
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 2
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                Case 4, 6
                    VBEAtkingVSS(12) = 1
                    VBEAtkingVSS(13) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
             End Select
             '=========================
             If LBound(VBEStageNumMain) = 0 Then
                    Select Case VBEStageNumMain(0)
                        Case 71  '執行階段71(普通-移動前)
                            VBEVSStageNum(1) = VBEStageNumMain(2)
                            VBEVSStageNum(2) = VBEStageNumMain(1)
                        Case 41, 46, 48 '執行階段41/46/48(角色交換/傷害/回復)
                            For i = 1 To UBound(VBEStageNumMain)
                                If VBEStageNumMain(i) = -1 Then
                                    VBEVSStageNum(i) = 2
                                ElseIf VBEStageNumMain(i) = -2 Then
                                    VBEVSStageNum(i) = 1
                                Else
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                                End If
                            Next
                        Case Else
                            For i = 1 To UBound(VBEStageNumMain)
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                            Next
                    End Select
             Else
                    VBEVSStageNum(1) = 0 '無資料
             End If
   End Select
End Sub
Sub 執行階段系統_初始_腳本讀入程序()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
Dim atknum As Integer, uscomn As Integer, pnnum As Integer, buffnum As Integer
atknum = 1: uscomn = 1: pnnum = 1: buffnum = 1
Dim tot As Integer, textlinea As String, str As String
tot = 1
Do
     textlinea = ""
     str = ""
     Select Case tot
         Case Is <= 24
                If VBEVSSAtkingStr(uscomn, pnnum, atknum, 1) <> "" Then
                    Open app_path & "character\" & VBEPerson(uscomn, pnnum, 1, 1, 2) & "\" & VBEVSSAtkingStr(uscomn, pnnum, atknum, 2) For Input As #1
                    
                    Do Until EOF(1)
                       Line Input #1, textlinea
                       str = str & textlinea & vbCrLf
                    Loop
                    Close
                    If str <> "" Then
                        FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum).AddCode str
                    End If
                End If
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
         Case Is <= 48
                If VBEVSSAtkingStr(uscomn, pnnum, atknum + 4, 1) <> "" Then
                    Open app_path & "character\" & VBEPerson(uscomn, pnnum, 1, 1, 2) & "\" & VBEVSSAtkingStr(uscomn, pnnum, atknum + 4, 2) For Input As #1
                    
                    Do Until EOF(1)
                       Line Input #1, textlinea
                       str = str & textlinea & vbCrLf
                    Loop
                    Close
                    If str <> "" Then
                        FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum + 24).AddCode str
                    End If
                End If
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
         Case Is <= 54
                
         Case Else
                Open VBEVSSBuffStr2(buffnum) For Input As #1
                
                Do Until EOF(1)
                   Line Input #1, textlinea
                   str = str & textlinea & vbCrLf
                Loop
                Close
                If str <> "" Then
                    FormMainMode.PEAFvssc(tot).AddCode str
                End If
                buffnum = buffnum + 1
    End Select
    tot = tot + 1
Loop Until tot > (Val(54) + Val(UBound(VBEVSSBuffStr2)))
'===============
Exit Sub
vssyserror:
If tot <= 48 Then
    執行階段系統_錯誤訊息通知 1, "3[" & VBEVSSAtkingStr(uscomn, pnnum, atknum, 1) & "]"
ElseIf tot > 48 And tot <= 54 Then
    執行階段系統_錯誤訊息通知 1, "3[" & VBEVSSActualStatusStr2(buffnum) & "]"
Else
    執行階段系統_錯誤訊息通知 1, "3[" & VBEVSSBuffStr2(buffnum) & "]"
End If
'===============
End Sub
Sub 執行階段系統遊戲初始總程序()
       執行階段系統類.執行階段系統_異常狀態類腳本搜尋
       執行階段系統類.執行階段系統_人物實際狀態類腳本搜尋
       執行階段系統類.執行階段系統_初始_腳本物件創立 (Val(54) + Val(UBound(VBEVSSBuffStr2)))
       執行階段系統類.執行階段系統_初始_腳本讀入程序
       執行階段系統類.執行階段系統_初始_人物主動及被動技能類資料讀入
       執行階段系統類.執行階段系統_初始_異常狀態類資料讀入
       執行階段系統類.執行階段系統_初始_人物實際狀態類資料讀入
End Sub
Sub 執行階段系統_初始_腳本物件創立(ByVal num As Integer)
       If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
        '==========
        For i = 1 To num
           Load FormMainMode.PEAFvssc(i)
        Next
        '==========
        '==========
        For i = 1 To num
           FormMainMode.PEAFvssc(i).Reset
        Next
        '==========
'===============
Exit Sub
vssyserror:
執行階段系統_錯誤訊息通知 1, "2"
'===============
End Sub
Sub 執行階段系統_異常狀態類腳本搜尋()
  If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
  Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
  mypath = App.Path & "\Buff\"
  mydir = Dir(mypath, vbDirectory) ' 找尋第一個子目錄。
  ReDim DirectoryBuff(0)
  ReDim VBEVSSBuffStr1(0)
  ReDim VBEVSSBuffStr2(0)
  Do While True
        Do While mydir <> ""
            ' 跳過目前的目錄及上層目錄。
            If mydir <> "." And mydir <> ".." Then
                ' 使用位元比對來確定 MyName 代表一目錄。
                If (GetAttr(mypath & mydir) And vbDirectory) = vbDirectory Then
                    Debug.Print mydir ' 將目錄名稱顯示出來。
                    ReDim Preserve DirectoryBuff(UBound(DirectoryBuff) + 1)
                    DirectoryBuff(UBound(DirectoryBuff)) = mypath + mydir
                Else
                    If 測試1.GetExtName(mydir) = "ulevsbf" And Index >= 1 Then
                        執行階段系統類.執行階段系統_初始_異常狀態類腳本加入紀錄 mydir, DirectoryBuff(Index) & "\"
                    ElseIf 測試1.GetExtName(mydir) = "ulevsbf" And Index = 0 Then
                        執行階段系統類.執行階段系統_初始_異常狀態類腳本加入紀錄 mydir, App.Path & "\Buff\"
                    End If
                End If
            End If
            mydir = Dir()
        Loop
        Index = Index + 1
        If Index > UBound(DirectoryBuff) Then Exit Do
        mypath = DirectoryBuff(Index) + "\"
        mydir = Dir(mypath, vbDirectory)
  Loop
'===============
Exit Sub
vssyserror:
執行階段系統_錯誤訊息通知 1, "1"
'===============
End Sub
Sub 執行階段系統_人物實際狀態類腳本搜尋()
  If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
  Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
  mypath = App.Path & "\CharacterActualStatus\"
  mydir = Dir(mypath, vbDirectory) ' 找尋第一個子目錄。
  ReDim DirectoryBuff(0)
  ReDim VBEVSSActualStatusStr1(0)
  ReDim VBEVSSActualStatusStr2(0)
  Do While True
        Do While mydir <> ""
            ' 跳過目前的目錄及上層目錄。
            If mydir <> "." And mydir <> ".." Then
                ' 使用位元比對來確定 MyName 代表一目錄。
                If (GetAttr(mypath & mydir) And vbDirectory) = vbDirectory Then
                    Debug.Print mydir ' 將目錄名稱顯示出來。
                    ReDim Preserve DirectoryBuff(UBound(DirectoryBuff) + 1)
                    DirectoryBuff(UBound(DirectoryBuff)) = mypath + mydir
                Else
                    If 測試1.GetExtName(mydir) = "ulevsc" And Index >= 1 Then
                        執行階段系統類.執行階段系統_初始_人物實際狀態類腳本加入紀錄 mydir, DirectoryBuff(Index) & "\"
                    ElseIf 測試1.GetExtName(mydir) = "ulevsc" And Index = 0 Then
                        執行階段系統類.執行階段系統_初始_人物實際狀態類腳本加入紀錄 mydir, App.Path & "\CharacterActualStatus\"
                    End If
                End If
            End If
            mydir = Dir()
        Loop
        Index = Index + 1
        If Index > UBound(DirectoryBuff) Then Exit Do
        mypath = DirectoryBuff(Index) + "\"
        mydir = Dir(mypath, vbDirectory)
  Loop
'===============
Exit Sub
vssyserror:
執行階段系統_錯誤訊息通知 1, "6"
'===============
End Sub
Sub 執行階段系統_初始_異常狀態類腳本加入紀錄(ByVal str1 As String, ByVal personname As String)
    ReDim Preserve VBEVSSBuffStr2(UBound(VBEVSSBuffStr2) + 1)
    VBEVSSBuffStr2(UBound(VBEVSSBuffStr2)) = personname & str1
End Sub
Sub 執行階段系統_初始_人物實際狀態類腳本加入紀錄(ByVal str1 As String, ByVal personname As String)
    ReDim Preserve VBEVSSActualStatusStr2(UBound(VBEVSSActualStatusStr2) + 1)
    VBEVSSActualStatusStr2(UBound(VBEVSSActualStatusStr2)) = personname & str1
End Sub
Sub 執行階段系統_初始_人物主動及被動技能類資料讀入()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
Dim vsstr As String, 文件字串() As String
Dim atknum As Integer, uscomn As Integer, pnnum As Integer
Dim tot As Integer
atknum = 1: uscomn = 1: pnnum = 1
tot = 1
Do
    textlinea = ""
    vsstr = ""
    Select Case tot
         Case Is <= 24
                If VBEVSSAtkingStr(uscomn, pnnum, atknum, 1) <> "" Then
                    For i = 3 To 7
                        vsstr = FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum).Run("main", i)
                        文件字串 = Split(vsstr, "#")
                        '==================
                        Select Case i
                            Case 3
                                VBEPerson(uscomn, pnnum, 3, atknum, 1) = 文件字串(0)
                            Case 4
                                VBEPerson(uscomn, pnnum, 3, atknum, 2) = 文件字串(0)
                                VBEPerson(uscomn, pnnum, 3, atknum, 8) = 文件字串(1)
                            Case 5
                                VBEPerson(uscomn, pnnum, 3, atknum, 3) = 文件字串(0)
                                VBEPerson(uscomn, pnnum, 3, atknum, 9) = 文件字串(1)
                            Case 6
                                VBEPerson(uscomn, pnnum, 3, atknum, 4) = 文件字串(0)
                                VBEPerson(uscomn, pnnum, 3, atknum, 10) = 文件字串(1)
                            Case 7
                                VBEPerson(uscomn, pnnum, 3, atknum, 5) = ""
                                For k = 0 To UBound(文件字串)
                                     VBEPerson(uscomn, pnnum, 3, atknum, 5) = VBEPerson(uscomn, pnnum, 3, atknum, 5) & 文件字串(k)
                                Next
                        End Select
                    Next
                End If
                '=================================================
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
         Case Is <= 48
                If VBEVSSAtkingStr(uscomn, pnnum, atknum + 4, 1) <> "" Then
                    For i = 3 To 4
                        vsstr = FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum + 24).Run("main", i)
                        文件字串 = Split(vsstr, "#")
                        '==================
                        Select Case i
                            Case 3
                                VBEPerson(uscomn, pnnum, 3, atknum + 4, 1) = 文件字串(0)
                            Case 4
                                VBEPerson(uscomn, pnnum, 3, atknum + 4, 2) = ""
                                For k = 0 To UBound(文件字串)
                                     VBEPerson(uscomn, pnnum, 3, atknum + 4, 2) = VBEPerson(uscomn, pnnum, 3, atknum + 4, 2) & 文件字串(k)
                                Next
                        End Select
                    Next
                End If
                '=================================================
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
    End Select
    tot = tot + 1
Loop Until tot > 48
'===============
Exit Sub
vssyserror:
If tot <= 24 Then
    執行階段系統_錯誤訊息通知 1, "4[" & uscomn & "," & atknum & "]"
Else
    執行階段系統_錯誤訊息通知 1, "4[" & uscomn & "," & atknum + 4 & "]"
End If
'===============
End Sub
Sub 執行階段系統_初始_異常狀態類資料讀入()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
ReDim VBEVSSBuffStr1(UBound(VBEVSSBuffStr2))
Dim vsstr As String
For i = 1 To UBound(VBEVSSBuffStr2)
    vsstr = FormMainMode.PEAFvssc(i + 54).Run("main", 1)
    VBEVSSBuffStr1(i) = vsstr
Next
'===============
Exit Sub
vssyserror:
執行階段系統_錯誤訊息通知 1, "5[" & VBEVSSBuffStr2(i) & "]"
'===============
End Sub
Sub 執行階段系統_初始_人物實際狀態類資料讀入()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
ReDim VBEVSSActualStatusStr1(UBound(VBEVSSActualStatusStr2))
Dim vsstr As String, textlinea As String, str As String
For i = 1 To UBound(VBEVSSActualStatusStr2)
    Open VBEVSSActualStatusStr2(i) For Input As #1
    Do Until EOF(1)
       Line Input #1, textlinea
       str = str & textlinea & vbCrLf
    Loop
    Close
    If str <> "" Then
        FormMainMode.PEAFvssc(49).AddCode str
    End If
    vsstr = FormMainMode.PEAFvssc(49).Run("main", 1)
    VBEVSSActualStatusStr1(i) = vsstr
    FormMainMode.PEAFvssc(49).Reset
Next
'===============
Exit Sub
vssyserror:
執行階段系統_錯誤訊息通知 1, "7[" & VBEVSSActualStatusStr2(i) & "]"
'===============
End Sub
Sub 執行階段系統總主要程序_異常狀態(ByVal uscom As Integer, ByVal personnum As Integer, ByVal personbuffnum As Integer, ByVal ns As Integer, ByVal BuffPersonType As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim buffvssnum As Integer '暫時變數
    If 執行階段系統類.執行階段系統_驗證(9, ns, 人物異常狀態資料庫(uscom, personnum, personbuffnum, 3), uscom, personnum) = True Then
           執行階段系統類.執行階段系統_準備變數統合資料 uscom, VBEStageNumMain
           vbecommadnum(6, vbecommadtotplayNow) = BuffPersonType
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           Erase VBEVSBuffNum '異常狀態專用-異常狀態之2個數值-VS版
           For p = 1 To UBound(VBEVSSBuffStr1)
                 If VBEVSSBuffStr1(p) = 人物異常狀態資料庫(uscom, personnum, personbuffnum, 3) Then
                     buffvssnum = p + 54
                     VBEVSBuffNum(1) = Val(人物異常狀態資料庫(uscom, personnum, personbuffnum, 1))
                     VBEVSBuffNum(2) = Val(人物異常狀態資料庫(uscom, personnum, personbuffnum, 2))
                     Exit For
                 End If
            Next
           執行指令集.執行指令集總程序執行 執行階段系統_執行腳本_異常狀態類(buffvssnum, ns, BuffPersonType, VBEVSSBuffStr1(p)), buffvssnum, uscom, 9, ns, vbecommadtotplayNow
    End If
End Sub
Sub 執行階段系統總主要程序_人物實際狀態(ByVal uscom As Integer, ByVal personnum As Integer, ByVal ns As Integer, ByVal ActualStatusPersonType As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim ActualStatusvssnum As Integer '暫時變數
    If 執行階段系統類.執行階段系統_驗證(10, ns, 人物實際狀態資料庫(uscom, personnum, 1), uscom, personnum) = True Then
           執行階段系統類.執行階段系統_準備變數統合資料 uscom, VBEStageNumMain
           vbecommadnum(6, vbecommadtotplayNow) = ActualStatusPersonType
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           ActualStatusvssnum = (((uscom - 1) * 3) + personnum) + 48
           執行指令集.執行指令集總程序執行 執行階段系統_執行腳本_人物實際狀態類(ActualStatusvssnum, ns, ActualStatusPersonType, 人物實際狀態資料庫(uscom, personnum, 1)), ActualStatusvssnum, uscom, 10, ns, vbecommadtotplayNow
    End If
End Sub
Function 執行階段系統_搜尋正在執行之執行階段(ByVal vscmdname As String) As Integer
    For i = 1 To vbecommadtotplay
         If vbecommadstr(1, i) = vscmdname Then
             執行階段系統_搜尋正在執行之執行階段 = i
             Exit Function
         End If
    Next
    '==========如果找不到時
    執行階段系統_搜尋正在執行之執行階段 = 0
End Function
Sub 執行階段系統_錯誤訊息通知(ByVal num As Integer, ByVal num1 As String)
MsgBox "執行階段錯誤(03-" & num & "-" & num1 & ")：" & Chr(10) & "系統於讀取及解釋腳本指令時發生錯誤。" & Chr(10) & Chr(10) & "(" & Err.Number & "):" & Err.Description, vbCritical
End
End Sub
