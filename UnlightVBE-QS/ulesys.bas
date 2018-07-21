Attribute VB_Name = "一般系統類"
Public app_path As String  '路徑設定碼
Public 角色人物對戰人數(1 To 2, 1 To 2) As Integer '雙方對戰角色人數紀錄數(1.使用者/2.電腦,1.總共人數/2.目前第幾位)
Public 角色待機人物紀錄數(1 To 2, 1 To 3) As Integer '雙方待機角色人物編號紀錄數(1.使用者/2.電腦,1.場上角色/2~3.待機角色第n位編號)
Public tr1num As Integer 'PEStartForm計數器暫時變數
Public PEAEtr1num As Integer 'PEAttackingEndingForm計數器暫時變數
Public st As Integer, sq As Integer, swq As Integer, cardusq As Integer, cardcomq As Integer   'PEAttackingStartForm計數器暫時變數
Public 第一次啟動讀入程序標記 As Boolean '第一次啟動程式讀入程序標記數
Public 接續讀入表單串 As String 'PEStartForm接續讀入表單暫時紀錄數
Public 音樂檢查播放目標數 As Integer '音樂檢查播放計數器目標數
Public 通知表單是否已出現 As Boolean '布勞通知表單是否已經出現暫時變數

Sub 判斷字型_FormMainMode()
Dim i, a As Integer
a = 14
If FormMainMode.PEStext1.FontName <> "Bradley Gratis" Then
    '===========PEAttackingForm
    For i = 1 To 3
'       FormMainMode.usbi1(i).FontSize = a
'        FormMainMode.usbi2(i).FontSize = a
'        FormMainMode.usbi3(i).FontSize = a
'        FormMainMode.cardcompi1(i).FontSize = a
'        FormMainMode.cardcompi2(i).FontSize = a
'        FormMainMode.cardcompi3(i).FontSize = a
        FormMainMode.compi4(i).FontSize = a
        FormMainMode.uspi4(i).FontSize = a
    Next
    FormMainMode.pageul.FontSize = 24
    FormMainMode.bloodnumcom1.FontSize = 20
    FormMainMode.bloodnumcom2.FontSize = 10
'    FormMainMode.turni.FontSize = 20
    FormMainMode.bloodnumus1.FontSize = 20
    FormMainMode.bloodnumus2.FontSize = 10
    '===========PEGameFreeModeSettingForm
    For i = 1 To 3
        FormMainMode.PEGFusbi1(i).FontSize = a
        FormMainMode.PEGFusbi2(i).FontSize = a
        FormMainMode.PEGFusbi3(i).FontSize = a
        FormMainMode.PEGFcardcompi1(i).FontSize = a
        FormMainMode.PEGFcardcompi2(i).FontSize = a
        FormMainMode.PEGFcardcompi3(i).FontSize = a
    Next
    '===========PEAttackingStartForm
    For i = 1 To 3
        FormMainMode.PEASusbi1(i).FontSize = a
        FormMainMode.PEASusbi2(i).FontSize = a
        FormMainMode.PEASusbi3(i).FontSize = a
        FormMainMode.PEAScardcompi1(i).FontSize = a
        FormMainMode.PEAScardcompi2(i).FontSize = a
        FormMainMode.PEAScardcompi3(i).FontSize = a
    Next
End If
End Sub
     
Sub 開始遊戲進行程序()
Dim i As Integer, m As Integer, n As Integer, u As Integer, personvsp As Integer, perosntempCount As Integer
Dim personnameg(1 To 2, 1 To 3) As String '人物隨機選擇紀錄數
一般系統類.音效播放 11
選單使用者事件 = False
選單電腦事件 = False
電腦方事件卡是否出完選擇數 = False
是否移動階段續估計判斷程序 = False
'==============清除所有變數值與恢復設定
一般系統類.清除戰鬥系統所有變數值
一般系統類.清除戰鬥系統開始表單設定值
'==============角色選擇(隨機)
For i = 1 To 3
    If FormMainMode.personnameus(i).Text = "《隨機》" Then
        personnameg(1, i) = ""
    Else
        personnameg(1, i) = FormMainMode.personnameus(i).Text
    End If
    If FormMainMode.personnamecom(i).Text = "《隨機》" Then
        personnameg(2, i) = ""
    Else
        personnameg(2, i) = FormMainMode.personnamecom(i).Text
    End If
Next
If FormMainMode.opnpersonvs(1).Value = True Then
    personvsp = 1
    If FormMainMode.personnameus(1).Text = "《隨機》" Then
        Randomize
        perosntempCount = 0
        Do
            m = Int(Rnd() * (FormMainMode.personnameus(1).ListCount - 1)) + 1
            人物系統類.卡片人物資訊讀入_二階段_使用者 FormMainMode.personnameus(1).List(m), 1
            If Formsetting.chkusesimilarlevel.Value = 0 Then
                n = Int(Rnd() * (FormMainMode.personlevelus(1).ListCount - 1 + 1)) + 0
                Exit Do
            End If
            n = 角色人物隨機相似等級選擇(1, 1)
            If n <> -1 Then
                Exit Do
            Else
                perosntempCount = perosntempCount + 1
                If perosntempCount > 3 Then
                    n = Int(Rnd() * (FormMainMode.personlevelus(1).ListCount - 1 + 1)) + 0
                End If
            End If
        Loop Until (perosntempCount > 3)
'        人物系統類.卡片人物資訊讀入_三階段_使用者 FormMainMode.personnameus(1).List(m), FormMainMode.personlevelus(1).List(n), 1, 1
        人物系統類.清除角色人物資訊變數 1, 1
        人物系統類.卡片人物資訊讀入_三階段 FormMainMode.personnameus(1).List(m), FormMainMode.personlevelus(1).List(n), 1, 1
'        人物系統類.卡片人物資訊讀入_四階段_使用者 FormMainMode.personnameus(1).List(m), 1
        人物系統類.卡片人物資訊讀入_四階段 FormMainMode.personnameus(1).List(m), 1, 1
    End If
    If FormMainMode.personnamecom(1).Text = "《隨機》" Then
        Randomize
        perosntempCount = 0
        Do
            m = Int(Rnd() * (FormMainMode.personnamecom(1).ListCount - 1)) + 1
            人物系統類.卡片人物資訊讀入_二階段_電腦 FormMainMode.personnamecom(1).List(m), 1
            If Formsetting.chkusesimilarlevel.Value = 0 Then
                n = Int(Rnd() * (FormMainMode.personlevelcom(1).ListCount - 1 + 1)) + 0
                Exit Do
            End If
            n = 角色人物隨機相似等級選擇(2, 1)
            If n <> -1 Then
                Exit Do
            Else
                perosntempCount = perosntempCount + 1
                If perosntempCount > 3 Then
                    n = Int(Rnd() * (FormMainMode.personlevelcom(1).ListCount - 1 + 1)) + 0
                End If
            End If
        Loop Until (perosntempCount > 3)
'        人物系統類.卡片人物資訊讀入_三階段_電腦 FormMainMode.personnamecom(1).List(m), FormMainMode.personlevelcom(1).List(n), 1, 2
        人物系統類.清除角色人物資訊變數 2, 1
        人物系統類.卡片人物資訊讀入_三階段 FormMainMode.personnamecom(1).List(m), FormMainMode.personlevelcom(1).List(n), 1, 2
'        人物系統類.卡片人物資訊讀入_四階段_電腦 formmainmode.personnamecom(1).List(m), 1
    End If
Else
    personvsp = 3
    For i = 1 To 3
        If FormMainMode.personnameus(i).Text = "《隨機》" Then
            Randomize
            perosntempCount = 0
            Do
                m = Int(Rnd() * (FormMainMode.personnameus(i).ListCount - 1)) + 1
                personnameg(1, i) = FormMainMode.personnameus(i).List(m)
                人物系統類.卡片人物資訊讀入_二階段_使用者 FormMainMode.personnameus(i).List(m), i
                If Formsetting.chkusesimilarlevel.Value = 0 Then
                    n = Int(Rnd() * (FormMainMode.personlevelus(i).ListCount - 1 + 1)) + 0
                    Exit Do
                End If
                n = 角色人物隨機相似等級選擇(1, i)
                If n <> -1 Then
                    Exit Do
                Else
                    perosntempCount = perosntempCount + 1
                    If perosntempCount > 3 Then
                        n = Int(Rnd() * (FormMainMode.personlevelus(i).ListCount - 1 + 1)) + 0
                    End If
                End If
            Loop Until (perosntempCount > 3)
'            人物系統類.卡片人物資訊讀入_三階段_使用者 FormMainMode.personnameus(i).List(m), FormMainMode.personlevelus(i).List(n), i, 1
'            人物系統類.卡片人物資訊讀入_四階段_使用者 FormMainMode.personnameus(i).List(m), i
            人物系統類.清除角色人物資訊變數 1, i
            人物系統類.卡片人物資訊讀入_三階段 FormMainMode.personnameus(i).List(m), FormMainMode.personlevelus(i).List(n), i, 1
            人物系統類.卡片人物資訊讀入_四階段 FormMainMode.personnameus(i).List(m), i, 1
            更新人物清單_使用者方_變更_開始隨機 i, personnameg(1, 1), personnameg(1, 2), personnameg(1, 3)
        End If
        If FormMainMode.personnamecom(i).Text = "《隨機》" Then
            Randomize
            perosntempCount = 0
            Do
                m = Int(Rnd() * (FormMainMode.personnamecom(i).ListCount - 1)) + 1
                personnameg(2, i) = FormMainMode.personnamecom(i).List(m)
                人物系統類.卡片人物資訊讀入_二階段_電腦 FormMainMode.personnamecom(i).List(m), i
                If Formsetting.chkusesimilarlevel.Value = 0 Then
                    n = Int(Rnd() * (FormMainMode.personlevelcom(i).ListCount - 1 + 1)) + 0
                    Exit Do
                End If
                n = 角色人物隨機相似等級選擇(2, i)
                If n <> -1 Then
                    Exit Do
                Else
                    perosntempCount = perosntempCount + 1
                    If perosntempCount > 3 Then
                        n = Int(Rnd() * (FormMainMode.personlevelcom(i).ListCount - 1 + 1)) + 0
                    End If
                End If
            Loop Until (perosntempCount > 3)
'            人物系統類.卡片人物資訊讀入_三階段_電腦 FormMainMode.personnamecom(i).List(m), FormMainMode.personlevelcom(i).List(n), i, 2
            人物系統類.清除角色人物資訊變數 2, i
            人物系統類.卡片人物資訊讀入_三階段 FormMainMode.personnamecom(i).List(m), FormMainMode.personlevelcom(i).List(n), i, 2
'            人物系統類.卡片人物資訊讀入_四階段_電腦 formmainmode.personnamecom(i).List(m), i
            更新人物清單_電腦方_變更_開始隨機 i, personnameg(2, 1), personnameg(2, 2), personnameg(2, 3)
        End If
    Next
End If
'=================大亂鬥選項
If Formsetting.大亂鬥選項.Value = 1 Then
   If personvsp = 1 Then
        VBEPerson(1, 1, 1, 3, 1) = 99
        VBEPerson(2, 1, 1, 3, 1) = 99
   ElseIf personvsp = 3 Then
        For i = 1 To 3
            VBEPerson(1, i, 1, 3, 1) = 99
            VBEPerson(2, i, 1, 3, 1) = 99
        Next
   End If
End If
'=======檢查設定
If 角色人物選擇空值檢查(personvsp) = False Then
   Form9.Left = FormMainMode.Left + 1185
   Form9.Top = FormMainMode.Top + 3030
   Form9.Show 1
Else
'-------------
    Select Case personvsp
       Case 1
            For i = 1 To 1
                FormMainMode.cardusname(i).Caption = VBEPerson(1, 1, 1, 1, 1)
                FormMainMode.cardusspname(i).Caption = VBEPerson(1, 1, 1, 1, 3)
                FormMainMode.cardcomname(i).Caption = VBEPerson(2, 1, 1, 1, 1)
                FormMainMode.cardcomspname(i).Caption = VBEPerson(2, 1, 1, 1, 3)
                FormMainMode.PEAScardus(i).Picture = LoadPicture(VBEPerson(1, 1, 1, 5, 5))
                FormMainMode.PEAScardcom(i).Picture = LoadPicture(VBEPerson(2, 1, 1, 5, 5))
                FormMainMode.PEASusbi1(i).Caption = VBEPerson(1, 1, 1, 3, 1)
                FormMainMode.PEASusbi2(i).Caption = VBEPerson(1, 1, 1, 3, 2)
                FormMainMode.PEASusbi3(i).Caption = VBEPerson(1, 1, 1, 3, 3)
                FormMainMode.PEAScardcompi1(i).Caption = VBEPerson(2, 1, 1, 3, 1)
                FormMainMode.PEAScardcompi2(i).Caption = VBEPerson(2, 1, 1, 3, 2)
                FormMainMode.PEAScardcompi3(i).Caption = VBEPerson(2, 1, 1, 3, 3)
            Next
       Case 3
            For i = 1 To 角色人物對戰人數(1, 1)
                FormMainMode.cardusname(i).Caption = VBEPerson(1, i, 1, 1, 1)
                FormMainMode.cardusspname(i).Caption = VBEPerson(1, i, 1, 1, 3)
                FormMainMode.PEAScardus(i).Picture = LoadPicture(VBEPerson(1, i, 1, 5, 5))
                FormMainMode.PEASusbi1(i).Caption = VBEPerson(1, i, 1, 3, 1)
                FormMainMode.PEASusbi2(i).Caption = VBEPerson(1, i, 1, 3, 2)
                FormMainMode.PEASusbi3(i).Caption = VBEPerson(1, i, 1, 3, 3)
            Next
            For i = 1 To 角色人物對戰人數(2, 1)
                FormMainMode.cardcomname(i).Caption = VBEPerson(2, i, 1, 1, 1)
                FormMainMode.cardcomspname(i).Caption = VBEPerson(2, i, 1, 1, 3)
                FormMainMode.PEAScardcom(i).Picture = LoadPicture(VBEPerson(2, i, 1, 5, 5))
                FormMainMode.PEAScardcompi1(i).Caption = VBEPerson(2, i, 1, 3, 1)
                FormMainMode.PEAScardcompi2(i).Caption = VBEPerson(2, i, 1, 3, 2)
                FormMainMode.PEAScardcompi3(i).Caption = VBEPerson(2, i, 1, 3, 3)
            Next
    End Select
    '===========================戰鬥系統主表單讀入(測)
    執行階段系統類.執行階段系統遊戲初始總程序
    戰鬥系統類.遊戲角色卡片物件創立
    '===========================
    For n = 1 To 4
        If VBEPerson(1, 1, 3, n, 1) = "" Then
           FormMainMode.personatk(n).Visible = False
        Else
           FormMainMode.personatk(n).Caption = VBEPerson(1, 1, 3, n, 1)
           FormMainMode.personatk(n).Visible = True
        End If
        '=============
        If VBEPerson(2, 1, 3, n, 1) = "" Then
           FormMainMode.comaiatk(n).Visible = False
        Else
           FormMainMode.comaiatk(n).Caption = VBEPerson(2, 1, 3, n, 1)
           FormMainMode.comaiatk(n).Visible = True
        End If
    Next
    FormMainMode.PEAFInterface.Passive_技能一方全重設 = 1
    FormMainMode.PEAFInterface.Passive_技能一方全重設 = 2
    For n = 5 To 8
        If VBEPerson(1, 1, 3, n, 1) = "" Then
           FormMainMode.PEAFInterface.Passive_使用者_技能隱藏 = n - 4
        Else
           FormMainMode.PEAFInterface.Passive_使用者_技能名稱 = VBEPerson(1, 1, 3, n, 1) & "#" & n - 4
           FormMainMode.PEAFInterface.Passive_使用者_技能顯示 = n - 4
        End If
        '=============
        If VBEPerson(2, 1, 3, n, 1) = "" Then
           FormMainMode.PEAFInterface.Passive_電腦_技能隱藏 = n - 4
        Else
           FormMainMode.PEAFInterface.Passive_電腦_技能名稱 = VBEPerson(2, 1, 3, n, 1) & "#" & n - 4
           FormMainMode.PEAFInterface.Passive_電腦_技能顯示 = n - 4
        End If
    Next
    '===============================對戰圖片載入(隨機組合)-前置階段
    If Formsetting.BGM選擇.Text = "《隨機-地圖組合》" Then
        If Formsetting.對戰地圖選擇.Text = "《隨機》" Then
            Randomize
            m = Int(Rnd() * Val(Formsetting.對戰地圖選擇.ListCount - 1)) + 1
            Formsetting.對戰地圖選擇.ListIndex = m
        Else
            Formsetting.對戰地圖選擇_Click
        End If
     End If
    '===============================對戰音樂載入
    FormMainMode.cMusicPlayer(0).MusicStop
    If Formsetting.BGM選擇.Text = "《隨機》" Then
        Randomize
        m = Int(Rnd() * Val(Formsetting.BGM選擇.ListCount - 1)) + 1
        Formsetting.BGM選擇.ListIndex = m
    End If
       Select Case Formsetting.BGM選擇.Text
         Case "舊版"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulbgm04.mp3"
         Case "冰封湖畔(新)"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm003.mp3"
         Case "人魂墓地"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm004.mp3"
         Case "萊丁貝魯格城堡"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm000.mp3"
         Case "誘惑森林"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm001.mp3"
         Case "垃圾之街"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm002.mp3"
         Case "盡頭之村"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm005.mp3"
         Case "風暴荒野"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm006.mp3"
         Case "魔都羅占布爾克"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm008.mp3"
         Case "烏波斯的黑湖"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm010.mp3"
         Case "藩骸兒的遺跡"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm007.mp3"
         Case "瘋狂山脈"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm009.mp3"
         Case "《其他》"
           FormMainMode.cMusicPlayer(0).Filepath = Formsetting.lopnmusictext.Caption
        End Select
      '===========================對戰圖片載入
    If Formsetting.對戰地圖選擇.Text = "《隨機》" Then
        Randomize
        m = Int(Rnd() * Val(Formsetting.對戰地圖選擇.ListCount - 1)) + 1
        Formsetting.對戰地圖選擇.ListIndex = m
    End If
    If Formsetting.對戰地圖選擇.Text = "《其他》" Then
        FormMainMode.PEAttackingForm.Picture = LoadPicture(Formsetting.lopnmapjpgtext.Caption)
        FormMainMode.PEAttackingStartForm.Picture = LoadPicture(Formsetting.lopnmapjpgtext.Caption)
        FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(Formsetting.lopnmapjpgtext.Caption)
    Else
        FormMainMode.PEAttackingForm.Picture = LoadPicture(app_path & "gif\system\map\" & Formsetting.對戰地圖選擇.ListIndex & ".jpg")
        FormMainMode.PEAttackingStartForm.Picture = LoadPicture(app_path & "gif\system\map\" & Formsetting.對戰地圖選擇.ListIndex & ".jpg")
        FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(app_path & "gif\system\map\" & Formsetting.對戰地圖選擇.ListIndex & ".jpg")
    End If
    '=============================================
    If Formsetting.chkusenewpage.Value = 1 Then
        戰鬥系統類.公用牌地圖牌種類配置 Formsetting.對戰地圖選擇.Text
    Else
        戰鬥系統類.公用牌地圖牌種類配置 0
    End If
    '=============================================
    FormMainMode.cMusicPlayer(0).IsLoop = True
    '=================================================
    戰鬥擲骰介面人物立繪圖路徑紀錄數(1) = VBEPerson(1, 1, 1, 5, 3)
    戰鬥擲骰介面人物立繪圖路徑紀錄數(2) = VBEPerson(2, 1, 1, 5, 3)
    '=================================================
    If Formsetting.chkusenewinterface.Value = 1 Then
        系統顯示界面紀錄數 = 2
    Else
        系統顯示界面紀錄數 = 1
    End If
    '=======================================事件卡/實體卡物件設定
    一般系統類.自由戰鬥模式設定表單各式設定讀入程序
    戰鬥系統類.遊戲實體牌物件宣告程序
    戰鬥系統類.事件卡處理_計算張數
    戰鬥系統類.事件卡處理_初始_使用者方
    戰鬥系統類.事件卡處理_初始_電腦方
    戰鬥系統類.事件卡處理_指定_使用者方
    戰鬥系統類.事件卡處理_指定_電腦方
    '===============================================
    一般系統類.戰鬥系統表單讀入程序
    一般系統類.戰鬥系統開始表單讀入程序
    '===================
    一般系統類.主選單_PEAttackingStartForm顯示
    FormMainMode.cMusicPlayer(0).MusicPlay
    FormMainMode.PEGameFreeModeSettingForm.Visible = False
End If
End Sub
Function 角色人物隨機相似等級選擇(ByVal uscom As Integer, ByVal num As Integer) As Integer
Dim levelmark As Integer
Dim tempnum As Integer
Dim personlist As ComboBox
Dim temppass As Boolean
levelmark = Formsetting.cbsimilarlevel.ListIndex
tempnum = 1
Select Case uscom
    Case 1
        Set personlist = FormMainMode.personlevelus(num)
    Case 2
        Set personlist = FormMainMode.personlevelcom(num)
End Select
temppass = False
Do
    If temppass = False Then
        For i = 0 To personlist.ListCount - 1
            If personlist.List(i) = Formsetting.cbsimilarlevel.List(levelmark) Then
                角色人物隨機相似等級選擇 = i
                Exit Function
            End If
        Next
    End If
    Select Case tempnum
        Case 1
            levelmark = Formsetting.cbsimilarlevel.ListIndex - 1
        Case 2
            levelmark = Formsetting.cbsimilarlevel.ListIndex + 1
        Case 3
            levelmark = Formsetting.cbsimilarlevel.ListIndex - 2
        Case 4
            levelmark = Formsetting.cbsimilarlevel.ListIndex + 2
    End Select
    tempnum = tempnum + 1
    If levelmark > Formsetting.cbsimilarlevel.ListCount - 1 Or levelmark < 0 Then temppass = True Else temppass = False
Loop Until (tempnum > 4)

'=======未找到任何相似之等級
角色人物隨機相似等級選擇 = -1
End Function
Function 角色人物選擇空值檢查(ByVal personvs As Integer) As Boolean
角色人物選擇空值檢查 = True
Dim i As Integer
Select Case personvs
   Case 1
      If FormMainMode.personnameus(1).Text = "" Or FormMainMode.personnamecom(1).Text = "" Then
          角色人物選擇空值檢查 = False
      Else
          角色人物選擇空值檢查 = True
          角色人物對戰人數(1, 1) = 1
          角色人物對戰人數(2, 1) = 1
      End If
   Case 3
      角色人物對戰人數(1, 1) = 3
      角色人物對戰人數(2, 1) = 3
      For i = 1 To 3
          If FormMainMode.personnameus(i).Text = "" Or FormMainMode.personnamecom(i).Text = "" Then
              角色人物選擇空值檢查 = False
'              Exit Function
          End If
       Next
       If i > 3 And 角色人物選擇空值檢查 = False Then
          If FormMainMode.personnameus(1).Text = "" Or FormMainMode.personnamecom(1).Text = "" Then
              '======不符至少有1位之規定
              Exit Function
          End If
          '===========檢查空值並隱藏
           For i = 2 To 3
              If FormMainMode.personnameus(i).Text = "" Then
                  角色人物對戰人數(1, 1) = 角色人物對戰人數(1, 1) - 1
              End If
              If FormMainMode.personnamecom(i).Text = "" Then
                  角色人物對戰人數(2, 1) = 角色人物對戰人數(2, 1) - 1
              End If
           Next
           角色人物選擇空值檢查 = True
        Else
           角色人物選擇空值檢查 = True
        End If
End Select
End Function
Sub 卡片人物資訊載入_搜尋檔案()
Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
  mypath = App.Path & "\character\"
  mydir = Dir(mypath, vbDirectory) ' 找尋第一個子目錄。
  ReDim DirectoryBuff(0)
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
                    If 測試1.GetExtName(mydir) = "uleci" Then
                        人物系統類.卡片人物資訊讀入_初階段 mypath + mydir
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
'MsgBox "1-5-2-4"
End Sub
Function 事件卡資料庫(ByVal name As String, ByVal 目標 As Integer) As String
'目標:1-傳回分類號/2-傳回事件卡檔案名稱/3-傳回類型數值
Select Case name
    Case "劍1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks1"
           Case 3
              事件卡資料庫 = "ATK-劍=1=ATK-劍=1"
        End Select
    Case "劍2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks2"
           Case 3
              事件卡資料庫 = "ATK-劍=2=ATK-劍=2"
        End Select
    Case "劍3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks3"
           Case 3
              事件卡資料庫 = "ATK-劍=3=ATK-劍=3"
        End Select
   Case "劍4"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks4"
           Case 3
              事件卡資料庫 = "ATK-劍=4=ATK-劍=4"
        End Select
    Case "劍5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks5"
           Case 3
              事件卡資料庫 = "ATK-劍=5=ATK-劍=5"
        End Select
    Case "劍6"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks6"
           Case 3
              事件卡資料庫 = "ATK-劍=6=ATK-劍=6"
        End Select
    Case "劍7"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks7"
           Case 3
              事件卡資料庫 = "ATK-劍=7=ATK-劍=7"
        End Select
    Case "劍8"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks8"
           Case 3
              事件卡資料庫 = "ATK-劍=8=ATK-劍=8"
        End Select
    Case "槍1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atkg1"
           Case 3
              事件卡資料庫 = "ATK-槍=1=ATK-槍=1"
        End Select
    Case "槍2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atkg2"
           Case 3
              事件卡資料庫 = "ATK-槍=2=ATK-槍=2"
        End Select
    Case "槍3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atkg3"
           Case 3
              事件卡資料庫 = "ATK-槍=3=ATK-槍=3"
        End Select
    Case "槍4"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg4"
           Case 3
              事件卡資料庫 = "ATK-槍=4=ATK-槍=4"
        End Select
    Case "槍5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg5"
           Case 3
              事件卡資料庫 = "ATK-槍=5=ATK-槍=5"
        End Select
    Case "槍6"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg6"
           Case 3
              事件卡資料庫 = "ATK-槍=6=ATK-槍=6"
        End Select
    Case "槍7"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg7"
           Case 3
              事件卡資料庫 = "ATK-槍=7=ATK-槍=7"
        End Select
    Case "槍8"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg8"
           Case 3
              事件卡資料庫 = "ATK-槍=8=ATK-槍=8"
        End Select
    Case "防1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-def1"
           Case 3
              事件卡資料庫 = "DEF=1=DEF=1"
        End Select
    Case "防2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-def2"
           Case 3
              事件卡資料庫 = "DEF=2=DEF=2"
        End Select
    Case "防3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-def3"
           Case 3
              事件卡資料庫 = "DEF=3=DEF=3"
        End Select
    Case "防4"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-def4"
           Case 3
              事件卡資料庫 = "DEF=4=DEF=4"
        End Select
    Case "防5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-def5"
           Case 3
              事件卡資料庫 = "DEF=5=DEF=5"
        End Select
    Case "防7"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-def7"
           Case 3
              事件卡資料庫 = "DEF=7=DEF=7"
        End Select
    Case "特1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-spe1"
           Case 3
              事件卡資料庫 = "SPE=1=SPE=1"
        End Select
    Case "特2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-spe2"
           Case 3
              事件卡資料庫 = "SPE=2=SPE=2"
        End Select
    Case "特3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 7
           Case 2
              事件卡資料庫 = "7-spe3"
           Case 3
              事件卡資料庫 = "SPE=3=SPE=3"
        End Select
    Case "特4"
       Select Case 目標
           Case 1
              事件卡資料庫 = 7
           Case 2
              事件卡資料庫 = "7-spe4"
           Case 3
              事件卡資料庫 = "SPE=4=SPE=4"
        End Select
    Case "特5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 7
           Case 2
              事件卡資料庫 = "7-spe5"
           Case 3
              事件卡資料庫 = "SPE=5=SPE=5"
        End Select
    Case "移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-mov1"
           Case 3
              事件卡資料庫 = "MOV=1=MOV=1"
        End Select
    Case "移2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 4
           Case 2
              事件卡資料庫 = "4-mov2"
           Case 3
              事件卡資料庫 = "MOV=2=MOV=2"
        End Select
    Case "移3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 4
           Case 2
              事件卡資料庫 = "4-mov3"
           Case 3
              事件卡資料庫 = "MOV=3=MOV=3"
        End Select
    Case "移4"
       Select Case 目標
           Case 1
              事件卡資料庫 = 4
           Case 2
              事件卡資料庫 = "4-mov4"
           Case 3
              事件卡資料庫 = "MOV=4=MOV=4"
        End Select
    Case "移5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 4
           Case 2
              事件卡資料庫 = "4-mov5"
           Case 3
              事件卡資料庫 = "MOV=5=MOV=5"
        End Select
    Case "機會1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-draw1"
           Case 3
              事件卡資料庫 = "DRAW=1=DRAW=1"
        End Select
    Case "機會2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 5
           Case 2
              事件卡資料庫 = "5-draw2"
           Case 3
              事件卡資料庫 = "DRAW=2=DRAW=2"
        End Select
    Case "機會3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 5
           Case 2
              事件卡資料庫 = "5-draw3"
           Case 3
              事件卡資料庫 = "DRAW=3=DRAW=3"
        End Select
    Case "機會4"
       Select Case 目標
           Case 1
              事件卡資料庫 = 5
           Case 2
              事件卡資料庫 = "5-draw4"
           Case 3
              事件卡資料庫 = "DRAW=4=DRAW=4"
        End Select
    Case "機會5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 5
           Case 2
              事件卡資料庫 = "5-draw5"
           Case 3
              事件卡資料庫 = "DRAW=5=DRAW=5"
        End Select
    Case "詛咒術1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 6
           Case 2
              事件卡資料庫 = "6-brk1"
           Case 3
              事件卡資料庫 = "BRK=1=BRK=1"
        End Select
    Case "詛咒術2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 6
           Case 2
              事件卡資料庫 = "6-brk2"
           Case 3
              事件卡資料庫 = "BRK=2=BRK=2"
        End Select
    Case "詛咒術3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 6
           Case 2
              事件卡資料庫 = "6-brk3"
           Case 3
              事件卡資料庫 = "BRK=3=BRK=3"
        End Select
    Case "詛咒術5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 6
           Case 2
              事件卡資料庫 = "6-brk5"
           Case 3
              事件卡資料庫 = "BRK=5=BRK=5"
        End Select
    Case "HP回復1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-hp1"
           Case 3
              事件卡資料庫 = "HPL=1=HPL=1"
        End Select
    Case "HP回復2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-hp2"
           Case 3
              事件卡資料庫 = "HPL=2=HPL=2"
        End Select
    Case "HP回復3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-hp3"
           Case 3
              事件卡資料庫 = "HPL=3=HPL=3"
        End Select
    Case "劍3/槍1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks3g1"
           Case 3
              事件卡資料庫 = "ATK-劍=3=ATK-槍=1"
        End Select
    Case "劍4/槍2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks4g2"
           Case 3
              事件卡資料庫 = "ATK-劍=4=ATK-槍=2"
        End Select
    Case "劍5/槍3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks5g3"
           Case 3
              事件卡資料庫 = "ATK-劍=5=ATK-槍=3"
        End Select
    Case "槍3/劍1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atkg3s1"
           Case 3
              事件卡資料庫 = "ATK-槍=3=ATK-劍=1"
        End Select
    Case "槍4/劍2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg4s2"
           Case 3
              事件卡資料庫 = "ATK-槍=4=ATK-劍=2"
        End Select
    Case "槍5/劍3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg5s3"
           Case 3
              事件卡資料庫 = "ATK-槍=5=ATK-劍=3"
        End Select
    Case "防3/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-def3mov1"
           Case 3
              事件卡資料庫 = "DEF=3=MOV=1"
        End Select
    Case "防4/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-def4mov1"
           Case 3
              事件卡資料庫 = "DEF=4=MOV=1"
        End Select
    Case "防5/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 3
           Case 2
              事件卡資料庫 = "3-def5mov1"
           Case 3
              事件卡資料庫 = "DEF=5=MOV=1"
        End Select
    Case "特1/防1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-spe1def1"
           Case 3
              事件卡資料庫 = "SPE=1=DEF=1"
        End Select
    Case "特2/防2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 7
           Case 2
              事件卡資料庫 = "7-spe2def2"
           Case 3
              事件卡資料庫 = "SPE=2=DEF=2"
        End Select
    Case "特3/防3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 7
           Case 2
              事件卡資料庫 = "7-spe3def3"
           Case 3
              事件卡資料庫 = "SPE=3=DEF=3"
        End Select
    Case "劍3/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks3mov1"
           Case 3
              事件卡資料庫 = "ATK-劍=3=MOV=1"
        End Select
    Case "劍4/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks4mov1"
           Case 3
              事件卡資料庫 = "ATK-劍=4=MOV=1"
        End Select
    Case "劍5/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 1
           Case 2
              事件卡資料庫 = "1-atks5mov1"
           Case 3
              事件卡資料庫 = "ATK-劍=5=MOV=1"
        End Select
    Case "槍3/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atkg3mov1"
           Case 3
              事件卡資料庫 = "ATK-槍=3=MOV=1"
        End Select
    Case "槍4/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg4mov1"
           Case 3
              事件卡資料庫 = "ATK-槍=4=MOV=1"
        End Select
    Case "槍5/移1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 2
           Case 2
              事件卡資料庫 = "2-atkg5mov1"
           Case 3
              事件卡資料庫 = "ATK-槍=5=MOV=1"
        End Select
    Case "劍3/防1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks3def1"
           Case 3
              事件卡資料庫 = "ATK-劍=3=DEF=1"
        End Select
    Case "槍3/防1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atkg3def1"
           Case 3
              事件卡資料庫 = "ATK-槍=3=DEF=1"
        End Select
    Case "移1/特1"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-mov1spe1"
           Case 3
              事件卡資料庫 = "MOV=1=SPE=1"
        End Select
    Case "移2/特2"
       Select Case 目標
           Case 1
              事件卡資料庫 = 4
           Case 2
              事件卡資料庫 = "4-mov2spe2"
           Case 3
              事件卡資料庫 = "MOV=2=SPE=2"
        End Select
    Case "移3/特3"
       Select Case 目標
           Case 1
              事件卡資料庫 = 4
           Case 2
              事件卡資料庫 = "4-mov3spe3"
           Case 3
              事件卡資料庫 = "MOV=3=SPE=3"
        End Select
    Case "劍5/槍5"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-atks5g5"
           Case 3
              事件卡資料庫 = "ATK-劍=5=ATK-槍=5"
        End Select
    Case "聖水"
       Select Case 目標
           Case 1
              事件卡資料庫 = 0
           Case 2
              事件卡資料庫 = "0-hpw"
           Case 3
              事件卡資料庫 = "HPW=1=HPW=1"
        End Select
    Case Else
        Select Case 目標
           Case 1
              事件卡資料庫 = 99
        End Select
End Select
End Function
Sub 戰鬥系統表單讀入程序()
Dim 暫時數(2) As Integer '按鈕座標暫時變數
Dim i As Integer, ckl As Integer, mm As Integer, w As Integer  '暫時變數
'    app_path = App.Path
'    If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'------------
goidefus = 0
movecp = 2
turnpageonin = 0
trend暫時變數 = 0
FormMainMode.PEAFInterface.MessageClear
'----------
'For i = 1 To 公用牌實體卡片分隔紀錄數(1)
' FormMainMode.cge(i).Visible = False
' FormMainMode.cgen(i).Visible = False
' FormMainMode.cqe(i).Visible = False
' FormMainMode.cqen(i).Visible = False
' FormMainMode.cgu(i).Visible = False
' FormMainMode.cqu(i).Visible = False
' FormMainMode.card(i).BackColor = RGB(0, 0, 0)
'Next

'----------------------------寫入技能欄座標
atkinghelpxy(1, 1, 1) = 2520
atkinghelpxy(1, 2, 1) = 4730
atkinghelpxy(1, 3, 1) = 6930
atkinghelpxy(1, 4, 1) = 9140

atkinghelpxy(2, 1, 1) = 7560
atkinghelpxy(2, 2, 1) = 5400
atkinghelpxy(2, 3, 1) = 3240
atkinghelpxy(2, 4, 1) = 1080
For i = 1 To 4
   atkinghelpxy(1, i, 2) = 3480
   atkinghelpxy(2, i, 2) = 960
Next
'-----------------------
For ckl = 1 To 公用牌實體卡片分隔紀錄數(1)
    FormMainMode.card(ckl).Visible = False
    FormMainMode.card(ckl).CardEnabledType = True
Next
'-------以下是設計物件顯示
If Formsetting.checktest.Value = 0 Then
    FormMainMode.pageusglead.Visible = False
    FormMainMode.pageusqlead.Visible = False
    FormMainMode.pagecomglead.Visible = False
    FormMainMode.pagecomqlead.Visible = False
    FormMainMode.Command1.Visible = False
    FormMainMode.Command2.Visible = False
    FormMainMode.影子設定.Visible = False
End If

FormMainMode.小人物角色基準線.Visible = False
For i = 1 To 6
   FormMainMode.小人物距離基準線(i).Visible = False
Next

If Formsetting.checktestpersondown.Value = 1 Then
    FormMainMode.影子設定.Visible = True
End If

If Formsetting.大亂鬥選項.Value = 0 Then
    暫時數(1) = 8400
    暫時數(2) = 7200
Else
   暫時數(1) = 9360
   暫時數(2) = 6720
End If
FormMainMode.cn1.Top = 暫時數(1)
FormMainMode.cnmove.Top = 暫時數(1)
FormMainMode.cnmove2.Top = 暫時數(1)
FormMainMode.cn2.Top = 暫時數(1)
FormMainMode.cn22.Top = 暫時數(1)
FormMainMode.cn3.Top = 暫時數(1)
FormMainMode.cn32.Top = 暫時數(1)
FormMainMode.cn4.Top = 暫時數(1)
FormMainMode.cn1.Left = 暫時數(2)
FormMainMode.cnmove.Left = 暫時數(2)
FormMainMode.cnmove2.Left = 暫時數(2)
FormMainMode.cn2.Left = 暫時數(2)
FormMainMode.cn22.Left = 暫時數(2)
FormMainMode.cn3.Left = 暫時數(2)
FormMainMode.cn32.Left = 暫時數(2)
FormMainMode.cn4.Left = 暫時數(2)
'=====================以下是技能欄顏色顯示
For i = 1 To 4
    FormMainMode.personatk(i).ForeColor = RGB(192, 192, 192)
    FormMainMode.personatk(i).BackColor = RGB(0, 0, 0)
    FormMainMode.comaiatk(i).ForeColor = RGB(192, 192, 192)
    FormMainMode.comaiatk(i).BackColor = RGB(0, 0, 0)
Next
'=====================
FormMainMode.顯示列1.goi1顯示 = False
FormMainMode.顯示列1.goi2顯示 = False
FormMainMode.cn2.Visible = False
FormMainMode.cnmove.Visible = False
FormMainMode.cnmove2.Visible = False
FormMainMode.cn22.Visible = False
FormMainMode.cn3.Visible = False
FormMainMode.cn32.Visible = False
FormMainMode.cn4.Visible = False
FormMainMode.atkinghelpc.Visible = False
'For i = 1 To 42
'   FormMainMode.personusspe(i).Visible = False
'   FormMainMode.personcomspe(i).Visible = False
'Next
'================================
'For i = 1 To 3
'    FormMainMode.cardbackus(i).大人物圖片 = app_path & "gif\system\cardblack.png"
'    FormMainMode.cardbackcom(i).大人物圖片 = app_path & "gif\system\cardblack.png"
'Next
'---------以下是設定技能字體大小
For i = 1 To 4
    If Val(VBEPerson(1, 1, 2, 3, 5)) = 0 Then
       FormMainMode.personatk(i).FontSize = VBEPerson(1, 1, 2, 3, i)
    Else
       FormMainMode.personatk(i).FontSize = 12
    End If
    If Val(VBEPerson(2, 1, 2, 3, 5)) = 0 Then
       FormMainMode.comaiatk(i).FontSize = VBEPerson(2, 1, 2, 3, i)
    Else
       FormMainMode.comaiatk(i).FontSize = 12
    End If
Next
'===========================系統顯示介面設定
If 系統顯示界面紀錄數 = 1 Then
    FormMainMode.PEAFInterface.Passive_介面顯示 = False
    FormMainMode.cardpagejpg.Visible = True
    FormMainMode.pageul.Visible = True
    FormMainMode.draw1.Visible = True
    FormMainMode.move1.Visible = True
    FormMainMode.move3.Visible = True
    FormMainMode.move4.Visible = True
    FormMainMode.draw2.Visible = False
    FormMainMode.atkdef1.Visible = False
    FormMainMode.atkdef2.Visible = False
    FormMainMode.move2.Visible = False
Else
    FormMainMode.PEAFInterface.Passive_介面顯示 = True
    FormMainMode.PEAFInterface.stagejpg = app_path & "gif\system\stageblack.gif"
    FormMainMode.cardpagejpg.Visible = False
    FormMainMode.pageul.Visible = False
    FormMainMode.draw1.Visible = False
    FormMainMode.move1.Visible = False
    FormMainMode.move3.Visible = False
    FormMainMode.move4.Visible = False
    FormMainMode.draw2.Visible = False
    FormMainMode.atkdef1.Visible = False
    FormMainMode.atkdef2.Visible = False
    FormMainMode.move2.Visible = False
End If
FormMainMode.顯示列1.顯示列圖片 = app_path & "gif\system\DRAW.png"
'-----------以下為牌組初始化
If Formsetting.大亂鬥選項.Value = 1 Then
    For mm = 1 To 3
       牌總階段數(mm) = Val(Formsetting.大亂鬥模式選項_牌數.Text)
       If 牌總階段數(mm) < 1 Then 牌總階段數(mm) = 1
    Next
ElseIf Formsetting.挑戰模式選項.Value = 1 Then
   For mm = 2 To 3
       牌總階段數(mm) = 5 + Val(Formsetting.挑戰模式選項_牌數.Text)
   Next
    牌總階段數(1) = 5
Else
    For mm = 1 To 3
       牌總階段數(mm) = 5
    Next
End If

'=============測試用途
'牌總階段數(1) = 57
'牌總階段數(2) = 0
'牌總階段數(3) = 57
'==================牌設定(重設)
For mm = 1 To 公用牌實體卡片分隔紀錄數(1)
   pagecardnum(mm, 6) = 4
Next
戰鬥系統類.公用牌未使用檢查
BattleCardNum = 公用牌各牌類型紀錄數(0, 2)
戰鬥系統類.執行動作_系統總卡牌張數更新
階段狀態數 = 0
FormMainMode.pageusqlead.Caption = 0
FormMainMode.pageusglead.Caption = 0
FormMainMode.pagecomqlead.Caption = 0
FormMainMode.pagecomglead.Caption = 0
'=======以下為角色人物設定(總初設)
FormMainMode.uspiin(1).Left = 0
FormMainMode.uspiin(1).Visible = False
FormMainMode.cardus(1).Left = 0
FormMainMode.cardus(1).Top = 6240
FormMainMode.cardus(1).ZOrder
FormMainMode.cardus(1).Visible = True
FormMainMode.compiin(1).Left = 0
'=======
For i = 2 To 3
   If 角色人物對戰人數(1, 1) >= i Then
       FormMainMode.uspiin(i).Visible = True
       FormMainMode.uspiin(i).Left = 2520 * (i - 1)
       FormMainMode.uspiin(i).Visible = True
       FormMainMode.cardus(i).Visible = False
   Else
       FormMainMode.uspiin(i).Visible = False
       FormMainMode.uspi4(i).Caption = 0
'       FormMainMode.usbi1(i).Caption = 0
       FormMainMode.cardus(i).CardMain_角色HP = 0
   End If
   If 角色人物對戰人數(2, 1) >= i Then
       FormMainMode.compiin(i).Visible = True
       FormMainMode.compiin(i).Left = 2520 * (i - 1)
   Else
       FormMainMode.compiin(i).Visible = False
       FormMainMode.compi4(i).Caption = 0
'       FormMainMode.cardcompi1(i).Caption = 0
       FormMainMode.cardcom(i).CardMain_角色HP = 0
   End If
Next
If 角色人物對戰人數(1, 1) > 1 Or 角色人物對戰人數(2, 1) > 1 Then
   FormMainMode.顯示列1.人物戰鬥人數 = 3
Else
   FormMainMode.顯示列1.人物戰鬥人數 = 1
End If
For w = 1 To 3
   角色待機人物紀錄數(1, w) = w
   角色待機人物紀錄數(2, w) = w
Next
'=======以下為角色人物設定(使用者)
角色人物對戰人數(1, 2) = 1
For w = 1 To 角色人物對戰人數(1, 1)
    FormMainMode.cardus(w).CardMain_角色圖片 = VBEPerson(1, w, 1, 5, 5)
    atkus(w) = Val(VBEPerson(1, w, 1, 3, 2))
    defus(w) = Val(VBEPerson(1, w, 1, 3, 3))
    liveus(w) = Val(VBEPerson(1, w, 1, 3, 1))
    uslevel(w) = Val(VBEPerson(1, w, 1, 2, 2))
    nameus(w) = VBEPerson(1, w, 1, 1, 1)
'    FormMainMode.usbi1(w).Caption = liveus(w)
'    FormMainMode.usbi2(w).Caption = atkus(w)
'    FormMainMode.usbi3(w).Caption = defus(w)
    FormMainMode.cardus(w).CardMain_角色HP = liveus(w)
    FormMainMode.cardus(w).CardMain_角色ATK = atkus(w)
    FormMainMode.cardus(w).CardMain_角色DEF = defus(w)
    FormMainMode.uspi1(w).Caption = nameus(w)
    liveusmax(w) = liveus(w)
    FormMainMode.cardus(w).CardMain_角色HPMAX = liveusmax(w)
    liveus41(w) = liveusmax(w) \ 3
    FormMainMode.uspi2(w).Caption = uslevel(w)
    FormMainMode.uspiatk(w).Caption = atkus(w)
    FormMainMode.uspidef(w).Caption = defus(w)
    FormMainMode.uspi4(w).Caption = liveus(w)
    FormMainMode.uspi5(w).Caption = liveusmax(w)
    '=================
    FormMainMode.cardus(w).Buff_異常狀態_全重設 = True
    FormMainMode.cardus(w).CardBack_全重設 = True
    戰鬥系統類.技能說明載入_人物卡片背面_使用者 w
Next
FormMainMode.bloodnumus1.Caption = liveus(1)
FormMainMode.bloodnumus2.Caption = liveusmax(1)
FormMainMode.personusminijpg.小人物重設 = True
FormMainMode.personusminijpg.小人物圖片 = VBEPerson(1, 1, 1, 5, 1)
FormMainMode.personusminijpg.小人物影子圖片 = VBEPerson(1, 1, 1, 5, 2)
FormMainMode.顯示列1.使用者方小人物圖片 = VBEPerson(1, 1, 1, 5, 4)
FormMainMode.顯示列1.使用者方小人物圖片left = -FormMainMode.顯示列1.使用者方小人物圖片width
FormMainMode.personusminijpg.小人物影子Left = Val(VBEPerson(1, 1, 2, 1, 5))
FormMainMode.personusminijpg.小人物影子top差 = Val(VBEPerson(1, 1, 2, 1, 6))
FormMainMode.personusminijpg.小人物影像反轉 = False
'戰鬥系統類.人物交換_使用者_指定交換 1
'=======以下為角色人物設定(電腦)
角色人物對戰人數(2, 2) = 1
For w = 1 To 角色人物對戰人數(2, 1)
    FormMainMode.cardcom(w).CardMain_角色圖片 = VBEPerson(2, w, 1, 5, 5)
    atkcom(w) = Val(VBEPerson(2, w, 1, 3, 2))
    defcom(w) = Val(VBEPerson(2, w, 1, 3, 3))
    livecom(w) = Val(VBEPerson(2, w, 1, 3, 1))
    comlevel(w) = Val(VBEPerson(2, w, 1, 2, 2))
    namecom(w) = VBEPerson(2, w, 1, 1, 1)
    livecommax(w) = livecom(w)
    FormMainMode.compi2(w).Caption = comlevel(w)
    FormMainMode.compiatk(w).Caption = atkcom(w)
    FormMainMode.compidef(w).Caption = defcom(w)
    FormMainMode.compi4(w).Caption = livecom(w)
    FormMainMode.compi5(w).Caption = livecommax(w)
    FormMainMode.compi1(w).Caption = namecom(w)
    livecom41(w) = livecommax(w) \ 3
'    FormMainMode.cardcompi1(w).Caption = livecom(w)
'    FormMainMode.cardcompi2(w).Caption = atkcom(w)
'    FormMainMode.cardcompi3(w).Caption = defcom(w)
    FormMainMode.cardcom(w).CardMain_角色HP = livecom(w)
    FormMainMode.cardcom(w).CardMain_角色HPMAX = livecommax(w)
    FormMainMode.cardcom(w).CardMain_角色ATK = atkcom(w)
    FormMainMode.cardcom(w).CardMain_角色DEF = defcom(w)
    '=================
    FormMainMode.cardcom(w).Buff_異常狀態_全重設 = True
    FormMainMode.cardcom(w).CardBack_全重設 = True
    戰鬥系統類.技能說明載入_人物卡片背面_電腦 w
Next
FormMainMode.bloodnumcom1.Caption = livecom(1)
FormMainMode.bloodnumcom2.Caption = livecommax(1)
FormMainMode.personcomminijpg.小人物重設 = True
FormMainMode.personcomminijpg.小人物圖片 = VBEPerson(2, 1, 1, 5, 1)
FormMainMode.personcomminijpg.小人物影子圖片 = VBEPerson(2, 1, 1, 5, 2)
FormMainMode.顯示列1.電腦方小人物圖片 = VBEPerson(2, 1, 1, 5, 4)
FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
FormMainMode.personcomminijpg.小人物影子Left = Val(VBEPerson(2, 1, 2, 1, 5))
FormMainMode.personcomminijpg.小人物影子top差 = Val(VBEPerson(2, 1, 2, 1, 6))
FormMainMode.personcomminijpg.小人物影像反轉 = True
'==================執行小人物立繪指定及距離指定
'戰鬥系統類.人物交換_使用者_指定交換 1
執行動作_距離變更 movecp, False
'================仿對戰模式設定
If Formsetting.chkpersonvsmode.Value = 1 Then
    For i = 2 To 3
        FormMainMode.compi1(i).Caption = ""
        FormMainMode.compi2(i).Caption = ""
        FormMainMode.compiatk(i).Caption = ""
        FormMainMode.compidef(i).Caption = ""
        FormMainMode.compi4(i).Caption = ""
        FormMainMode.compi5(i).Caption = ""
'        FormMainMode.cardcompi1(i).Caption = "?"
'        FormMainMode.cardcompi2(i).Caption = "?"
'        FormMainMode.cardcompi3(i).Caption = "?"
        FormMainMode.cardcom(i).CardMain_角色HP = -99
        FormMainMode.cardcom(i).CardMain_角色ATK = -99
        FormMainMode.cardcom(i).CardMain_角色DEF = -99
        FormMainMode.cardcom(i).CardMain_角色圖片 = app_path & "gif\system\personunknown.jpg"
        FormMainMode.cardcom(i).CardBack_全重設 = True
    Next
End If
'--------------------------計算距離單位(HP血條)
距離單位(1, 1, 1) = 5295 \ liveusmax(1)
距離單位(1, 2, 1) = (11340 - 6060) \ livecommax(1)
'==============血量載入動畫設定
FormMainMode.bloodlineout1.Width = 1
FormMainMode.bloodlineout2.Left = 11325
Erase 血量計數器動畫暫時變數
血量計數器動畫暫時變數(1, 1) = 106
血量計數器動畫暫時變數(2, 1) = 106
'==============時間軸顏色設定
戰鬥系統類.時間軸_重設
'===============重設雙方異常狀態設定
'戰鬥系統類.執行動作_清除所有異常狀態_使用者 1
'戰鬥系統類.執行動作_清除所有異常狀態_電腦 1
Erase 人物異常狀態資料庫
'==================
'一般系統類.判斷字型_FormMainMode
'==================
BattleTurn = 1
FormMainMode.PEAFInterface.turn = BattleTurn
End Sub
Sub 自由戰鬥模式設定表單讀入程序()
Dim i, j As Integer
'MsgBox "1-5-2-1"
選單使用者事件 = True
選單電腦事件 = True
FormMainMode.cdgpersonus.Filter = "UnlightVBE 卡片人物資訊檔(*.uleci)|*.uleci"
卡片人物資訊檔案讀取失敗紀錄串 = ""
'======================清除人物相關記錄變數
Erase VBEPerson
Erase VBEVSSAtkingStr
totpersonnumber = 0
總共人物名稱 = ""
總共人物檔案名 = ""
'======================
For i = 1 To 3
    FormMainMode.personnameus(i).Clear
    FormMainMode.personnamecom(i).Clear
    FormMainMode.personlevelus(i).Clear
    FormMainMode.personlevelcom(i).Clear
Next
'MsgBox "1-5-2-2"
一般系統類.卡片人物資訊載入_搜尋檔案
'MsgBox "1-5-2-5"
'===============調整預設
If FormMainMode.personnameus(1).ListCount > 0 Then
    For i = 1 To 3
       FormMainMode.personnameus(i).ListIndex = 0
       FormMainMode.personnamecom(i).ListIndex = 0
    Next
End If
FormMainMode.opnpersonvs(2).Value = True


'一般系統類.判斷字型_formgamesetting
FormMainMode.cMusicPlayer(0).MusicPlay
'一般系統類.檢查音樂播放 0
FormMainMode.personreadifus.Visible = False
'---------以下是設計物件顯示
For i = 1 To 3
    FormMainMode.personsettingus(i).Caption = "人物資訊"
    FormMainMode.personsettingcom(i).Caption = "人物資訊"
    FormMainMode.personsettingus(i).Visible = False
    FormMainMode.personsettingcom(i).Visible = False
Next
'MsgBox "1-5-2-6"
End Sub
Sub 遊戲初始讀入程序()
'=====以下是背景音樂及SE初始設定
    For i = 1 To 8
        Load FormMainMode.cMusicPlayer(i)
    Next
    FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulbgm03.mp3"
    FormMainMode.cMusicPlayer(0).Volume = 50
    FormMainMode.cMusicPlayer(0).IsLoop = True
    For i = 1 To FormMainMode.cMusicPlayer.UBound
          FormMainMode.cMusicPlayer(i).Volume = 45
    Next
End Sub
Sub 主選單_PEStartForm顯示()
FormMainMode.PEStartForm.Left = 0
FormMainMode.PEStartForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10335
FormMainMode.PEStartForm.Visible = True
FormMainMode.PEStartForm.ZOrder
'====================
FormMainMode.PEStext1.Visible = False
tr1num = 0
FormMainMode.tr1.Enabled = True
End Sub
Sub 主選單_PEGameFreeModeSettingForm顯示()
FormMainMode.PEGameFreeModeSettingForm.Left = 0
FormMainMode.PEGameFreeModeSettingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10335
FormMainMode.PEGameFreeModeSettingForm.Visible = True
FormMainMode.PEGameFreeModeSettingForm.Enabled = True
FormMainMode.PEGameFreeModeSettingForm.ZOrder
'===================臨時通知顯示
If 通知表單是否已出現 = False Then
    通知表單是否已出現 = True
    If 卡片人物資訊檔案讀取失敗紀錄串 <> "" Then
        一般系統類.通知表單顯示 1
    End If
End If

End Sub
Sub 主選單_PEAttackingForm顯示()
FormMainMode.PEAttackingForm.Left = 0
FormMainMode.PEAttackingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10335
FormMainMode.PEAttackingForm.Visible = True
FormMainMode.PEAttackingForm.ZOrder
End Sub
Sub 主選單_PEAttackingStartForm顯示()
FormMainMode.PEAttackingStartForm.Left = 0
FormMainMode.PEAttackingStartForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10335
FormMainMode.PEASpersontalk.Visible = False
FormMainMode.PEAttackingStartForm.Visible = True
FormMainMode.PEAttackingStartForm.ZOrder
'=============
FormMainMode.PEASpersontalk.Visible = False
End Sub
Sub 主選單_PEAttackingEndingForm顯示()
FormMainMode.PEAttackingEndingForm.Left = 0
FormMainMode.PEAttackingEndingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10335
FormMainMode.PEAttackingEndingForm.Visible = True
FormMainMode.PEAttackingEndingForm.ZOrder
'================
FormMainMode.bnreturnt.Visible = False
FormMainMode.bnreturn.Visible = False
FormMainMode.bn.Visible = False
FormMainMode.bnt.Visible = False
End Sub
Sub 戰鬥系統開始表單讀入程序()
Dim i As Integer '暫時變數
'======================
FormMainMode.downjpg.Top = Val(FormMainMode.Height)
FormMainMode.upjpg_2.大人物圖片 = App.Path & "\gif\system\startupjpg.png"
'FormMainMode.upjpg.Top = -Val(FormMainMode.upjpg.Height)
FormMainMode.upjpg_2.Top = -Val(FormMainMode.upjpg.Height)
For i = 1 To 3
'   FormMainMode.cardus(i).Top = -Val(FormMainMode.Height)
'   FormMainMode.cardcom(i).Top = -Val(FormMainMode.Height)
   FormMainMode.PEAScardus(i).Top = -Val(FormMainMode.PEAScardus(i).Height)
   FormMainMode.PEAScardcom(i).Top = -Val(FormMainMode.PEAScardcom(i).Height)
Next
FormMainMode.大人物形像_使用者.大人物圖片 = VBEPerson(1, 1, 1, 5, 3)
FormMainMode.大人物形像_使用者.大人物影像反轉 = False
'FormMainMode.大人物形像_使用者.Height = formsettingpersonus.bight.Text
FormMainMode.大人物形像_使用者.Top = 8400 - FormMainMode.大人物形像_使用者.大人物圖片height
'If FormMainMode.大人物形像_使用者.Top < 0 Then FormMainMode.大人物形像_使用者.Top = 0
FormMainMode.大人物形像_使用者.Width = FormMainMode.大人物形像_使用者.大人物圖片width
FormMainMode.大人物形像_使用者.Left = -FormMainMode.大人物形像_使用者.大人物圖片width
FormMainMode.大人物形像_電腦.大人物圖片 = VBEPerson(2, 1, 1, 5, 3)
FormMainMode.大人物形像_電腦.大人物影像反轉 = True
'FormMainMode.大人物形像_電腦.Height = formsettingpersoncom.bight.Text
FormMainMode.大人物形像_電腦.Top = 8400 - FormMainMode.大人物形像_電腦.大人物圖片height
'If FormMainMode.大人物形像_電腦.Top < 0 Then FormMainMode.大人物形像_電腦.Top = 0
FormMainMode.大人物形像_電腦.Width = FormMainMode.大人物形像_電腦.大人物圖片width
FormMainMode.大人物形像_電腦.Left = FormMainMode.ScaleWidth
st = 0
sq = 0
'一般系統類.判斷字型_form8
FormMainMode.start1.Enabled = True
End Sub
Sub 自由戰鬥模式設定表單基本設定程序()
Dim i As Integer '暫時變數
'MsgBox "1-5-3-1"
Formsetting.對戰地圖選擇.ListIndex = 0
Formsetting.BGM選擇.ListIndex = 0
For i = 1 To 18
    Formsetting.personus(i).ListIndex = 0
    Formsetting.personcom(i).ListIndex = 0
Next
Formsetting.persontgruonus(1).Value = True
Formsetting.persontgruoncom(1).Value = True
Formsetting.lopnmusictext.Visible = False
Formsetting.lopnmapjpgtext.Visible = False
Formsetting.ckendturnnum.Text = 18
Formsetting.t1.Tab = 0
Formsetting.chkusenewai.Value = 1
Formsetting.chkusenewpage.Value = 1
Formsetting.chkusenewinterface.Value = 1
'=============================
Formsetting.cbsimilarlevel.AddItem "LV1"
Formsetting.cbsimilarlevel.AddItem "LV2"
Formsetting.cbsimilarlevel.AddItem "LV3"
Formsetting.cbsimilarlevel.AddItem "LV4"
Formsetting.cbsimilarlevel.AddItem "LV5"
Formsetting.cbsimilarlevel.AddItem "R1"
Formsetting.cbsimilarlevel.AddItem "R2"
Formsetting.cbsimilarlevel.AddItem "R3"
Formsetting.cbsimilarlevel.AddItem "R4"
Formsetting.cbsimilarlevel.AddItem "R5"
Formsetting.cbsimilarlevel.AddItem "N1"
Formsetting.cbsimilarlevel.ListIndex = 4
'=============================
If FormMainMode.personsettingus(1).Caption = "人物資訊" Then
'    Formsetting.其他設定.Visible = False
    Formsetting.chkpersonvsmode.Value = 1
    Formsetting.persontgruoncom(4).Value = True
    Formsetting.persontgruonus(4).Value = True
    Formsetting.ckendturn.Value = 1
'    Formsetting.chkusenewaipersonauto.Visible = False
'    FormMainMode.Caption = FormMainMode.Tag & "  [" & Form2.aboutvn.Caption & "]"
End If
'MsgBox "1-5-3-2"
End Sub
Sub 檢查音樂播放(ByVal num As Integer)
'音樂檢查播放目標數 = num
'FormMainMode.PEMtr1.Enabled = True
End Sub
Sub 清除戰鬥系統所有變數值()
'Erase atkingno '技能發動排序暫時圖片路徑儲存變數(技能發動順序8~1,1.圖片路徑/2.(1)使用者/(2)電腦方/3.Left/4.Top(座標)/5.視窗寬度(Width)/6.視窗高度(Height)/7.技能編號/8.技能執行中時啟動值/9.技能執行中換圖片檢查值/10.第2張圖片路徑)
Erase goicheck   '攻擊/防禦模式加骰數值檢查碼
Erase liveus
Erase livecom
Erase liveusmax
Erase livecommax
Erase pageusleadmax  '使用者牌順序計數表(0.手牌/1.出牌)
Erase pagecomleadmax   '電腦牌順序計數表(0.手牌/1.出牌)
Erase pageqlead   '出牌計數變數(1.使用者/2.電腦)
Erase pageglead   '手牌計數變數(1.使用者/2.電腦)
movedsus = 0   '使用者移動階段決定值變數
turnpageonin = 0 '階段是否可出牌變數(一般)
turnpageoninatking = 0 '階段是否可出牌變數(技能使用)
goickus = 0 '牌值一次檢查碼
Erase atkingck   '技能階段啟動碼(x.人物技能編號,1.技能執行階段/2.技能啟動檢查值)
'Erase atkingckai   'AI技能階段啟動碼(x.人物技能編號,1.技能執行階段/2.技能啟動檢查值)
Erase atkingtrn  '技能計數器暫時儲存變數(1.使用者(現)/2.電腦(現)/3.使用者(備份)/4.電腦(備份))
HP檢查變數 = False 'HP檢查階段是否已檢查變數
HP檢查階段數 = 0  'HP檢查階段變數(1.移動階段後,2.攻擊/防禦階段前,3.攻/防禦階段後)
Erase 距離單位  '距離單位暫時儲存資料(1.HP血條/2.牌移動,1.使用者/2.電腦,1.Left單位/2.Top單位)
Erase personminixy '小人物圖片座標指定資料(1.使用者/2.電腦,第n位,1.近距離/2.中距離/3.遠距離,1.Left/2.Top(座標))
Erase 人物異常狀態資料庫 '異常狀態資料(1.使用者/2.電腦,第x個異常狀態,1.狀態數值/2.狀態統計數(剩餘回合/累計)/3.狀態編號)
Erase 異常狀態檢查數 '異常狀態啟動碼(x.異常狀態編號,1.狀態執行階段/2.狀態啟動檢查值)
技能動畫顯示階段數 = 0 '技能動畫計數器階段碼(1.攻擊/防禦階段-普通,2.移動階段-普通/3.發牌階段後、移動階段前/4.移動階段後/5.攻擊階段後/6.防禦階段後/7.回合結束時)
Erase 攻擊防禦骰子總數 '攻擊/防禦模式骰子數量資料(1.使用者(總)/2.電腦(總)/3.使用者(原)/4.電腦(原))
Erase atkingpagetot  '每階段出牌種類及數值統計資料(1.使用者/2.電腦,1.劍/2.防/3.移/4.特/5.槍)
Erase 骰數零檢查值  '當前階段骰子數量是否為零檢查數(1.使用者/2.電腦)
Erase pagecardnum '公用牌資料(第x編號(1~70-公牌/71~88-使用者事件牌/89~106-電腦事件牌),1.正面類型/2.正面數值/3.反面類型/4.反面數值/5.(1)使用者-(2)電腦/6.(1)手牌-(2)出牌-(3)藏牌-(4)牌堆/7.出牌順序/8.圖片編號/9.目前Left(座標)/10.目前Top(座標)/11.(1)電腦方出牌(裏)-(2)電腦發出牌(外))
Erase 牌總階段數 '牌擁有總階段數(1.使用者/2.電腦/3.總計)
Erase 牌移動暫時變數 '牌移動計數器暫時變數(1.Left單位/2.Top單位/3.牌張編號)
Erase 目前數  '總暫時變數
Erase 出牌順序統計暫時變數 '出牌順序統計總暫時資料(1.使用者出牌/2.使用者手牌/3.電腦出牌/4.電腦手牌,第x順序,1.目前牌出牌順序/2.牌張編號)
Erase 距離單位_收牌暫時數  '收牌個別距離單位暫時儲存變數(第x順序,1.Left單位/2.Top單位/3.牌張編號)
階段狀態數 = 0 '每階段開始結束狀態檢查數(1.開始階段/2.結束階段)
Erase 小人物頭像移動方向數   '小人物頭像移動方向狀態數(1.使用者/2.電腦[1.向內,2.向外])
Erase 血量計數器動畫暫時變數 '開始初始階段-血量動畫計數器暫時變數(1.使用者血條/2.電腦血條,1.每次移動量/2.是否已完成)
Erase 時間軸顏色變化紀錄暫時變數 '時間軸進行顏色變化階段紀錄暫時變數(1~3(1)單位變化量(1(1).時間軸(內外))/2.目前累計量/3.目前顏色(R,G,B),4.(1)時間軸(外)階段數-(1)黑變紅-(2)紅變黑/2.目前累計量/3.目前顏色(R))
Erase 開始卡片移動動畫完成數   '開始時每張卡片移動動畫完成紀錄數(1.使用者/2.電腦,1~3.卡片/4.目前第幾張)
Erase 交換角色紀錄暫時變數 '交換角色雙方紀錄暫時數(1.使用者/2.電腦/3.是否當下首次/4.交換角色完執行階段數)
Erase pageeventnum '事件卡排列紀錄資料(1.使用者/2.電腦,1~18-編號,1.事件卡名稱/2.事件卡檔案名稱)
'擲骰後骰傷害數 = 0 '戰鬥系統表單-擲骰表單溝通暫時變數(2)的變數表示
戰鬥模式勝敗紀錄數 = 0 '戰鬥系統當前勝敗紀錄暫時變數(1.使用者方勝利/2.使用者方敗北/3.平手)
電腦方移動階段選擇數 = 0
電腦方事件卡是否出完選擇數 = False
是否系統公骰 = False
Erase 人物卡面背面編號紀錄數  '人物卡片背面技能說明人物編號暫時變數(1.(1).使用者/(2).電腦,2.第n位)
Erase 擲骰表單溝通暫時變數 'Form6表單值溝通暫時變數(1.一回合中先後判斷(1.前/2.後),2.原始骰值(使用者)-擲骰後有效傷害數,3.原始骰值(電腦)-擲骰後傷害對象(1.使用者/2.電腦),4.(1.使用者先攻/2.電腦先攻))
Erase 公用牌各牌類型紀錄數 '各場景公用牌牌類型紀錄暫時變數(0.(1)目前已發牌總數量/(2)目前場景牌總數量,1~31.(1)目前已使用之牌數/(2)該牌型能使用之總數量)
atkingsecondjpg = ""
Erase 公用牌實體卡片分隔紀錄數 '戰鬥系統實體牌相關紀錄數(1.總共牌數/2.公牌牌數/3.使用者事件卡最底編號/4.電腦事件卡最底編號)
Erase 戰鬥擲骰介面人物立繪圖路徑紀錄數 '戰鬥系統擲骰介面雙方人物立繪圖路徑紀錄數(1.使用者方/2.電腦方)
Erase 人物實際狀態資料庫 '人物實際狀態資料
'===================
Erase 事件卡記錄暫時數 '事件卡使用紀錄暫時變數(0.(1)總共給予回合數,1.使用者/2.電腦,1.總共數值/2.目前處理數值/3.目前階段/4.事件卡牌編號/5.事件分類/6.是否啟動)
'Erase 異常狀態_混沌紀錄數 '異常狀態-混沌-骰量紀錄暫時變數(1.紀錄數值(原始)/2.紀錄數值(變更後)/3.數值紀錄是否啟動/4.攻擊防禦模式階段數)
''===================
'atking_sheri_4_tot = 0  '技能-雪莉-飛刃雨出牌量儲存變數
'atking_sheri_4_tot_ai = 0 '技能-AI-雪莉-飛刃雨出牌量儲存變數
'Erase atking_帕茉_慈悲的藍眼_tot  '技能-帕茉-慈悲的藍眼骰子量紀錄暫時變數(1.數值/2.是否啟動)
'Erase atking_艾茵_十三隻眼_tot '技能.艾茵_十三隻眼骰子量紀錄暫時變數(1.數值/2.是否啟動)
'Erase atking_史塔夏_殺戮模式狀態數 '史塔夏殺戮模式狀態檢查數(1.狀態執行階段/2.狀態啟動檢查值/3.紀錄數值(原始)/4.紀錄數值(變更後)/5.數值紀錄是否啟動)
'Erase atking_音音夢_成長模式狀態數 '音音夢成長模式狀態檢查數(1.狀態執行階段/2.狀態啟動檢查值)
'atking_蕾_守護模式狀態啟動值 = False '技能-蕾-Ex-協奏曲-加百烈的守護免除直傷模式啟動值
'Erase atking_羅莎琳_黑霧幻影紀錄狀態數 '技能-羅莎琳-黑霧幻影(普、EX)紀錄對手出牌編號數
'Erase atking_伊芙琳_怠惰的墓表紀錄數 '技能-伊芙琳-怠惰的墓表紀錄對手牌編號暫時數(0.總共張數值/1~2牌編號)
'Erase atking_伊芙琳_赤紅石榴階段紀錄數 '技能-伊芙琳-赤紅石榴紀錄效果及階段暫時數(0.(1).當前效果/(2).當前效果階段,1~106.(1)牌號選定紀錄值)
'Erase atking_古魯瓦爾多_精神力吸收紀錄數 '技能-古魯瓦爾多-精神力吸收紀錄對手牌編號暫時數(0.總共張數值/1~106牌編號選擇值)
'Erase atking_梅倫_Jackpot紀錄數 '技能-梅倫-Jackpot抽牌紀錄數(1.總共數/2.目前數)
'Erase atking_艾伯李斯特_雷擊紀錄數 '技能-艾伯李斯特-雷擊丟棄對手牌紀錄數(1.總共數/2.目前數)
'atking_艾伯李斯特_智略紀錄數 = 0 '技能-艾伯李斯特-智略抽牌目前數
'Erase atking_艾依查庫_神速之劍計算數值紀錄數  '技能-艾依查庫-神速之劍計算劍數值紀錄暫時數(1.目前計算數值/2.(廢除))
'atking_布勞_發條機構紀錄數 = 0 '技能-布勞-發條機構抽牌目前數
'Erase atking_利恩_反擊的狼煙紀錄數 '技能-利恩-反擊的狼煙抽牌目前數(1.總共數/2.目前數)
'Erase atking_夏洛特_大聖堂骰量紀錄數 '技能-夏洛特-大聖堂擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後結果)
'Erase atking_瑪格莉特_月光紀錄數 '技能-瑪格莉特-月光紀錄對手牌編號暫時數(0.目前丟棄張數值/1~106牌編號選擇值/107.總共能丟棄張數值)
'atking_庫勒尼西_瘋狂眼窩紀錄數 = 0 '技能-庫勒尼西-瘋狂眼窩丟棄對手牌紀錄目前數
'Erase atking_傑多_因果之刻記錄數 '技能-傑多-因果之刻紀錄對手出牌編號數(1~106.記錄牌編號/107.總共回張數/108.目前數)
'Erase atking_傑多_因果之幻骰量紀錄數 '技能-傑多-因果之幻擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後結果)
'atking_阿奇波爾多_防護射擊_槍數值紀錄數 = 0 '技能-阿奇波爾多-防護射擊目前累計加槍數值紀錄數
'Erase atking_洛洛妮_逆轉戰局的槍響_抽牌紀錄數 '技能-洛洛妮-逆轉戰局的槍響抽牌目前數(1.總共數/2.目前數)
'atking_洛洛妮_貪婪之刃與嗜血之槍_搶牌紀錄數 = 0 '技能-洛洛妮-貪婪之刃與嗜血之槍搶牌目前數
'Erase atking_克頓_竊取資料_奪牌紀錄數  '技能-克頓-竊取資料奪取對手出牌牌號紀錄數(1.奪牌編號/2.奪牌原方出牌順序)
'Erase atking_克頓_隱蔽射擊骰量紀錄數 '技能-克頓-隱蔽射擊擲骰量紀錄數(1.第1次(公骰)結果/2.第2次(技能)結果/3.分析後總結果)
'Erase atking_克頓_惡意情報紀錄數 '技能-克頓-惡意情報紀錄對手牌編號暫時數(0.目前階段/1~106牌編號選擇值)
'atking_露緹亞_渦騎劍閃計算張數紀錄數 = 0 '技能-露緹亞-渦騎劍閃計算劍卡張數值紀錄暫時數
'atking_艾蕾可_王座之炎計算出牌張數紀錄數 = 0 '技能-艾蕾可-王座之炎計算出牌張數值紀錄暫時數
'Erase atking_艾蕾可_聖王威光紀錄數  '技能-艾蕾可-聖王威光紀錄暫時數(1.對手當回合防禦力/2.對手當回合出牌數/3.使用者當回合攻擊力)
'Erase atking_梅莉_綿羊幻夢_抽牌紀錄數 '技能-梅莉-綿羊幻夢抽牌目前數(1.總共數/2.目前數)
'Erase atking_AI_梅莉_綿羊幻夢_抽牌紀錄數 '技能-AI-梅莉-綿羊幻夢抽牌目前數(1.總共數/2.目前數)
'Erase atking_貝琳達_雪光_抽牌紀錄數  '技能-貝琳達-雪光抽牌目前數(1.總共數/2.目前數)
'Erase atking_貝琳達_水晶幻鏡紀錄狀態數   '技能-貝琳達-水晶幻鏡紀錄對手出牌編號數
'Erase atking_貝琳達_溶魂之雨_攻擊力加成紀錄數  '技能-貝琳達-溶魂之雨攻擊力加成暫時紀錄數(1.是否10張已+10/2.是否15張已+15)
'atking_蕾_終曲_無盡輪迴的終結紀錄數 = 0 '技能-蕾-Ex-終曲-無盡輪迴的終結紀錄對手之防禦牌值暫時數
''================
'Erase 夏洛特_階段處理記錄數 '智慧型AI-夏洛特-戰略判斷紀錄數(1.當前階段實行/2.目標結束之回合數)
vbecommadtotplay = 0
ReDim vbecommadnum(1 To 7, vbecommadtotplay)
ReDim vbecommadstr(1 To 3, vbecommadtotplay)
Erase Vss_PersonAtkingOffNum
Erase Vss_AtkingInformationRecordStr
ReDim VBEStageNum(0) As Integer
End Sub
Sub 清除戰鬥系統開始表單設定值()
Dim i As Integer, j As Integer
For i = 1 To 3
    FormMainMode.PEAScardus(i).Picture = LoadPicture()
    FormMainMode.PEASusbi1(i).Caption = 0
    FormMainMode.PEAScardcom(i).Picture = LoadPicture()
    FormMainMode.PEAScardcompi1(i).Caption = 0
    FormMainMode.cardusname(i).Visible = True
    FormMainMode.cardusspname(i).Visible = True
    FormMainMode.cardcomname(i).Visible = True
    FormMainMode.cardcomspname(i).Visible = True
    For j = 1 To 3
        FormMainMode.compiin(j).Left = 2520 * (j - 1)
        FormMainMode.uspiin(j).Left = 2520 * (j - 1)
    Next
    '================
    If i >= 2 Then
        Formchangeperson.card(i - 1).Visible = True
        Formchangeperson.bnok(i - 1).Visible = True
    End If
Next
End Sub
Sub 自由戰鬥模式設定表單各式設定讀入程序()
Dim ne As Integer, nd As Integer '暫時變數
'========角色色格讀入
For i = 1 To 18
    ne = i Mod 6
    nd = i \ 6
    If ne = 0 Then ne = 6
    If i = 6 Or i = 12 Or i = 18 Then
       nd = (i \ 6) - 1
    End If
    Formsetting.persontgus(i).Caption = Mid(VBEPerson(1, nd + 1, 1, 3, 4), ne, 1)
    Formsetting.persontgcom(i).Caption = Mid(VBEPerson(2, nd + 1, 1, 3, 4), ne, 1)
Next
End Sub
Sub 通知表單顯示(ByVal num As Integer)
Dim pstr() As String
Select Case num
    Case 1
        FormMessage.Label2 = "大小姐您好" & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & "在讀取某些卡片人物資訊檔案時發生了錯誤，" & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & "請大小姐對以下檔案進行個別檢查：" & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & Chr(10)
        pstr = Split(卡片人物資訊檔案讀取失敗紀錄串, "=")
        For k = 1 To UBound(pstr)
             FormMessage.Label2 = FormMessage.Label2 & pstr(k) & Chr(10)
        Next
        For k = 1 To 6
             FormMessage.Label2 = FormMessage.Label2 & Chr(10)
        Next
        FormMessage.Label2 = FormMessage.Label2 & "布勞" & Chr(10)
        FormMessage.Label2.Visible = True
        FormMessage.Text1.Visible = False
        FormMessage.Show 1
End Select
End Sub
Sub 音效播放(ByVal num As Integer)
Select Case num
    Case 1
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse06.mp3"
    Case 2
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse09.mp3"
    Case 3
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse08.mp3"
    Case 4
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse29.mp3"
    Case 5
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse13.mp3"
    Case 6
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse12.mp3"
    Case 7
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse11.mp3"
    Case 8
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse10_f.mp3"
    Case 9
        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse23.mp3"
    Case 10
        FormMainMode.cMusicPlayer(1).Filepath = app_path & "mp3\ulse22.mp3"
        FormMainMode.cMusicPlayer(1).MusicPlay
        Exit Sub
    Case 11
        FormMainMode.cMusicPlayer(3).Filepath = app_path & "mp3\ulse01.mp3"
        FormMainMode.cMusicPlayer(3).MusicPlay
        Exit Sub
End Select
FormMainMode.cMusicPlayer(num).MusicPlay
End Sub
