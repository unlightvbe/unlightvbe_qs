Attribute VB_Name = "一般系統類"
Option Explicit
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
Public ProgramIsOnWine As Boolean '程式是否處於Wine環境下執行

Sub 判斷字型_FormMainMode()
Dim i, a As Integer
a = 14
If FormMainMode.PEStext1.FontName <> "Bradley Gratis" Then
    '===========PEAttackingForm
    FormMainMode.pageul.FontSize = 24
    FormMainMode.bloodnumcom1.FontSize = 20
    FormMainMode.bloodnumcom2.FontSize = 10
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
   FormHint.Left = FormMainMode.Left + 1185
   FormHint.Top = FormMainMode.Top + 3030
   FormHint.Show 1
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
    '===============================
    For i = 1 To 角色人物對戰人數(1, 1)
        If Val(VBEPerson(1, i, 1, 3, 5)) = 1 Then
            FormMainMode.PEASusbi1(i).Left = 300
            FormMainMode.PEASusbi1(i).Top = 3220
            FormMainMode.PEASusbi2(i).Left = 960
            FormMainMode.PEASusbi2(i).Top = 3220
            FormMainMode.PEASusbi3(i).Left = 1820
            FormMainMode.PEASusbi3(i).Top = 3220
        Else
            FormMainMode.PEASusbi1(i).Left = 555
            FormMainMode.PEASusbi1(i).Top = 3240
            FormMainMode.PEASusbi2(i).Left = 1200
            FormMainMode.PEASusbi2(i).Top = 3240
            FormMainMode.PEASusbi3(i).Left = 1920
            FormMainMode.PEASusbi3(i).Top = 3240
        End If
    Next
    For i = 1 To 角色人物對戰人數(2, 1)
        If Val(VBEPerson(2, i, 1, 3, 5)) = 1 Then
            FormMainMode.PEAScardcompi1(i).Left = 230
            FormMainMode.PEAScardcompi1(i).Top = 3220
            FormMainMode.PEAScardcompi2(i).Left = 960
            FormMainMode.PEAScardcompi2(i).Top = 3220
            FormMainMode.PEAScardcompi3(i).Left = 1820
            FormMainMode.PEAScardcompi3(i).Top = 3220
        Else
            FormMainMode.PEAScardcompi1(i).Left = 480
            FormMainMode.PEAScardcompi1(i).Top = 3240
            FormMainMode.PEAScardcompi2(i).Left = 1200
            FormMainMode.PEAScardcompi2(i).Top = 3240
            FormMainMode.PEAScardcompi3(i).Left = 1920
            FormMainMode.PEAScardcompi3(i).Top = 3240
        End If
    Next
    '===========================戰鬥系統主表單讀入(測)
    執行階段系統類.執行階段系統遊戲初始總程序
    戰鬥系統類.遊戲角色卡片物件創立
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
Dim i As Integer
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
                    If Utils.GetExtName(mydir) = "uleci" Then
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
Dim i As Integer, j As Integer, ckl As Integer, mm As Integer, w As Integer, n As Integer '暫時變數
'------------
goidefus = 0
movecp = 2
turnpageonin = 0
trend暫時變數 = 0
FormMainMode.PEAFInterface.MessageClear
FormMainMode.PEAFInterface.BnOKStopListen
FormMainMode.PEAFInterface.BnOKVisable False
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
    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\stageblack.gif"
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
FormMainMode.PEAFpersoncardus(1).Left = 0
FormMainMode.PEAFpersoncardus(1).Visible = False
FormMainMode.cardus(1).Left = 0
FormMainMode.cardus(1).Top = 6240
FormMainMode.cardus(1).ZOrder
FormMainMode.cardus(1).Visible = True
FormMainMode.PEAFpersoncardcom(1).Left = 0
戰鬥系統類.PersonCardShowOnMode(1, 1) = True
戰鬥系統類.PersonCardShowOnMode(2, 1) = True
FormMainMode.PEAFpersoncardus(1).ShowOnMode = True
FormMainMode.PEAFpersoncardcom(1).ShowOnMode = True
'=======
For i = 2 To 3
   If 角色人物對戰人數(1, 1) >= i Then
       FormMainMode.PEAFpersoncardus(i).ShowOnMode = True
       FormMainMode.PEAFpersoncardus(i).Left = 2520 * (i - 1)
       FormMainMode.PEAFpersoncardus(i).Visible = True
       FormMainMode.cardus(i).Visible = False
       戰鬥系統類.PersonCardShowOnMode(1, i) = True
   Else
       FormMainMode.PEAFpersoncardus(i).Visible = False
       FormMainMode.cardus(i).CardMain_角色HP = 0
   End If
   If 角色人物對戰人數(2, 1) >= i Then
       戰鬥系統類.PersonCardShowOnMode(2, i) = False
       FormMainMode.PEAFpersoncardcom(i).ShowOnMode = False
       FormMainMode.PEAFpersoncardcom(i).Left = 2520 * (i - 1)
       FormMainMode.PEAFpersoncardcom(i).Visible = True
   Else
       FormMainMode.PEAFpersoncardcom(i).Visible = False
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
    FormMainMode.cardus(w).CardMain_角色HP = liveus(w)
    FormMainMode.cardus(w).CardMain_角色ATK = atkus(w)
    FormMainMode.cardus(w).CardMain_角色DEF = defus(w)
    liveusmax(w) = liveus(w)
    liveus41(w) = liveusmax(w) \ 3
    '=================
    戰鬥系統類.介面角色小卡資訊寫入 1, w
    '=================
    FormMainMode.cardus(w).CardMain_角色HPMAX = liveusmax(w)
    FormMainMode.cardus(w).CardMain_是否為新樣式資訊 = CBool(Val(VBEPerson(1, w, 1, 3, 5)) = 1)
    '=================
    FormMainMode.cardus(w).異常狀態全重設
    FormMainMode.cardus(w).CardBack全重設
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
    '=================
    戰鬥系統類.介面角色小卡資訊寫入 2, w
    '=================
    livecom41(w) = livecommax(w) \ 3
    FormMainMode.cardcom(w).CardMain_角色HP = livecom(w)
    FormMainMode.cardcom(w).CardMain_角色HPMAX = livecommax(w)
    FormMainMode.cardcom(w).CardMain_角色ATK = atkcom(w)
    FormMainMode.cardcom(w).CardMain_角色DEF = defcom(w)
    FormMainMode.cardcom(w).CardMain_是否為新樣式資訊 = CBool(Val(VBEPerson(2, w, 1, 3, 5)) = 1)
    '=================
    FormMainMode.cardcom(w).異常狀態全重設
    FormMainMode.cardcom(w).CardBack全重設
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
'==================寫入技能欄座標
atkinghelpxy(1, 1, 1) = 2520
atkinghelpxy(1, 2, 1) = 4730
atkinghelpxy(1, 3, 1) = 6930
atkinghelpxy(1, 4, 1) = 9140

atkinghelpxy(2, 1, 1) = 6840
atkinghelpxy(2, 2, 1) = 4560
atkinghelpxy(2, 3, 1) = 2280
atkinghelpxy(2, 4, 1) = 0
For i = 1 To 4
    atkinghelpxy(1, i, 2) = 3000
    atkinghelpxy(2, i, 2) = 960
Next
'===========================技能說明載入
戰鬥系統類.技能說明載入_使用者
戰鬥系統類.技能說明載入_電腦
FormMainMode.PEAFInterface.Passive_技能一方全重設 1
FormMainMode.PEAFInterface.Passive_技能一方全重設 2
For n = 5 To 8
    If VBEPerson(1, 1, 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.Passive_使用者_技能隱藏 n - 4
    Else
       FormMainMode.PEAFInterface.Passive_使用者_技能名稱 VBEPerson(1, 1, 3, n, 1), n - 4
       FormMainMode.PEAFInterface.Passive_使用者_技能顯示 n - 4
    End If
    '=============
    If VBEPerson(2, 1, 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.Passive_電腦_技能隱藏 n - 4
    Else
       FormMainMode.PEAFInterface.Passive_電腦_技能名稱 VBEPerson(2, 1, 3, n, 1), n - 4
       FormMainMode.PEAFInterface.Passive_電腦_技能顯示 n - 4
    End If
Next
FormMainMode.PEAFatkinghelpc.Visible = False
'=====================以下是技能欄顏色顯示
For i = 1 To 2
    For j = 1 To 4
        FormMainMode.PEAFInterface.ActiveSkillLight i, j, False
    Next
Next
'==================執行小人物立繪指定及距離指定
執行動作_距離變更 movecp, False
'================仿對戰模式設定
If Formsetting.chkpersonvsmode.Value = 1 Then
    For i = 2 To 3
        FormMainMode.PEAFpersoncardcom(i).ShowOnMode = False
        FormMainMode.cardcom(i).CardMain_角色HP = -99
        FormMainMode.cardcom(i).CardMain_角色ATK = -99
        FormMainMode.cardcom(i).CardMain_角色DEF = -99
        FormMainMode.cardcom(i).CardMain_角色圖片 = app_path & "gif\system\personunknown.jpg"
        FormMainMode.cardcom(i).CardBack全重設
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
For i = 1 To 2
    For j = 1 To UBound(人物異常狀態列表, 2)
        Set 人物異常狀態列表(i, j) = New Collection
    Next
Next
'==================
BattleTurn = 1
FormMainMode.PEAFInterface.turn = BattleTurn
FormMainMode.PEAFAnimateInterface.MusicPlayerObj = FormMainMode.cMusicPlayer(5)
FormMainMode.PEAFAnimateInterface.ImageMaskUse = Formsetting.chkImageMaskUse.Value
End Sub
Sub 自由戰鬥模式設定表單讀入程序()
Dim i, j As Integer
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
一般系統類.卡片人物資訊載入_搜尋檔案
'===============調整預設
If FormMainMode.personnameus(1).ListCount > 0 Then
    For i = 1 To 3
       FormMainMode.personnameus(i).ListIndex = 0
       FormMainMode.personnamecom(i).ListIndex = 0
    Next
End If
FormMainMode.opnpersonvs(2).Value = True
FormMainMode.cMusicPlayer(0).MusicPlay
FormMainMode.personreadifus.Visible = False
End Sub
Sub 遊戲初始讀入程序()
    Dim i As Integer
    '=====以下是背景音樂及SE初始設定
    For i = FormMainMode.cMusicPlayer.UBound + 1 To 11
        Load FormMainMode.cMusicPlayer(i)
    Next
    FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulbgm03.mp3"
    FormMainMode.cMusicPlayer(0).Volume = 50
    FormMainMode.cMusicPlayer(0).IsLoop = True
    一般系統類.音效初始設定
    For i = 1 To FormMainMode.cMusicPlayer.UBound
          FormMainMode.cMusicPlayer(i).Volume = 45
    Next
End Sub
Sub 主選單_PEStartForm顯示()
FormMainMode.PEStartForm.Left = 0
FormMainMode.PEStartForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
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
FormMainMode.Height = 10325
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
If Formsetting.chkautocontinuemode.Value = 1 Then
    FormMainMode.PEGFbnstart_Click
End If
End Sub
Sub 主選單_PEAttackingForm顯示()
FormMainMode.PEAttackingForm.Left = 0
FormMainMode.PEAttackingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
FormMainMode.PEAttackingForm.Visible = True
FormMainMode.PEAttackingForm.ZOrder
End Sub
Sub 主選單_PEAttackingStartForm顯示()
FormMainMode.PEAttackingStartForm.Left = 0
FormMainMode.PEAttackingStartForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
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
FormMainMode.Height = 10325
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
FormMainMode.upjpg_2.Top = -Val(FormMainMode.upjpg.Height)
For i = 1 To 3
   FormMainMode.PEAScardus(i).Top = -Val(FormMainMode.PEAScardus(i).Height)
   FormMainMode.PEAScardcom(i).Top = -Val(FormMainMode.PEAScardcom(i).Height)
Next
FormMainMode.大人物形像_使用者.大人物圖片 = VBEPerson(1, 1, 1, 5, 3)
FormMainMode.大人物形像_使用者.大人物影像反轉 = False
FormMainMode.大人物形像_使用者.Top = 8400 - FormMainMode.大人物形像_使用者.大人物圖片height
FormMainMode.大人物形像_使用者.Width = FormMainMode.大人物形像_使用者.大人物圖片width
FormMainMode.大人物形像_使用者.Left = -FormMainMode.大人物形像_使用者.大人物圖片width
FormMainMode.大人物形像_電腦.大人物圖片 = VBEPerson(2, 1, 1, 5, 3)
FormMainMode.大人物形像_電腦.大人物影像反轉 = True
FormMainMode.大人物形像_電腦.Top = 8400 - FormMainMode.大人物形像_電腦.大人物圖片height
FormMainMode.大人物形像_電腦.Width = FormMainMode.大人物形像_電腦.大人物圖片width
FormMainMode.大人物形像_電腦.Left = FormMainMode.ScaleWidth
st = 0
sq = 0
FormMainMode.start1.Enabled = True
End Sub
Sub 自由戰鬥模式設定表單基本設定程序()
Dim i As Integer '暫時變數
Formsetting.對戰地圖選擇.ListIndex = 0
Formsetting.BGM選擇.ListIndex = 0
Formsetting.comboeventcarrdus.ListIndex = 2
Formsetting.comboeventcarrdcom.ListIndex = 2
Formsetting.cbsimilarlevel.ListIndex = 4
For i = 1 To 18
    Formsetting.personus(i).ListIndex = 0
    Formsetting.personcom(i).ListIndex = 0
Next
Formsetting.lopnmusictext.Visible = False
Formsetting.lopnmapjpgtext.Visible = False
Formsetting.ckendturnnum.Text = 18
Formsetting.t1.Tab = 0
Formsetting.chkusenewai.Value = 1
Formsetting.chkusenewpage.Value = 1
Formsetting.chkusenewinterface.Value = 1
Formsetting.chkpersonvsmode.Value = 1
Formsetting.ckendturn.Value = 1
End Sub
Sub 清除戰鬥系統所有變數值()
Dim i As Integer, j As Integer '暫時變數
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
'Erase 人物異常狀態資料庫 '異常狀態資料(1.使用者/2.電腦,第x個異常狀態,1.狀態數值/2.狀態統計數(剩餘回合/累計)/3.狀態編號)
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
Erase 公用牌實體卡片分隔紀錄數 '戰鬥系統實體牌相關紀錄數(1.總共牌數/2.公牌牌數/3.使用者事件卡最底編號/4.電腦事件卡最底編號)
Erase 戰鬥擲骰介面人物立繪圖路徑紀錄數 '戰鬥系統擲骰介面雙方人物立繪圖路徑紀錄數(1.使用者方/2.電腦方)
Erase 人物實際狀態資料庫 '人物實際狀態資料
'===================
Erase 事件卡記錄暫時數 '事件卡使用紀錄暫時變數(0.(1)總共給予回合數,1.使用者/2.電腦,1.總共數值/2.目前處理數值/3.目前階段/4.事件卡牌編號/5.事件分類/6.是否啟動)
''===================
vbecommadtotplay = 0
ReDim vbecommadnum(1 To 7, vbecommadtotplay)
ReDim vbecommadstr(1 To 3, vbecommadtotplay)
Erase Vss_PersonAtkingOffNum
Erase Vss_AtkingInformationRecordStr
ReDim VBEStageNum(0) As Integer
For i = 1 To 2
    For j = 1 To UBound(人物異常狀態列表, 2)
        Set 人物異常狀態列表(i, j) = Nothing
    Next
Next
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
        FormMainMode.PEAFpersoncardcom(j).Left = 2520 * (j - 1)
        FormMainMode.PEAFpersoncardus(j).Left = 2520 * (j - 1)
    Next
    '================
    If i >= 2 Then
        Formchangeperson.card(i - 1).Visible = True
        Formchangeperson.bnok(i - 1).Visible = True
    End If
Next
End Sub
Sub 自由戰鬥模式設定表單各式設定讀入程序()
Dim i As Integer, ne As Integer, nd As Integer  '暫時變數
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
Dim k As Integer
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
'Select Case num
'    Case 1
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse06.mp3"
'    Case 2
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse09.mp3"
'    Case 3
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse08.mp3"
'    Case 4
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse29.mp3"
'    Case 5
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse13.mp3"
'    Case 6
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse12.mp3"
'    Case 7
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse11.mp3"
'    Case 8
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse10_f.mp3"
'    Case 9
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse23.mp3"
'    Case 10
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse22.mp3"
'    Case 11
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse01.mp3"
'End Select
FormMainMode.cMusicPlayer(num).MusicStop
FormMainMode.cMusicPlayer(num).MusicPlay
End Sub
Sub 音效初始設定()
FormMainMode.cMusicPlayer(1).Filepath = app_path & "mp3\ulse06.mp3"
FormMainMode.cMusicPlayer(2).Filepath = app_path & "mp3\ulse09.mp3"
FormMainMode.cMusicPlayer(3).Filepath = app_path & "mp3\ulse08.mp3"
FormMainMode.cMusicPlayer(4).Filepath = app_path & "mp3\ulse29.mp3"
FormMainMode.cMusicPlayer(5).Filepath = app_path & "mp3\ulse13.mp3"
FormMainMode.cMusicPlayer(6).Filepath = app_path & "mp3\ulse12.mp3"
FormMainMode.cMusicPlayer(7).Filepath = app_path & "mp3\ulse11.mp3"
FormMainMode.cMusicPlayer(8).Filepath = app_path & "mp3\ulse10_f.mp3"
FormMainMode.cMusicPlayer(9).Filepath = app_path & "mp3\ulse23.mp3"
FormMainMode.cMusicPlayer(10).Filepath = app_path & "mp3\ulse22.mp3"
FormMainMode.cMusicPlayer(11).Filepath = app_path & "mp3\ulse01.mp3"
End Sub
Sub 離開遊戲提示(Cancel As Integer, UnloadMode As Integer)
Dim YesNo As VbMsgBoxResult
If UnloadMode = 0 Then
   YesNo = MsgBox("確定離開遊戲?", 36, "UnlightVBE-系統提示")
   If YesNo = VbMsgBoxResult.vbYes Then
    End
   Else
    Cancel = 1
   End If
End If

End Sub
