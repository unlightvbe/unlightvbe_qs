Attribute VB_Name = "人物系統類"
Public totpersonnumber As Integer '現在目前處理第幾人暫時數
Public 總共人物名稱 As String '目前總共讀入人物名稱
Public 總共人物檔案名 As String '目前總共讀入人物檔案名
Public 選單使用者事件 As Boolean '選單類是否為使用者事件暫時數
Public 選單電腦事件 As Boolean '選單類是否為電腦事件暫時數
Public VBEPerson(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11) As String 'VBE人物統一資料記錄變數
Public VBEVSSAtkingStr(1 To 2, 1 To 3, 1 To 8, 1 To 2) As String 'VBE人物技能相關名稱紀錄(1.使用者/2.電腦,1~4主動技/5~8被動技,1.技能唯一識別碼/2.技能腳本檔案名稱)
Public VBEVSSBuffStr1() As String 'VBE人物異常狀態相關名稱紀錄(1~n異常狀態-技能唯一識別碼)
Public VBEVSSBuffStr2() As String 'VBE人物異常狀態相關名稱紀錄(1~n異常狀態-技能腳本檔案名稱)
Public VBEVSSActualStatusStr1() As String 'VBE人物實際狀態相關名稱紀錄(1~n人物實際狀態-技能唯一識別碼)
Public VBEVSSActualStatusStr2() As String 'VBE人物實際狀態相關名稱紀錄(1~n人物實際狀態-技能腳本檔案名稱)
Dim app_path As String  '路徑設定碼
Public VBETalkLevelStr(1 To 2) As String 'VBE人物對話適用人物等級類別字串(1.使用者方/2.電腦方)
Sub 卡片人物資訊讀入_初階段(ByVal filename As String)
Dim textlinea As String '讀入文件時一行處理暫時變數
Dim 文件字串() As String
Dim textcheck As Boolean '判斷文件檢查碼正確性變數
'MsgBox filename
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   文件字串 = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "Q4VX435S" Then
'           MsgBox "讀入檔案時發生錯誤!"
           卡片人物資訊檔案讀取失敗紀錄串 = 卡片人物資訊檔案讀取失敗紀錄串 & "=" & filename
           Exit Do
       Else
           textcheck = True
           加入總共人物檔案名字串 filename
       End If
   End If
   If textlinea = "" Then
       GoTo 略過本行字串
   End If
   Select Case 文件字串(0)
       Case "MenuList"
           
       Case "MenuName"
           加入總共人物名稱字串 文件字串(1)
           更新人物清單_使用者方_初設
           更新人物清單_電腦方_初設
       Case "EndFirst"
           Exit Do
   End Select
略過本行字串:
Loop
Close
End Sub
Sub 卡片人物資訊讀入_二階段_使用者(ByVal personname As String, ByVal Index As Integer)
Dim textlinea As String '讀入文件時一行處理暫時變數
Dim 文件字串() As String
Dim textcheck As Boolean '判斷文件檢查碼正確性變數
Dim filename As String '目標人物檔案名暫時數
Dim at() As String
at = Split(總共人物名稱, "=")
Dim aw() As String
aw = Split(總共人物檔案名, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
FormMainMode.personlevelus(Index).Clear
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   文件字串 = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "Q4VX435S" Then
           MsgBox "讀入檔案時發生錯誤!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo 略過本行字串
   End If
   Select Case 文件字串(0)
       Case "MenuList"
           For i = 1 To UBound(文件字串)
               FormMainMode.personlevelus(Index).AddItem 文件字串(i)
           Next
       Case "EndFirst"
           Exit Do
   End Select
略過本行字串:
Loop
Close
End Sub
Sub 卡片人物資訊讀入_二階段_電腦(ByVal personname As String, ByVal Index As Integer)
Dim textlinea As String '讀入文件時一行處理暫時變數
Dim 文件字串() As String
Dim textcheck As Boolean '判斷文件檢查碼正確性變數
Dim filename As String '目標人物檔案名暫時數
Dim at() As String
at = Split(總共人物名稱, "=")
Dim aw() As String
aw = Split(總共人物檔案名, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
FormMainMode.personlevelcom(Index).Clear
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   文件字串 = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "Q4VX435S" Then
           MsgBox "讀入檔案時發生錯誤!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo 略過本行字串
   End If
   Select Case 文件字串(0)
       Case "MenuList"
           For i = 1 To UBound(文件字串)
               FormMainMode.personlevelcom(Index).AddItem 文件字串(i)
           Next
       Case "EndFirst"
           Exit Do
   End Select
略過本行字串:
Loop
Close
End Sub
Sub 卡片人物資訊讀入_三階段(ByVal personname As String, ByVal personlevel As String, ByVal Index As Integer, ByVal uscom As Integer)
Dim textlinea As String '讀入文件時一行處理暫時變數
Dim 文件字串() As String
Dim textcheck As Boolean '判斷文件檢查碼正確性變數
Dim filename As String '目標人物檔案名暫時數
Dim 略過資訊 As Boolean '是否略過目前區段暫時數
Dim at() As String
at = Split(總共人物名稱, "=")
Dim aw() As String
aw = Split(總共人物檔案名, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
'======================
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   文件字串 = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "Q4VX435S" Then
           MsgBox "讀入檔案時發生錯誤!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo 略過本行字串
   End If
   If 略過資訊 = False Then
        Select Case 文件字串(0)
            Case "StartPerson"
               If 文件字串(1) <> personlevel Or 文件字串(2) <> personname Then
                   略過資訊 = True
               End If
            Case "cardjpg"
               VBEPerson(uscom, Index, 1, 5, 5) = app_path & "gif\" & 文件字串(1)
            Case "personhp"
               VBEPerson(uscom, Index, 1, 3, 1) = 文件字串(1)
            Case "personatk"
               VBEPerson(uscom, Index, 1, 3, 2) = 文件字串(1)
            Case "persondef"
               VBEPerson(uscom, Index, 1, 3, 3) = 文件字串(1)
            Case "personname"
               VBEPerson(uscom, Index, 1, 1, 1) = 文件字串(1)
            Case "personengname"
               VBEPerson(uscom, Index, 1, 1, 2) = 文件字串(1)
            Case "personpname"
               VBEPerson(uscom, Index, 1, 1, 3) = 文件字串(1)
            Case "personlevel1"
               VBEPerson(uscom, Index, 1, 2, 1) = 文件字串(1)
            Case "personlevel2"
               VBEPerson(uscom, Index, 1, 2, 2) = 文件字串(1)
            Case "cardid"
               VBEPerson(uscom, Index, 1, 4, 1) = 文件字串(1)
            Case "persontg"
               VBEPerson(uscom, Index, 1, 3, 4) = 文件字串(1)
            Case "personbig"
               VBEPerson(uscom, Index, 1, 5, 3) = app_path & "gif\" & 文件字串(1)
            Case "personmini"
               VBEPerson(uscom, Index, 1, 5, 1) = app_path & "gif\" & 文件字串(1)
            Case "personf"
               VBEPerson(uscom, Index, 1, 5, 4) = app_path & "gif\" & 文件字串(1)
            Case "personsmalldown"
               VBEPerson(uscom, Index, 1, 5, 2) = app_path & "gif\" & 文件字串(1)
'            Case "personfleftall"
'               VBEPerson(uscom, Index, 2, 4, 1) = 文件字串(1)
            Case "atkingfontck"
               VBEPerson(uscom, Index, 2, 3, 5) = 文件字串(1)
            Case "bight"
               VBEPerson(uscom, Index, 2, 2, 1) = 文件字串(1)
            Case "bigtop"
               VBEPerson(uscom, Index, 2, 2, 3) = 文件字串(1)
            Case "bigwh"
               VBEPerson(uscom, Index, 2, 2, 2) = 文件字串(1)
            Case "minileft1"
               VBEPerson(uscom, Index, 2, 1, 1) = 文件字串(1)
            Case "minileft2"
               VBEPerson(uscom, Index, 2, 1, 2) = 文件字串(1)
            Case "minileft3"
               VBEPerson(uscom, Index, 2, 1, 3) = 文件字串(1)
            Case "minitop"
               VBEPerson(uscom, Index, 2, 1, 4) = 文件字串(1)
            Case "atkingjpgleftallzero"
               VBEPerson(uscom, Index, 2, 2, 5) = 文件字串(1)
            Case "bigleftall"
               VBEPerson(uscom, Index, 2, 2, 4) = 文件字串(1)
            '======================================
            Case "smalldownleftus"
               If uscom = 1 Then
                   VBEPerson(1, Index, 2, 1, 5) = 文件字串(1)
               End If
            Case "smalldowntopus"
               If uscom = 1 Then
                   VBEPerson(1, Index, 2, 1, 6) = 文件字串(1)
               End If
            Case "smalldownleftcom"
               If uscom = 2 Then
                   VBEPerson(2, Index, 2, 1, 5) = 文件字串(1)
               End If
            Case "smalldowntopcom"
               If uscom = 2 Then
                   VBEPerson(2, Index, 2, 1, 6) = 文件字串(1)
               End If
            '=======================================
            Case "atkingfont1"
               VBEPerson(uscom, Index, 2, 3, 1) = 文件字串(1)
            Case "atkingfont2"
               VBEPerson(uscom, Index, 2, 3, 2) = 文件字串(1)
            Case "atkingfont3"
               VBEPerson(uscom, Index, 2, 3, 3) = 文件字串(1)
            Case "atkingfont4"
               VBEPerson(uscom, Index, 2, 3, 4) = 文件字串(1)
            Case "atkingcfont(1)"
               VBEPerson(uscom, Index, 3, 1, 6) = 文件字串(1)
            Case "atkingcfont(2)"
               VBEPerson(uscom, Index, 3, 2, 6) = 文件字串(1)
            Case "atkingcfont(3)"
               VBEPerson(uscom, Index, 3, 3, 6) = 文件字串(1)
            Case "atkingcfont(4)"
               VBEPerson(uscom, Index, 3, 4, 6) = 文件字串(1)
            Case "atkingdfont(1)"
               VBEPerson(uscom, Index, 3, 1, 7) = 文件字串(1)
            Case "atkingdfont(2)"
               VBEPerson(uscom, Index, 3, 2, 7) = 文件字串(1)
            Case "atkingdfont(3)"
               VBEPerson(uscom, Index, 3, 3, 7) = 文件字串(1)
            Case "atkingdfont(4)"
               VBEPerson(uscom, Index, 3, 4, 7) = 文件字串(1)
            Case "atkingname(1)"
               VBEPerson(uscom, Index, 3, 1, 11) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 1, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 1, 2) = 文件字串(2)
            Case "atkingname(2)"
               VBEPerson(uscom, Index, 3, 2, 11) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 2, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 2, 2) = 文件字串(2)
            Case "atkingname(3)"
               VBEPerson(uscom, Index, 3, 3, 11) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 3, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 3, 2) = 文件字串(2)
            Case "atkingname(4)"
               VBEPerson(uscom, Index, 3, 4, 11) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 4, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 4, 2) = 文件字串(2)
            Case "atkingname(5)"
               VBEPerson(uscom, Index, 3, 5, 3) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 5, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 5, 2) = 文件字串(2)
            Case "atkingname(6)"
               VBEPerson(uscom, Index, 3, 6, 3) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 6, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 6, 2) = 文件字串(2)
            Case "atkingname(7)"
               VBEPerson(uscom, Index, 3, 7, 3) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 7, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 7, 2) = 文件字串(2)
            Case "atkingname(8)"
               VBEPerson(uscom, Index, 3, 8, 3) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 8, 1) = 文件字串(1)
               VBEVSSAtkingStr(uscom, Index, 8, 2) = 文件字串(2)
'===========================================================
        End Select
   End If
   If 文件字串(0) = "EndPerson" Then
       略過資訊 = False
   End If
略過本行字串:
Loop
Close
End Sub

Sub 卡片人物資訊讀入_四階段(ByVal personname As String, ByVal Index As Integer, ByVal uscom As Integer)
Dim textlinea As String '讀入文件時一行處理暫時變數
Dim 文件字串() As String
Dim textcheck As Boolean '判斷文件檢查碼正確性變數
Dim filename As String '目標人物檔案名暫時數
Dim 略過資訊 As Boolean '是否略過目前區段暫時數
Dim persontalka As Integer '資料寫入暫時變數
Dim at() As String
at = Split(總共人物名稱, "=")
Dim aw() As String
aw = Split(總共人物檔案名, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
'======================
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   文件字串 = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "Q4VX435S" Then
           MsgBox "讀入檔案時發生錯誤!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
        GoTo 略過本行字串
   End If
   If 略過資訊 = False Then
        If Left(文件字串(0), 4) = "Talk" Then
           If 文件字串(1) = "" Then
                GoTo 略過本行字串
           End If
        End If
        '=====================
        Select Case 文件字串(0)
           Case "StartTalk"
               略過資訊 = True
               '========================
               If 文件字串(1) = personname Then
                   If UBound(文件字串) >= 2 Then
                       For i = 2 To UBound(文件字串)
                           If VBEPerson(uscom, Index, 1, 2, 1) = 文件字串(i) Then
                               略過資訊 = False
                               For k = 2 To UBound(文件字串)
                                    VBETalkLevelStr(uscom) = VBETalkLevelStr(uscom) & "=" & 文件字串(k)
                                Next
                               Exit For
                           End If
                       Next
                   End If
               End If
            Case "TalkA1"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA2"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA3"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA4"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA5"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA6"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA7"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA8"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA9"
               persontalka = Right(文件字串(0), 1)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA10"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA11"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA12"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA13"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA14"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA15"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA16"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA17"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA18"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA19"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkA20"
               persontalka = Right(文件字串(0), 2)
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
               VBEPerson(uscom, Index, 4, persontalka, 2) = 文件字串(2)
            Case "TalkB1"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB2"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB3"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB4"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB5"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB6"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB7"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB8"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB9"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
            Case "TalkB10"
               persontalka = Val(Right(文件字串(0), 1)) + 20
               VBEPerson(uscom, Index, 4, persontalka, 1) = 文件字串(1)
        End Select
   End If
   If 文件字串(0) = "EndTalk" Then
       略過資訊 = False
   End If
略過本行字串:
Loop
Close
End Sub

Sub 加入總共人物名稱字串(ByVal name As String)
總共人物名稱 = 總共人物名稱 & "=" & name
End Sub
Sub 加入總共人物檔案名字串(ByVal name As String)
總共人物檔案名 = 總共人物檔案名 & "=" & name
End Sub
Sub 更新人物清單_使用者方_初設()
Dim at() As String
at = Split(總共人物名稱, "=")
For i = 1 To 3
   FormMainMode.personnameus(i).Clear
   FormMainMode.personnameus(i).AddItem "《隨機》"
   For j = 1 To UBound(at)
       FormMainMode.personnameus(i).AddItem at(j)
   Next
Next
End Sub
Sub 更新人物清單_電腦方_初設()
Dim at() As String
at = Split(總共人物名稱, "=")
For i = 1 To 3
   FormMainMode.personnamecom(i).Clear
   FormMainMode.personnamecom(i).AddItem "《隨機》"
   For j = 1 To UBound(at)
       FormMainMode.personnamecom(i).AddItem at(j)
   Next
Next
End Sub
Sub 更新人物清單_使用者方_變更(ByVal 現在所在數 As Integer)
Dim at() As String
at = Split(總共人物名稱, "=")
Dim ag(1 To 3) As String  '紀錄目前選項暫時數
Dim ap, au As Integer '暫時變數
For i = 1 To 3
   ag(i) = FormMainMode.personnameus(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnameus(i).Clear
   FormMainMode.personnameus(i).AddItem "《隨機》"
   For j = 1 To UBound(at)
       FormMainMode.personnameus(i).AddItem at(j)
   Next
Next
 '===========================================
選單使用者事件 = False
'formmainmode.personnameus(現在所在數).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnameus(p).ListCount - 1
           If FormMainMode.personnameus(p).List(q) = ag(p) Then
               FormMainMode.personnameus(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnameus(p).ListIndex = -1
   End If
Next
選單使用者事件 = True
'========================================
For i = 1 To 3
   ap = FormMainMode.personnameus(i).ListCount - 1
   au = 0
        Do Until au > ap
            If FormMainMode.personnameus(i).List(au) <> "《隨機》" Then
                Select Case i
                    Case 1
                       If FormMainMode.personnameus(2).Text = FormMainMode.personnameus(i).List(au) Or FormMainMode.personnameus(3).Text = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If FormMainMode.personnameus(1).Text = FormMainMode.personnameus(i).List(au) Or FormMainMode.personnameus(3).Text = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If FormMainMode.personnameus(2).Text = FormMainMode.personnameus(i).List(au) Or FormMainMode.personnameus(1).Text = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
            End If
            au = au + 1
        Loop
Next
'===========檢查選單是否只有「隨機」一項
For i = 1 To 3
   If FormMainMode.personnameus(i).ListCount = 1 Then
       FormMainMode.personnameus(i).Clear
   End If
Next
選單使用者事件 = False
'formmainmode.personnameus(現在所在數).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnameus(i).ListCount - 1
           If FormMainMode.personnameus(i).List(k) = ag(i) Then
               FormMainMode.personnameus(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnameus(i).ListIndex = -1
   End If
Next
選單使用者事件 = True
End Sub
Sub 更新人物清單_使用者方_變更_開始隨機(ByVal 現在所在數 As Integer, ByVal name1 As String, ByVal name2 As String, ByVal name3 As String)
Dim at() As String
at = Split(總共人物名稱, "=")
Dim ag(1 To 3) As String  '紀錄目前選項暫時數
Dim ap, au As Integer '暫時變數
For i = 1 To 3
   ag(i) = FormMainMode.personnameus(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnameus(i).Clear
   FormMainMode.personnameus(i).AddItem "《隨機》"
   For j = 1 To UBound(at)
       FormMainMode.personnameus(i).AddItem at(j)
   Next
Next
 '===========================================
選單使用者事件 = False
'formmainmode.personnameus(現在所在數).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnameus(p).ListCount - 1
           If FormMainMode.personnameus(p).List(q) = ag(p) Then
               FormMainMode.personnameus(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnameus(p).ListIndex = -1
   End If
Next
'========================================
For i = 1 To 3
   ap = FormMainMode.personnameus(i).ListCount - 1
   au = 0
        Do Until au > ap
'            If formmainmode.personnameus(i).List(au) <> "《隨機》" Then
                Select Case i
                    Case 1
                       If name2 = FormMainMode.personnameus(i).List(au) Or name3 = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If name1 = FormMainMode.personnameus(i).List(au) Or name3 = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If name2 = FormMainMode.personnameus(i).List(au) Or name1 = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
'            End If
            au = au + 1
        Loop
Next
'===========檢查選單是否只有「隨機」一項
For i = 1 To 3
   If FormMainMode.personnameus(i).ListCount = 1 Then
       FormMainMode.personnameus(i).Clear
   End If
Next
選單使用者事件 = False
'formmainmode.personnameus(現在所在數).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnameus(i).ListCount - 1
           If FormMainMode.personnameus(i).List(k) = ag(i) Then
               FormMainMode.personnameus(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnameus(i).ListIndex = -1
   End If
Next
End Sub

Sub 更新人物清單_電腦方_變更(ByVal 現在所在數 As Integer)
Dim at() As String
at = Split(總共人物名稱, "=")
Dim ag(1 To 3) As String  '紀錄目前選項暫時數
Dim ap, au As Integer '暫時變數
For i = 1 To 3
   ag(i) = FormMainMode.personnamecom(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnamecom(i).Clear
   FormMainMode.personnamecom(i).AddItem "《隨機》"
   For j = 1 To UBound(at)
       FormMainMode.personnamecom(i).AddItem at(j)
   Next
Next
 '===========================================
選單電腦事件 = False
'formmainmode.personnamecom(現在所在數).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnamecom(p).ListCount - 1
           If FormMainMode.personnamecom(p).List(q) = ag(p) Then
               FormMainMode.personnamecom(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnamecom(p).ListIndex = -1
   End If
Next
選單電腦事件 = True
'========================================
For i = 1 To 3
   ap = FormMainMode.personnamecom(i).ListCount - 1
   au = 0
        Do Until au > ap
            If FormMainMode.personnamecom(i).List(au) <> "《隨機》" Then
                Select Case i
                    Case 1
                       If FormMainMode.personnamecom(2).Text = FormMainMode.personnamecom(i).List(au) Or FormMainMode.personnamecom(3).Text = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If FormMainMode.personnamecom(1).Text = FormMainMode.personnamecom(i).List(au) Or FormMainMode.personnamecom(3).Text = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If FormMainMode.personnamecom(2).Text = FormMainMode.personnamecom(i).List(au) Or FormMainMode.personnamecom(1).Text = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
            End If
            au = au + 1
        Loop
Next
'===========檢查選單是否只有「隨機」一項
For i = 1 To 3
   If FormMainMode.personnamecom(i).ListCount = 1 Then
       FormMainMode.personnamecom(i).Clear
   End If
Next
選單電腦事件 = False
'formmainmode.personnamecom(現在所在數).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnamecom(i).ListCount - 1
           If FormMainMode.personnamecom(i).List(k) = ag(i) Then
               FormMainMode.personnamecom(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnamecom(i).ListIndex = -1
   End If
Next
選單電腦事件 = True
End Sub
Sub 更新人物清單_電腦方_變更_開始隨機(ByVal 現在所在數 As Integer, ByVal name1 As String, ByVal name2 As String, ByVal name3 As String)
Dim at() As String
at = Split(總共人物名稱, "=")
Dim ag(1 To 3) As String  '紀錄目前選項暫時數
Dim ap, au As Integer '暫時變數
For i = 1 To 3
   ag(i) = FormMainMode.personnamecom(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnamecom(i).Clear
   FormMainMode.personnamecom(i).AddItem "《隨機》"
   For j = 1 To UBound(at)
       FormMainMode.personnamecom(i).AddItem at(j)
   Next
Next
 '===========================================
選單電腦事件 = False
'formmainmode.personnamecom(現在所在數).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnamecom(p).ListCount - 1
           If FormMainMode.personnamecom(p).List(q) = ag(p) Then
               FormMainMode.personnamecom(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnamecom(p).ListIndex = -1
   End If
Next
'========================================
For i = 1 To 3
   ap = FormMainMode.personnamecom(i).ListCount - 1
   au = 0
        Do Until au > ap
'            If formmainmode.personnamecom(i).List(au) <> "《隨機》" Then
                Select Case i
                    Case 1
                       If name2 = FormMainMode.personnamecom(i).List(au) Or name3 = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If name1 = FormMainMode.personnamecom(i).List(au) Or name3 = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If name2 = FormMainMode.personnamecom(i).List(au) Or name1 = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
'            End If
            au = au + 1
        Loop
Next
'===========檢查選單是否只有「隨機」一項
For i = 1 To 3
   If FormMainMode.personnamecom(i).ListCount = 1 Then
       FormMainMode.personnamecom(i).Clear
   End If
Next
選單電腦事件 = False
'formmainmode.personnamecom(現在所在數).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnamecom(i).ListCount - 1
           If FormMainMode.personnamecom(i).List(k) = ag(i) Then
               FormMainMode.personnamecom(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnamecom(i).ListIndex = -1
   End If
Next
End Sub
Sub 重設人物圖片(ByVal uscom As Integer, ByVal Index As Integer)
Select Case uscom
   Case 1
   
   Case 2
   
End Select
End Sub
Sub 卡片人物資訊顯示_使用者(ByVal Index As Integer)
FormMainMode.PEGFusbi1(Index).Caption = VBEPerson(1, Index, 1, 3, 1)
FormMainMode.PEGFusbi2(Index).Caption = VBEPerson(1, Index, 1, 3, 2)
FormMainMode.PEGFusbi3(Index).Caption = VBEPerson(1, Index, 1, 3, 3)
FormMainMode.PEGFcardus(Index).Picture = LoadPicture(VBEPerson(1, Index, 1, 5, 5))
End Sub
Sub 卡片人物資訊顯示_電腦(ByVal Index As Integer)
FormMainMode.PEGFcardcompi1(Index).Caption = VBEPerson(2, Index, 1, 3, 1)
FormMainMode.PEGFcardcompi2(Index).Caption = VBEPerson(2, Index, 1, 3, 2)
FormMainMode.PEGFcardcompi3(Index).Caption = VBEPerson(2, Index, 1, 3, 3)
FormMainMode.PEGFcardcom(Index).Picture = LoadPicture(VBEPerson(2, Index, 1, 5, 5))
End Sub
Sub 角色隨機_使用者(ByVal Index As Integer)
Dim i As Integer
For i = 1 To 4
    For j = 1 To 30
        For k = 1 To 10
           VBEPerson(1, Index, i, j, k) = ""
        Next
    Next
Next
'==============
VBEPerson(1, Index, 1, 5, 5) = App.Path & "\gif\system\personunknown.jpg"
VBEPerson(1, Index, 1, 3, 1) = "?"
VBEPerson(1, Index, 1, 3, 2) = "?"
VBEPerson(1, Index, 1, 3, 3) = "?"
VBEPerson(1, Index, 1, 1, 1) = "?"
VBEPerson(1, Index, 1, 1, 2) = "?"
VBEPerson(1, Index, 1, 1, 3) = "?"
VBEPerson(1, Index, 1, 2, 1) = "?"
VBEPerson(1, Index, 1, 2, 2) = "?"
VBEPerson(1, Index, 1, 4, 1) = "??????"
VBEPerson(1, Index, 2, 3, 5) = 1
VBEPerson(1, Index, 1, 3, 4) = "000000"
End Sub
Sub 角色隨機_電腦(ByVal Index As Integer)
Dim i As Integer
For i = 1 To 4
    For j = 1 To 30
        For k = 1 To 10
           VBEPerson(2, Index, i, j, k) = ""
        Next
    Next
Next
'==============
VBEPerson(2, Index, 1, 5, 5) = App.Path & "\gif\system\personunknown.jpg"
VBEPerson(2, Index, 1, 3, 1) = "?"
VBEPerson(2, Index, 1, 3, 2) = "?"
VBEPerson(2, Index, 1, 3, 3) = "?"
VBEPerson(2, Index, 1, 1, 1) = "?"
VBEPerson(2, Index, 1, 1, 2) = "?"
VBEPerson(2, Index, 1, 1, 3) = "?"
VBEPerson(2, Index, 1, 2, 1) = "?"
VBEPerson(2, Index, 1, 2, 2) = "?"
VBEPerson(2, Index, 1, 4, 1) = "??????"
VBEPerson(2, Index, 2, 3, 5) = 1
VBEPerson(2, Index, 1, 4, 3) = "?.?.?"
VBEPerson(2, Index, 1, 3, 4) = "000000"
End Sub
Function 人物對話選擇() As String
Dim personcomname As String '電腦方人物名稱暫時紀錄變數
Dim personcomlevel() As String '電腦方人物等級暫時紀錄變數
Dim talkname() As String  '每句對話人物記錄分別變數
Dim persontalkname As String  '每句對話人物記錄總變數
Dim persontalkrec As String '總共可選擇指定對話紀錄編號串
Dim persontalkrecnum As Integer '總共可選擇指定對話紀錄數
Dim at() As String '選擇對話暫時變數
Dim m As Integer '暫時變數
Dim atbo(1 To 10) As Boolean '人物本身對話空白標記記錄數
personcomname = VBEPerson(2, 1, 1, 1, 1)
personcomlevel = Split(VBETalkLevelStr(1), "=")

For i = 1 To 20
'   persontalkname = formsettingpersonus.persontalking2(i).Text
   persontalkname = VBEPerson(1, 1, 4, i, 2)
   talkname = Split(persontalkname, "&")
   For k = 0 To UBound(talkname)
       If talkname(k) = personcomname Then
           For p = 0 To UBound(personcomlevel)
               If VBEPerson(2, 1, 1, 2, 1) = personcomlevel(p) Then
                    persontalkrec = persontalkrec & i & "="
                    persontalkrecnum = persontalkrecnum + 1
                    k = UBound(talkname) '象徵ExitFor
                    p = UBound(personcomlevel) '象徵ExitFor
               End If
           Next
       End If
    Next
Next

If persontalkrecnum >= 1 Then
    m = Int(Rnd() * persontalkrecnum) + 1
    at = Split(persontalkrec, "=")
'    人物對話選擇 = formsettingpersonus.persontalking1(at(m - 1)).Text
    人物對話選擇 = VBEPerson(1, 1, 4, at(m - 1), 1)
Else
    Do
        Randomize
        m = Int(Rnd() * 10) + 1
        If atbo(m) = False Then
            人物對話選擇 = VBEPerson(1, 1, 4, m + 20, 1)
            atbo(m) = True
        End If
        If 人物對話選擇 <> "" Then
            Exit Do
        ElseIf atbo(1) = True And atbo(2) = True And atbo(3) = True And atbo(4) = True And atbo(5) = True _
        And atbo(6) = True And atbo(7) = True And atbo(8) = True And atbo(9) = True And atbo(10) = True Then
            人物對話選擇 = ""
            Exit Do
        Else
            atbo(m) = True
        End If
    Loop
End If
End Function
Sub 清除角色人物資訊變數(ByVal uscom As Integer, ByVal num As Integer)
For i = 1 To UBound(VBEPerson, 3)
    For j = 1 To UBound(VBEPerson, 4)
        For k = 1 To UBound(VBEPerson, 5)
            VBEPerson(uscom, num, i, j, k) = ""
        Next
    Next
Next
For i = 1 To UBound(VBEVSSAtkingStr, 3)
    For j = 1 To UBound(VBEVSSAtkingStr, 4)
         VBEVSSAtkingStr(uscom, num, i, j) = ""
    Next
Next
VBETalkLevelStr(uscom) = ""
End Sub
