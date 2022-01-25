# UnlightVBE-QS Origin

遊戲介紹：http://unlightvbe.blogspot.com/

開發語言：VB6、VBScript

開發環境：Visual Studio 6.0 SP6

本專案為主程式source code，若只想單純下載遊玩者請左轉到[UnlightVBE-QS (User Edition)](https://github.com/unlightvbe/unlightvbe_qs_user "UnlightVBE-QS (User Edition)")

___

在服用之前須知以下幾點：

1. 請以`ULVBE.vbg`為開啟專案的進入點
2. 請做好身處雜亂無章的環境中之心理準備
3. 某些source code檔案並未在專案中使用，此為正常現象，這些檔案有些是歷史文物，有些為開發過程中測試之用途。
4. 在專案中你會看到一些被註解起來的code，此為正常現象。
5. 因為本人我也很久沒碰了，所以......請加油。

#### 注意！本專案為Big5編碼，CRLF換行，請在push時將git的換行符號自動轉換的功能關閉
命令列輸入：```git config --global core.autocrlf false```
___

## Linux Wine 環境配置（試驗）  
#### PlayOnLinux 
1. install PlayOnLinux
2. 新增虛擬磁碟(32 bits installation)
3. 安裝套件: vbrun6, wsh57, quartz
4. 將UnlightVBE-QS遊戲資料夾複製到該虛擬磁碟上
5. 將ttf內所有字體複製至`windows/fonts`內
6. 將`COMCTL32.OCX`/`COMDLG32.OCX`/`Imagex.ocx`/`TABCTL32.OCX`/`msscript.ocx`覆蓋至`windows/system32`內
7. PlayOnLinux設定: 建立新的捷徑>`UnlightVBE-QS.exe`, 程式設定: 參數>```/wine```
8. 配置wine: 版本>`Windows XP`, 函式庫>`quartz`:原生, `winegstreamer`:停用
9. 執行，請耐心等待。
10. 若成功進入loading畫面卻閃退者可能是因為wine字體編碼問題，請Google: `wine big5`

## 遊戲簡介

![ULVBE_logo](http://3.bp.blogspot.com/-TyrMtORJqrE/UhzAREQ4twI/AAAAAAAAABQ/nUKTAy2q7e8/s1600/unlightvbelong.jpg "ULVBE logo")  
#### Unlight -The Visual Basic Edition-  

一個因突發奇想而以VB6語法寫成的一套單機小遊戲。  
本作仿製Unlight的角色戰鬥系統，您可以自由選擇角色、戰鬥的地圖、BGM、以及插入您想要的事件卡組合。~宗旨為創造自由、無限想像的戰鬥平台。~  

#### Type QS
UnlightVBE-QS為一「應用程式框架」，作為全開放式文件設計之開發路線，其遊戲角色與相關技能內容是以腳本形式去做外加的。  
VBE在經歷了2個正式版本的發布（Version α、Version ζ）後，雖然其遊戲系統完整性日益成熟，不過在人物角色的資訊的自由編輯及可攜性方面，仍然是相當的有限。在追求人物角色技能程式碼外部化，達到角色資訊完全獨立的期望之下，本版本（QS）也就此誕生。

## 素材授權
![CC-BY-ND 4.0](https://i.creativecommons.org/l/by-nd/4.0/88x31.png)  
本程式內使用之Unlight相關素材，皆授權自CPA。  
(CC BY-ND 4.0) CPA Co.,Ltd.
