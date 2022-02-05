VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWineobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private oVBEPersonVS()  As Variant  'VBE人物統一變數-VS版
Private oatkingpagetotVS() As Variant  '每階段出牌種類及數值統計資料-VS版
Private oVBEPersonBuffVSF() As Variant '異常狀態資料-VS-F版
Private oVBEPersonBuffVSS() As Variant '異常狀態資料-VS-S版
Private oAtkingckVSS() As Variant  '技能資訊一覽-S版(技能啟動碼)
Private oAtkingckVSF() As Variant '技能資訊一覽-F版(備註字串)
Private oVBEAtkingVSF() As Variant 'VBE>VS給予變數統一資料-F版
Private oVBEAtkingVSS() As Variant 'VBE>VS給予變數統一資料-S版
Private oVBEPageCardNumVS() As Variant '公用牌資料-VS版
Private oVBEVSBuffNum() As Variant '異常狀態專用-異常狀態之2個數值-VS版
Private oVBEVSStageNum() As Variant '執行階段系統-執行階段多用途紀錄變數-VS版
Private oVBEActualStatusVS() As Variant '人物實際狀態資料-VS版
Public oPersonType As Variant  '該角色目前場上狀態
Public oNs As Variant '執行階段號

Public Sub AddInformation(ByVal str As Variant, ByRef Varray() As Variant)
Select Case str
    Case "VBEAtkingVSF"
        oVBEAtkingVSF = Varray
    Case "VBEAtkingVSS"
        oVBEAtkingVSS = Varray
    Case "AtkingPagetotVS"
        oatkingpagetotVS = Varray
    Case "VBEPersonVS"
        oVBEPersonVS = Varray
    Case "VBEPageCardNumVS"
        oVBEPageCardNumVS = Varray
    Case "AtkingckVSS"
        oAtkingckVSS = Varray
    Case "AtkingckVSF"
        oAtkingckVSF = Varray
    Case "VBEPersonBuffVSF"
        oVBEPersonBuffVSF = Varray
    Case "VBEPersonBuffVSS"
        oVBEPersonBuffVSS = Varray
    Case "VBEActualStatusVS"
        oVBEActualStatusVS = Varray
    Case "VBEVSBuffNum"
        oVBEVSBuffNum = Varray
    Case "VBEVSStageNum"
        oVBEVSStageNum = Varray
End Select
End Sub
Public Function GetArray(ByVal str As Variant)
Select Case str
    Case "VBEAtkingVSF"
        GetArray = oVBEAtkingVSF
    Case "VBEAtkingVSS"
        GetArray = oVBEAtkingVSS
    Case "AtkingPagetotVS"
        GetArray = oatkingpagetotVS
    Case "VBEPersonVS"
        GetArray = oVBEPersonVS
    Case "VBEPageCardNumVS"
        GetArray = oVBEPageCardNumVS
    Case "AtkingckVSS"
        GetArray = oAtkingckVSS
    Case "AtkingckVSF"
        GetArray = oAtkingckVSF
    Case "VBEPersonBuffVSF"
        GetArray = oVBEPersonBuffVSF
    Case "VBEPersonBuffVSS"
        GetArray = oVBEPersonBuffVSS
    Case "VBEActualStatusVS"
        GetArray = oVBEActualStatusVS
    Case "VBEVSBuffNum"
        GetArray = oVBEVSBuffNum
    Case "VBEVSStageNum"
        GetArray = oVBEVSStageNum
End Select
End Function
