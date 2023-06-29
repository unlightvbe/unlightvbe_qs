VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Text As String
Private eventActiveAIScoreFlag As Boolean    'EventActiveAIScore指令啟動註記
Private eventActiveAIScoretmpMain(1 To 2) As Integer    'EventActiveAIScore主指令暫存
Private eventActiveAIScoretmpCardNum() As Integer    'EventActiveAIScore個別推薦牌編號暫存

Public Function Output() As String
    If eventActiveAIScoreFlag = True Then
        Call eventActiveAIScoreOutput

        eventActiveAIScoreFlag = False
        Erase eventActiveAIScoretmpMain
        Erase eventActiveAIScoretmpCardNum
    End If

    Output = Me.Text
    Me.Text = ""
End Function
Public Sub AtkingLineLight(ByVal n1 As Integer)
    Me.Text = Me.Text & "AtkingLineLight" & "#" & CStr(n1) & "="
End Sub
Public Sub AtkingTurnOnOff(ByVal n1 As Integer)
    Me.Text = Me.Text & "AtkingTurnOnOff" & "#" & CStr(n1) & "="
End Sub
Public Sub AtkingLineLightAnother(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "AtkingLineLightAnother" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub AtkingTurnOnOffAnother(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "AtkingTurnOnOffAnother" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub AtkingStartPlay(ByVal n1 As String, ByVal n2 As String)
    Me.Text = Me.Text & "AtkingStartPlay" & "#" & n1 & "," & n2 & "="
End Sub
Public Sub AtkingStartPlayAnimate(ByVal n1 As String, Optional ByVal n2 As Integer, Optional ByVal n3 As Integer)
    Me.Text = Me.Text & "AtkingStartPlayAnimate" & "#" & n1
    If n2 > 0 Then Me.Text = Me.Text & "," & CStr(n2)
    If n3 > 0 Then Me.Text = Me.Text & "," & CStr(n3)
    Me.Text = Me.Text & "="
End Sub
Public Sub BattleTurnControl(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "BattleTurnControl" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub BattleSendMessage(ByVal n1 As String)
    Me.Text = Me.Text & "BattleSendMessage" & "#" & n1 & "="
End Sub
Public Sub BattleStartDice()
    Me.Text = Me.Text & "BattleStartDice" & "#0" & "="
End Sub
Public Sub AtkingInformationRecord(ByVal n1 As String)
    Me.Text = Me.Text & "AtkingInformationRecord" & "#" & n1 & "="
End Sub
Public Sub EventPlayerAllActionOff(ByVal n1 As Integer)
    Me.Text = Me.Text & "EventPlayerAllActionOff" & "#" & CStr(n1) & "="
End Sub
Public Sub EventActiveAIScore(ByVal n1 As Integer, ByVal n2 As Integer)
    Erase eventActiveAIScoretmpMain
    ReDim eventActiveAIScoretmpCardNum(0) As Integer
    eventActiveAIScoreFlag = True
    eventActiveAIScoretmpMain(1) = n1
    eventActiveAIScoretmpMain(2) = n2
End Sub
Public Sub EventActiveAIScoreAddCardNum(ByVal n1 As Integer)
    ReDim Preserve eventActiveAIScoretmpCardNum(UBound(eventActiveAIScoretmpCardNum) + 1) As Integer
    eventActiveAIScoretmpCardNum(UBound(eventActiveAIScoretmpCardNum)) = n1
End Sub
Private Sub eventActiveAIScoreOutput()
    Dim i As Integer

    Me.Text = Me.Text & "EventActiveAIScore" & "#" & CStr(eventActiveAIScoretmpMain(1)) & "," & CStr(eventActiveAIScoretmpMain(2))
    If UBound(eventActiveAIScoretmpCardNum) > 0 Then
        For i = 1 To UBound(eventActiveAIScoretmpCardNum)
            Me.Text = Me.Text & "," & eventActiveAIScoretmpCardNum(i)
        Next
    End If
    Me.Text = Me.Text & "="
End Sub
Public Sub BattleMoveControl(ByVal n1 As Integer)
    Me.Text = Me.Text & "BattleMoveControl" & "#" & CStr(n1) & "="
End Sub
Public Sub EventMoveActionOff()
    Me.Text = Me.Text & "EventMoveActionOff" & "#0" & "="
End Sub
Public Sub PersonMoveControl(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "PersonMoveControl" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub PersonMoveActionChange(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "PersonMoveActionChange" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub PersonAttackFirstControl(ByVal n1 As Integer)
    Me.Text = Me.Text & "PersonAttackFirstControl" & "#" & CStr(n1) & "="
End Sub
Public Sub EventTotalDiceChange(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "EventTotalDiceChange" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub EventPersonAbilityDiceChange(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "EventPersonAbilityDiceChange" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub PersonTotalDiceControl(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "PersonTotalDiceControl" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub AtkingTrueDiceControl(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "AtkingTrueDiceControl" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub PersonBloodControl(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer)
    Me.Text = Me.Text & "PersonBloodControl" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "," & CStr(n4) & "="
End Sub
Public Sub EventBloodActionOff()
    Me.Text = Me.Text & "EventBloodActionOff" & "#0" & "="
End Sub
Public Sub EventBloodActionChange(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "EventBloodActionChange" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub EventBloodReflection(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "EventBloodReflection" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub EventHPLActionOff()
    Me.Text = Me.Text & "EventHPLActionOff" & "#0" & "="
End Sub
Public Sub EventHPLActionChange(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "EventHPLActionChange" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub EventHPLReflection(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "EventHPLReflection" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub PersonAtkingOff(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "PersonAtkingOff" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub PersonPassiveOff(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "PersonPassiveOff" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub PersonAtkingOffSelect(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer)
    Me.Text = Me.Text & "PersonAtkingOffSelect" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "," & CStr(n4) & "="
End Sub
Public Sub PersonPassiveOffSelect(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer)
    Me.Text = Me.Text & "PersonPassiveOffSelect" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "," & CStr(n4) & "="
End Sub
Public Sub PersonAtkingInvalid(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "PersonAtkingInvalid" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub PersonResurrect(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "PersonResurrect" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub EventPersonResurrectActionOff()
    Me.Text = Me.Text & "EventPersonResurrectActionOff" & "#0" & "="
End Sub
Public Sub PersonChangeBattleImage(ByVal n1 As Integer, ByVal n2 As String)
    Me.Text = Me.Text & "PersonChangeBattleImage" & "#" & CStr(n1) & "," & n2 & "="
End Sub
Public Sub AtkingSeizeEnemyCards(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "AtkingSeizeEnemyCards" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub EventAtkingSeizeEnemyCardsActionOff()
    Me.Text = Me.Text & "EventAtkingSeizeEnemyCardsActionOff" & "#0" & "="
End Sub
Public Sub AtkingDrawCards(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "AtkingDrawCards" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub EventAtkingDrawCardsActionOff()
    Me.Text = Me.Text & "EventAtkingDrawCardsActionOff" & "#0" & "="
End Sub
Public Sub EventAtkingDrawCardsAddOnce()
    Me.Text = Me.Text & "EventAtkingDrawCardsAddOnce" & "#0" & "="
End Sub
Public Sub EventAtkingDrawCardsContinue()
    Me.Text = Me.Text & "EventAtkingDrawCardsContinue" & "#0" & "="
End Sub
Public Sub BattleDeckShuffle()
    Me.Text = Me.Text & "BattleDeckShuffle" & "#0" & "="
End Sub
Public Sub AtkingDestroyCards(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "AtkingDestroyCards" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub AtkingGiveCards(ByVal n1 As Integer)
    Me.Text = Me.Text & "AtkingGiveCards" & "#" & CStr(n1) & "="
End Sub
Public Sub AtkingGetUsedCards(ByVal n1 As Integer)
    Me.Text = Me.Text & "AtkingGetUsedCards" & "#" & CStr(n1) & "="
End Sub
Public Sub AtkingOneSelfCardControl(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "AtkingOneSelfCardControl" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub PersonMaxCardsNumControl(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer)
    Me.Text = Me.Text & "PersonMaxCardsNumControl" & "#" & CStr(n1) & "," & CStr(n2) & "," & CStr(n3) & "="
End Sub
Public Sub BattleInsertEventCard(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As String)
    Me.Text = Me.Text & "BattleInsertEventCard" & "#" & CStr(n1) & "," & CStr(n2) & "," & n3 & "="
End Sub
Public Sub PersonAddActualStatus(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As String, ByVal n4 As Integer)
    Me.Text = Me.Text & "PersonAddActualStatus" & "#" & CStr(n1) & "," & CStr(n2) & "," & n3 & "," & CStr(n4) & "="
End Sub
Public Sub EventAddActualStatusData(ByVal n1 As String, ByVal n2 As String, ByVal n3 As String, ByVal n4 As String, ByVal n5 As Integer, ByVal n6 As Integer, ByVal n7 As Integer, ByVal n8 As Integer)
    Me.Text = Me.Text & "EventAddActualStatusData" & "#" & n1 & "," & n2 & "," & n3 & "," & n4 & "," & CStr(n5) & "," & CStr(n6) & "," & CStr(n7) & "," & CStr(n8) & "="
End Sub
Public Sub ActualStatusEnd()
    Me.Text = Me.Text & "ActualStatusEnd" & "#0" & "="
End Sub
Public Sub PersonRemoveActualStatus(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "PersonRemoveActualStatus" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub EventRemoveActualStatusActionOff()
    Me.Text = Me.Text & "EventRemoveActualStatusActionOff" & "#0" & "="
End Sub
Public Sub PersonAddBuff(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As String, ByVal n4 As Integer, ByVal n5 As Integer)
    Me.Text = Me.Text & "PersonAddBuff" & "#" & CStr(n1) & "," & CStr(n2) & "," & n3 & "," & CStr(n4) & "," & CStr(n5) & "="
End Sub
Public Sub BuffTurnEnd()
    Me.Text = Me.Text & "BuffTurnEnd" & "#0" & "="
End Sub
Public Sub PersonRemoveBuffAll(ByVal n1 As Integer, ByVal n2 As Integer)
    Me.Text = Me.Text & "PersonRemoveBuffAll" & "#" & CStr(n1) & "," & CStr(n2) & "="
End Sub
Public Sub PersonRemoveBuffSelect(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As String)
    Me.Text = Me.Text & "PersonRemoveBuffSelect" & "#" & CStr(n1) & "," & CStr(n2) & "," & n3 & "="
End Sub
Public Sub PersonBuffTurnChange(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As String, ByVal n4 As Integer, ByVal n5 As Integer)
    Me.Text = Me.Text & "PersonBuffTurnChange" & "#" & CStr(n1) & "," & CStr(n2) & "," & n3 & "," & CStr(n4) & "," & CStr(n5) & "="
End Sub
Public Sub EventRemoveBuffActionOff()
    Me.Text = Me.Text & "EventRemoveBuffActionOff" & "#0" & "="
End Sub
