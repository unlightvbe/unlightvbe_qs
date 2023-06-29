VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActionCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CardNum As Integer    '卡片編號
Public UpperType As String    '上方類型[1]
Public UpperNum As Integer    '上方數值[2]
Public LowerType As String    '下方類型[3]
Public LowerNum As Integer    '下方數值[4]
Public Owner As Integer    '擁有者[5]((0)系統-(1)使用者-(2)電腦)
Public Location As Integer    '所屬位置[6]((1)手牌-(2)出牌-(3)墓地牌/棄牌-(4)牌堆)
Public ImageStr As String    '圖片編號[8]
Public XYLeft As Integer    '目前Left(座標)[9]
Public XYTop As Integer    '目前Top(座標)[10]
Public ComMark As Integer    '電腦方標記[11]((1)電腦方出牌(裏)-(2)電腦發出牌(外)-(3)使用者方出牌標記-(4)使用者方出牌標記[待反轉])
Public CardOnIn As Integer    '牌目前反轉紀錄數(1.正面/2.轉牌)
Public CardType As Integer    '卡片發行類型(1.公用牌/2.事件卡)

Public Sub Reverse()
    Dim tmpType As String, tmpNum As Integer

    tmpType = Me.UpperType
    tmpNum = Me.UpperNum
    Me.UpperType = Me.LowerType
    Me.UpperNum = Me.LowerNum
    Me.LowerType = tmpType
    Me.LowerNum = tmpNum

    If Me.CardOnIn = 1 Then
        Me.CardOnIn = 2
    Else
        Me.CardOnIn = 1
    End If
End Sub
