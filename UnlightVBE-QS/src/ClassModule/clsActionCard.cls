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
Public CardNum As Integer    '�d���s��
Public UpperType As String    '�W������[1]
Public UpperNum As Integer    '�W��ƭ�[2]
Public LowerType As String    '�U������[3]
Public LowerNum As Integer    '�U��ƭ�[4]
Public Owner As Integer    '�֦���[5]((0)�t��-(1)�ϥΪ�-(2)�q��)
Public Location As Integer    '���ݦ�m[6]((1)��P-(2)�X�P-(3)�Ӧa�P/��P-(4)�P��)
Public ImageStr As String    '�Ϥ��s��[8]
Public XYLeft As Integer    '�ثeLeft(�y��)[9]
Public XYTop As Integer    '�ثeTop(�y��)[10]
Public ComMark As Integer    '�q����аO[11]((1)�q����X�P(��)-(2)�q���o�X�P(�~)-(3)�ϥΪ̤�X�P�аO-(4)�ϥΪ̤�X�P�аO[�ݤ���])
Public CardOnIn As Integer    '�P�ثe���������(1.����/2.��P)
Public CardType As Integer    '�d���o������(1.���εP/2.�ƥ�d)

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