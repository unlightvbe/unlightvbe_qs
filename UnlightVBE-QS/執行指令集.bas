Attribute VB_Name = "������O��"
Option Explicit
Public vbecommadnum() As Integer '���涥�q���O���ܼ�-�ƭ���(1.�ثe������O����/2.�ثe������O�����q/3.�ثe����}������/4.�ثe�����涥�q��/5.�ثe���涥�q���O�`�p/6.�ثe�H������W����/7.�ثe�H�������ڽs��, ���涥�q���椤�p�ƭ�)
Public vbecommadstr() As String '���涥�q���O���ܼ�-�r����(1.�ثe������O�W��/2.�ثe���涥�q���O��, ���涥�q���椤�p�ƭ�)
Public vbecommadtotplay As Integer '�ثe���椧���涥�q�p�ƭ�
Public Vss_AtkingDrawCardsNum(1 To 3) As Integer '������O��-�ޯ��P�P�Ƭ����Ȯ��ܼ�(1.�w��i��[�t���L]/2.�w�p�`�i��/3.�w���L�i��)
Public Vss_AtkingSeizeEnemyCardsNum As Integer '������O��-�ܨ����d�P�����Ȯ��ܼ�
Public Vss_AtkingStartPlayNum(1 To 3) As Integer '������O��-�ޯ�ʵe��������Ȯ��ܼ�
Public Vss_PersonAtkingOffNum(1 To 2, 1 To 3, 1 To 8) As Integer '������O��-�T�����H���D�ʧޤγQ�ʧާޯ�����Ȯ��ܼ�(1.�ϥΪ�/2.�q��,1~3�H���s��,1~4.�D�ʧ޼аO/5~8.�Q�ʧ޼аO)
Public Vss_EventActiveAIScoreNum() As Integer '������O��-���z��AI�ӧO�ޯ���������Ȯ��ܼ�(1.�ӱƦC�զX�ޯ�����^�_/2.�����зǦ^�_/3~.�ޯ���ˤ��ӧO������˵P�s��)
Public Vss_PersonMoveControlNum(1 To 2, 1 To 2) As Integer  '������O��-���ʫe�`���ʶq����Ȯ��ܼ�(1.�ϥΪ̤�/2.�q����,1.�����ܤƶq/2.�O�_�����w)
Public Vss_AtkingInformationRecordStr(1 To 2, 1 To 3, 1 To 8) As String '������O��-�ޯ�Ƶ���T�x�s�Ȯ��ܼ�(1.�ϥΪ�/2.�q��,1~3�H���s��,�ޯ�ۦ�Ƶ��r��)
Public Vss_EventPlayerAllActionOffNum(1 To 2) As Integer '������O��-�T��a�i��Ҧ��ާ@�����Ȯ��ܼ�(1.�ϥΪ̤�/2.�q����)
Public Vss_PersonMoveActionChangeNum(1 To 2, 1 To 2) As Integer  '������O��-�H�����Ⲿ�ʶ��q��ʱ���Ȯ��ܼ�(1.�ϥΪ̤�/2.�q����,1.�O�_����/2.�����ܼ�)
Public Vss_PersonAttackFirstControlNum As Integer '������O��-�H�������u��������������Ȯ��ܼ�(1.�ϥΪ̤��/2.�q�����)
Public Vss_BattleStartDiceNum(0 To 5) As Integer '������O��-�����Y��l���q��T�Ȯ��ܼ�(0.���涥�q��/1.�ۨ��`���/2.����`���/3.�Y���ۨ�����ƶq/4.�Y����⥿��ƶq/5.�Y����`����ƶq)
Public Vss_EventPersonAbilityDiceChangeNum(1 To 2, 1 To 2) As Integer '������O��-�����O�����ܤƶq����Ȯ��ܼ�(1.�ϥΪ̤�/2.�q����,1.�ܤƶq/2.�O�_�����w)

Sub ������O���`�{��_�^�����O(ByVal str As String, ByVal ns As Integer, ByVal vbecommadtotplayNow As Integer)
      Dim commadstr1() As String
      vbecommadstr(2, vbecommadtotplayNow) = str
      vbecommadnum(1, vbecommadtotplayNow) = 1
      vbecommadnum(2, vbecommadtotplayNow) = 1
      '===============
      commadstr1 = Split(vbecommadstr(2, vbecommadtotplayNow), "=")
      vbecommadnum(5, vbecommadtotplayNow) = UBound(commadstr1)
End Sub
Sub ������O���`�{��_���涥�q����(ByVal ns As Integer, ByVal vbecommadtotplayNow As Integer)
    vbecommadnum(1, vbecommadtotplayNow) = 0
    vbecommadnum(2, vbecommadtotplayNow) = 0
    vbecommadnum(3, vbecommadtotplayNow) = 0
    vbecommadnum(4, vbecommadtotplayNow) = 0
    vbecommadstr(1, vbecommadtotplayNow) = ""
    vbecommadstr(2, vbecommadtotplayNow) = ""
End Sub
Sub ������O���`�{��_���O�I�s����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal ns As Integer, ByVal vbecommadtotplayNow As Integer)
'======commadtype(1.�@����涥�q/2.�ʵe���ĪG���涥�q)
     If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
     Dim cmdnumnow As Integer
     Dim PersonCheckAtking As Boolean
     Dim commadstr1()  As String, commadstr2() As String
     '===============================
     Do While vbecommadnum(1, vbecommadtotplayNow) <= vbecommadnum(5, vbecommadtotplayNow)
        commadstr1 = Split(vbecommadstr(2, vbecommadtotplayNow), "=")
        commadstr2 = Split(commadstr1(vbecommadnum(1, vbecommadtotplayNow) - 1), "#")
        vbecommadnum(2, vbecommadtotplayNow) = 1
        cmdnumnow = vbecommadnum(1, vbecommadtotplayNow)
        vbecommadstr(1, vbecommadtotplayNow) = commadstr2(0)
        vbecommadstr(3, vbecommadtotplayNow) = commadstr2(1)
        '=============================================
        PersonCheckAtking = ������O��_��������(uscom, commadtype, atkingnum, vbecommadtotplayNow)
        If PersonCheckAtking = False And _
               commadstr2(0) <> "AtkingLineLight" And commadstr2(0) <> "AtkingTurnOnOff" Then
               ������O��.������O_���O�����аO vbecommadtotplayNow
        Else
            Do
                Select Case commadstr2(0)
                        Case "AtkingLineLight"
                               ������O��.������O_�ޯ�O���� uscom, commadtype, atkingnum, vbecommadtotplayNow '(���q1)
                        Case "AtkingTurnOnOff"
                               ������O��.������O_�ޯ�ҰʽX���� uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                     '=======================================================
                        Case "EventTotalDiceChange"
                               ������O��.������O_�`����ܤƶq���� uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "PersonTotalDiceControl"
                               ������O��.������O_�`����`�q���� uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "PersonBloodControl"
                               ������O��.������O_�H����q���� uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "PersonAtkingInvalid"
                               ������O��.������O_�H���ޯ�L�Ĥ� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "BattleMoveControl"
                               ������O��.������O_���a�Z������ uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingStartPlay"
                               ������O��.������O_�ޯ�ʵe���� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingStartPlayAnimate"
                               ������O��.������O_�ޯ�ʵe����_�v��ʵe uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "AtkingSeizeEnemyCards"
                               ������O��.������O_�ܨ����d�P uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingDrawCards"
                               ������O��.������O_�P���P uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "BattleDeckShuffle"
                               ������O��.������O_�t�αj��~�P uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "BattleTurnControl"
                               ������O��.������O_�t�Φ^�X�Ʊ��� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingDestroyCards"
                               ������O��.������O_�֦��d�P��P uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingGiveCards"
                               ������O��.������O_�e�P�d�P uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingGetUsedCards"
                               ������O��.������O_�Ӧa�P�^�P uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "BattleSendMessage"
                               ������O��.������O_�ǰe�T�� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingTrueDiceControl"
                               ������O��.������O_������Ʊ��� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingOneSelfCardControl"
                               ������O��.������O_�֦����d�P���� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "BattleStartDice"
                               ������O��.������O_�����Y��l uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonMaxCardsNumControl"
                               ������O��.������O_�H���̤j�d��Ʊ��� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "BattleInsertEventCard"
                               ������O��.������O_���J�ƥ�d uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonAddBuff"
                               ������O��.������O_���`���A����_�[�J uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonRemoveBuffAll"
                               ������O��.������O_���`���A����_�����M��_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonRemoveBuffSelect"
                               ������O��.������O_���`���A����_�S�w�M��_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonBuffTurnChange"
                               ������O��.������O_���`���A����_�ܧ�^�X�� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonAddActualStatus"
                               ������O��.������O_�H����ڪ��A����_�[�J uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonRemoveActualStatus"
                               ������O��.������O_�H����ڪ��A����_�S�w�Ѱ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonAtkingOff"
                               ������O��.������O_�T�����H���D�ʧާޯ�_���� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonPassiveOff"
                               ������O��.������O_�T�����H���Q�ʧާޯ�_���� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonAtkingOffSelect"
                               ������O��.������O_�T�����H���D�ʧާޯ�_��� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonPassiveOffSelect"
                               ������O��.������O_�T�����H���Q�ʧާޯ�_��� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonMoveControl"
                               ������O��.������O_���ʫe�`���ʶq���� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonMoveActionChange"
                               ������O��.������O_�H�����Ⲿ�ʶ��q��ʱ��� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "PersonAttackFirstControl"
                               ������O��.������O_�H�������u���������� uscom, commadtype, atkingnum, vbecommadtotplayNow    '(���q1)
                        Case "AtkingInformationRecord"
                               ������O��.������O_�ޯ���O�Ƶ��r�� uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingLineLightAnother"
                               ������O��.������O_�ޯ�O����_��L uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "AtkingTurnOnOffAnother"
                               ������O��.������O_�ޯ�ҰʽX����_��L uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "PersonResurrect"
                               ������O��.������O_�H������_�� uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "EventPersonAbilityDiceChange"
                               ������O��.������O_�H������խȹ����ܤƶq���� uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "PersonChangeBattleImage"
                                ������O��.������O_�ܧ�H���԰���ø uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        '========================================================
                        Case "BuffTurnEnd"
                               ������O��.������O_���`���A����_��^�X����_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventBloodActionOff"
                               ������O��.������O_�H����q����_�ˮ`�L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventBloodActionChange"
                               ������O��.������O_�H����q����_�ˮ`�ĪG�ܧ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventBloodReflection"
                               ������O��.������O_�H����q����_�ˮ`�ĪG�Ϯg_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventHPLReflection"
                               ������O��.������O_�H����q����_�^�_�ĪG�Ϯg_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventHPLActionOff"
                               ������O��.������O_�H����q����_�^�_�L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventHPLActionChange"
                               ������O��.������O_�H����q����_�^�_�ĪG�ܧ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventMoveActionOff"
                               ������O��.������O_���a�Z������_�L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventRemoveBuffActionOff"
                               ������O��.������O_���椧���`���A�����L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventAddActualStatusData"
                               ������O��.������O_�H����ڪ��A�[�J���_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "ActualStatusEnd"
                               ������O��.������O_�H����ڪ��A����_�ŧi����_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventRemoveActualStatusActionOff"
                               ������O��.������O_���椧�H����ڪ��A�����L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventPlayerAllActionOff"
                                ������O��.������O_�T��a�i��Ҧ��ާ@ uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventPersonResurrectActionOff"
                                ������O��.������O_�H������_��_�L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventAtkingSeizeEnemyCardsActionOff"
                                ������O��.������O_�ܨ����d�P_�L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow   '(���q1)
                        Case "EventAtkingDrawCardsActionOff"
                                ������O��.������O_�P���P_�L�Ĥ�_�M uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "EventAtkingDrawCardsAddOnce"
                                ������O��.������O_�P���P_�ƶq�W�[_�M uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                        Case "EventAtkingDrawCardsContinue"
                                ������O��.������O_�P���P_���L_�M uscom, commadtype, atkingnum, vbecommadtotplayNow  '(���q1)
                     '========================================================
                        Case Else
                               GoTo vss_cmdlocalerr
                End Select
                DoEvents
            Loop Until vbecommadnum(1, vbecommadtotplayNow) > cmdnumnow
        End If
     Loop
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "Run-CommadNotFound[" & commadstr2(0) & "]", 0, vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O���`�{�ǰ���(ByVal cmdstr As String, ByVal vsscnum As Integer, ByVal uscom As Integer, ByVal atkingnum As Integer, ByVal ns As Integer, ByVal vbecommadtotplayNow As Integer)
     Dim commadtype As Integer
     vbecommadnum(3, vbecommadtotplayNow) = vsscnum
     vbecommadnum(4, vbecommadtotplayNow) = ns
     ������O��.������O���`�{��_�^�����O cmdstr, ns, vbecommadtotplayNow
     commadtype = ������O��.������O���`�{��_�P�_���涥�q���O(ns)
     ������O��.������O���`�{��_���O�I�s���� uscom, commadtype, atkingnum, ns, vbecommadtotplayNow
     ������O���`�{��_���涥�q���� ns, vbecommadtotplayNow
End Sub
Function ������O���`�{��_�P�_���涥�q���O(ByVal ns As Integer) As Integer
Select Case ns
    Case 42, 43, 44, 45, 92, 93, 94, 99 '�S��
        ������O���`�{��_�P�_���涥�q���O = 2
    Case 41, 46, 47, 48, 49, 61, 62, 72, 73, 74, 75, 76, 77, 101, 102, 103, 104, 105, 106, 107 '�ƥ�
        ������O���`�{��_�P�_���涥�q���O = 3
    Case Else  '���q��
        ������O���`�{��_�P�_���涥�q���O = 1
End Select
End Function
Sub ������O_�ޯ�O����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or vbecommadnum(3, vbecommadtotplayNow) > 48 Or _
       ((commadtype <> 1 And commadtype <> 3) And _
       Not (vbecommadnum(4, vbecommadtotplayNow) >= 42 And vbecommadnum(4, vbecommadtotplayNow) <= 44) And _
       Not (vbecommadnum(4, vbecommadtotplayNow) >= 92 And vbecommadnum(4, vbecommadtotplayNow) <= 94)) Then GoTo VssCommadExit
    If ����H����ԤH��(uscom, 2) <> vbecommadnum(7, vbecommadtotplayNow) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case vbecommadnum(3, vbecommadtotplayNow)
                Case Is <= 12 '==�D�ʧ�-�ϥΪ̤�
                        If ((uscom = 1 And liveus(����H����ԤH��(uscom, 2)) <= 0) Or _
                           (uscom = 2 And livecom(����H����ԤH��(uscom, 2)) <= 0)) And Val(commadstr3(0)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case Val(commadstr3(0))
                            Case 1
                                �԰��t����.�H���ޯ���O�}�� True, atkingnum
                            Case 2
                                �԰��t����.�H���ޯ���O�}�� False, atkingnum
                        End Select
                Case Is <= 24
                        GoTo VssCommadExit
                Case Is <= 48 '==�Q�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(0)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case uscom
                            Case 1
                                 Select Case Val(commadstr3(0))
                                    Case 1
                                        FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�O�o�G atkingnum - 4
                                    Case 2
                                        FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�O�ܷt atkingnum - 4
                                  End Select
                            Case 2
                                  Select Case Val(commadstr3(0))
                                    Case 1
                                        FormMainMode.PEAFInterface.Passive_�q��_�ޯ�O�o�G atkingnum - 4
                                    Case 2
                                        FormMainMode.PEAFInterface.Passive_�q��_�ޯ�O�ܷt atkingnum - 4
                                  End Select
                        End Select
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingLineLight", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ޯ�O����_��L(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or vbecommadnum(3, vbecommadtotplayNow) > 48 Or _
        (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    If ����H����ԤH��(uscom, 2) <> vbecommadnum(7, vbecommadtotplayNow) Then GoTo VssCommadExit
    If Val(commadstr3(1)) < 1 Or Val(commadstr3(1)) > 4 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(0))
                Case 1 '�D�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(2)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case Val(commadstr3(2))
                            Case 1
                                �԰��t����.�H���ޯ���O�}�� True, Val(commadstr3(1))
                            Case 2
                                �԰��t����.�H���ޯ���O�}�� False, Val(commadstr3(1))
                        End Select
                Case 2 '�Q�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(2)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case uscom
                            Case 1
                                 Select Case Val(commadstr3(2))
                                    Case 1
                                        FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�O�o�G Val(commadstr3(1))
                                    Case 2
                                        FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�O�ܷt Val(commadstr3(1))
                                  End Select
                            Case 2
                                  Select Case Val(commadstr3(2))
                                    Case 1
                                        FormMainMode.PEAFInterface.Passive_�q��_�ޯ�O�o�G Val(commadstr3(1))
                                    Case 2
                                        FormMainMode.PEAFInterface.Passive_�q��_�ޯ�O�ܷt Val(commadstr3(1))
                                  End Select
                        End Select
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingLineLightAnother", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ޯ�ҰʽX����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or vbecommadnum(3, vbecommadtotplayNow) > 48 Or _
       ((commadtype <> 1 And commadtype <> 3) And _
       Not (vbecommadnum(4, vbecommadtotplayNow) >= 42 And vbecommadnum(4, vbecommadtotplayNow) <= 44) And _
       Not (vbecommadnum(4, vbecommadtotplayNow) >= 92 And vbecommadnum(4, vbecommadtotplayNow) <= 94)) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case vbecommadnum(3, vbecommadtotplayNow)
                Case Is <= 24 '==�D�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(0)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case Val(commadstr3(0))
                            Case 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 1) = 1
                            Case 2
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 1) = 0
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 2) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 2)) + 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 3) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 3)) + 1
                        End Select
                Case Is <= 48 '==�Q�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(0)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case Val(commadstr3(0))
                            Case 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 1) = 1
                            Case 2
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 1) = 0
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 2) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 2)) + 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 3) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 3)) + 1
                        End Select
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingTurnOnOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ޯ�ҰʽX����_��L(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or vbecommadnum(3, vbecommadtotplayNow) > 48 Or _
       (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    If Val(commadstr3(1)) < 1 Or Val(commadstr3(1)) > 4 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(0))
                    Case 1 '�D�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(2)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case Val(commadstr3(2))
                            Case 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)), 1) = 1
                            Case 2
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)), 1) = 0
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)), 2) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)), 2)) + 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)), 3) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)), 3)) + 1
                        End Select
                Case 2 '�Q�ʧ�
                        If ((uscom = 1 And liveus(vbecommadnum(7, vbecommadtotplayNow)) <= 0) Or _
                           (uscom = 2 And livecom(vbecommadnum(7, vbecommadtotplayNow)) <= 0)) And Val(commadstr3(2)) = 1 Then
                           GoTo VssCommadExit
                        End If
                        Select Case Val(commadstr3(2))
                            Case 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)) + 4, 1) = 1
                            Case 2
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)) + 4, 1) = 0
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)) + 4, 2) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)) + 4, 2)) + 1
                                atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)) + 4, 3) = Val(atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), Val(commadstr3(1)) + 4, 3)) + 1
                        End Select
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingTurnOnOffAnother", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�`����ܤƶq����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or vbecommadnum(4, vbecommadtotplayNow) <> 45 Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case commadstr3(1)
                Case 1
                     If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                         atkingckdice(uscom, uscomt, 1) = atkingckdice(uscom, uscomt, 1) & "+" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                         atkingckdice(uscom, uscomt, 2) = atkingckdice(uscom, uscomt, 2) & "+" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                         atkingckdice(uscom, uscomt, 4) = atkingckdice(uscom, uscomt, 4) & "+" & commadstr3(2) & "="
                     Else
                         atkingckdice(uscom, uscomt, 3) = atkingckdice(uscom, uscomt, 3) & "+" & commadstr3(2) & "="
                     End If
                Case 2
                     If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                         atkingckdice(uscom, uscomt, 1) = atkingckdice(uscom, uscomt, 1) & "-" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                         atkingckdice(uscom, uscomt, 2) = atkingckdice(uscom, uscomt, 2) & "-" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                         atkingckdice(uscom, uscomt, 4) = atkingckdice(uscom, uscomt, 4) & "-" & commadstr3(2) & "="
                     Else
                         atkingckdice(uscom, uscomt, 3) = atkingckdice(uscom, uscomt, 3) & "-" & commadstr3(2) & "="
                     End If
                Case 3
                     If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                         atkingckdice(uscom, uscomt, 1) = atkingckdice(uscom, uscomt, 1) & "*" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                         atkingckdice(uscom, uscomt, 2) = atkingckdice(uscom, uscomt, 2) & "*" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                         atkingckdice(uscom, uscomt, 4) = atkingckdice(uscom, uscomt, 4) & "*" & commadstr3(2) & "="
                     Else
                         atkingckdice(uscom, uscomt, 3) = atkingckdice(uscom, uscomt, 3) & "*" & commadstr3(2) & "="
                     End If
                Case 4
                     If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                         atkingckdice(uscom, uscomt, 1) = atkingckdice(uscom, uscomt, 1) & "\" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                         atkingckdice(uscom, uscomt, 2) = atkingckdice(uscom, uscomt, 2) & "\" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                         atkingckdice(uscom, uscomt, 4) = atkingckdice(uscom, uscomt, 4) & "\" & commadstr3(2) & "="
                     Else
                         atkingckdice(uscom, uscomt, 3) = atkingckdice(uscom, uscomt, 3) & "\" & commadstr3(2) & "="
                     End If
                Case 5
                     If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                         atkingckdice(uscom, uscomt, 1) = atkingckdice(uscom, uscomt, 1) & "/" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                         atkingckdice(uscom, uscomt, 2) = atkingckdice(uscom, uscomt, 2) & "/" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                         atkingckdice(uscom, uscomt, 4) = atkingckdice(uscom, uscomt, 4) & "/" & commadstr3(2) & "="
                     Else
                         atkingckdice(uscom, uscomt, 3) = atkingckdice(uscom, uscomt, 3) & "/" & commadstr3(2) & "="
                     End If
                Case 6
                     If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                         atkingckdice(uscom, uscomt, 1) = atkingckdice(uscom, uscomt, 1) & "@" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                         atkingckdice(uscom, uscomt, 2) = atkingckdice(uscom, uscomt, 2) & "@" & commadstr3(2) & "="
                     ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                         atkingckdice(uscom, uscomt, 4) = atkingckdice(uscom, uscomt, 4) & "@" & commadstr3(2) & "="
                     Else
                         atkingckdice(uscom, uscomt, 3) = atkingckdice(uscom, uscomt, 3) & "@" & commadstr3(2) & "="
                     End If
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventTotalDiceChange", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�`����`�q����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (vbecommadnum(4, vbecommadtotplayNow) <> 10 And vbecommadnum(4, vbecommadtotplayNow) <> 11 And vbecommadnum(4, vbecommadtotplayNow) <> 30 And vbecommadnum(4, vbecommadtotplayNow) <> 31) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case commadstr3(1)
                Case 1
                     �������m��l�`��(uscomt) = �������m��l�`��(uscomt) + Val(commadstr3(2))
                Case 2
                     �������m��l�`��(uscomt) = �������m��l�`��(uscomt) - Val(commadstr3(2))
                Case 3
                     �������m��l�`��(uscomt) = �������m��l�`��(uscomt) * Val(commadstr3(2))
                Case 4
                     �������m��l�`��(uscomt) = �������m��l�`��(uscomt) \ Val(commadstr3(2))
                Case 5
                     �������m��l�`��(uscomt) = Int(�������m��l�`��(uscomt) / Val(commadstr3(2)) + 0.9)
                Case 6
                     �������m��l�`��(uscomt) = Val(commadstr3(2))
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonTotalDiceControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 3 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) = 46 Or vbecommadnum(4, vbecommadtotplayNow) = 48 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    '=====================
    Dim uscomt As Integer, statusnum As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case atkingnum
                Case Is <= 4
                    statusnum = 1
                Case Is <= 8
                    statusnum = 2
                Case 9
                    statusnum = 3
                Case 10
                    statusnum = 4
            End Select
            '===============================�[�J�Ӷ��q������T
            Dim stageInfoListObj As New clsVSStageObj
            stageInfoListObj.StageNum = vbecommadtotplayNow
            stageInfoListObj.CommandStr = "PersonBloodControl"
            stageInfoListObj.Value = "0"
            ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
            '===============================
            Select Case commadstr3(2)
                Case 1
                     ReDim VBEStageNum(0 To 6) As Integer
                     VBEStageNum(5) = -uscom
                     VBEStageNum(6) = statusnum
                     Select Case uscomt
                          Case 1
                                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� commadstr3(3), commadstr3(1), True
                          Case 2
                                �԰��t����.�ˮ`����_�ޯઽ��_�q�� commadstr3(3), commadstr3(1), True
                     End Select
                Case 2
                     ReDim VBEStageNum(0 To 5) As Integer
                     VBEStageNum(4) = -uscom
                     VBEStageNum(5) = statusnum
                     Select Case uscomt
                          Case 1
                                �԰��t����.�^�_����_�ϥΪ� commadstr3(3), commadstr3(1), statusnum, True, False
                          Case 2
                                �԰��t����.�^�_����_�q�� commadstr3(3), commadstr3(1), statusnum, True, False
                     End Select
                Case 3
                     ReDim VBEStageNum(0 To 6) As Integer
                     VBEStageNum(5) = -uscom
                     VBEStageNum(6) = statusnum
                     Select Case uscomt
                          Case 1
                                �԰��t����.�ˮ`����_�ߧY���`_�ϥΪ� commadstr3(1)
                          Case 2
                                �԰��t����.�ˮ`����_�ߧY���`_�q�� commadstr3(1)
                     End Select
            End Select
            ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonBloodControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H������_��(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case Else
            GoTo VssCommadExit
    End Select
    '=====================
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    If Val(commadstr3(1)) < 1 Or Val(commadstr3(1)) > ����H����ԤH��(uscomt, 1) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            '===================================
            Dim stageInfoListObj As New clsVSStageObj
            stageInfoListObj.StageNum = vbecommadtotplayNow
            stageInfoListObj.CommandStr = "PersonResurrect"
            stageInfoListObj.Value = "0"
            ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
            '===================================
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(3) = -uscom
            '===================================
            Select Case uscomt
                Case 1
                     �԰��t����.����_��_�ϥΪ� Val(commadstr3(1))
                Case 2
                     �԰��t����.����_��_�q�� Val(commadstr3(1))
            End Select
            
            ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonResurrect", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H���ޯ�L�Ĥ�(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim i As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case commadstr3(1)
                Case 1 '==�D�ʧ�
                        For i = 1 To 4
                            atkingck(uscomt, ����H����ԤH��(uscomt, 2), i, 1) = 0
                            �԰��t����.�H���ޯ���O�}�� False, i
                        Next
                        atkingckdice(uscomt, uscom, 1) = 0
                        atkingckdice(uscomt, uscomt, 1) = 0
                Case 2 '==�Q�ʧ�
                        For i = 5 To 8
                            atkingck(uscomt, ����H����ԤH��(uscomt, 2), i, 1) = 0
                        Next
                        atkingckdice(uscomt, uscom, 2) = 0
                        atkingckdice(uscomt, uscomt, 2) = 0
            End Select
            �԰��t����.��q��s���
            '============
            FormMainMode.trgoi1_Timer
            FormMainMode.trgoi2_Timer
            '============
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonAtkingInvalid", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���a�Z������(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) = 47 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    '=====================
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(3) = -uscom
            '===============================�[�J�Ӷ��q������T
            Dim stageInfoListObj As New clsVSStageObj
            stageInfoListObj.StageNum = vbecommadtotplayNow
            stageInfoListObj.CommandStr = "BattleMoveControl"
            stageInfoListObj.Value = "0"
            ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
            '===============================
            Select Case Val(commadstr3(0))
                Case 1
                    �԰��t����.����ʧ@_�Z���ܧ� 1, True, False
                Case 2
                    �԰��t����.����ʧ@_�Z���ܧ� 2, True, False
                Case 3
                    �԰��t����.����ʧ@_�Z���ܧ� 3, True, False
            End Select
            ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BattleMoveControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ޯ�ʵe����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or commadtype <> 1 Or atkingnum = 9 Or (vbecommadnum(4, vbecommadtotplayNow) = 13 Or vbecommadnum(4, vbecommadtotplayNow) = 33) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If commadstr3(1) = "0" Then commadstr3(1) = ""
            ������O��.Sub_�ޯ�ʵe����_�R�A commadstr3(0), commadstr3(1), uscom
            vbecommadnum(2, vbecommadtotplayNow) = 3 '==���ݮɶ�
        Case 3
            If Vss_AtkingStartPlayNum(2) = 1 Then
                Dim vbecommadnumSecond As Integer '���h���涥�q�s����
                '=======================
                vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
                '=======================
                Dim VBEStageNumMainSec(1 To 1) As Integer
                Dim buffvssnum As String
                If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_�H���D�ʧޯ� uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_�H���Q�ʧޯ� uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_�H����ڪ��A uscom, vbecommadnum(7, vbecommadtotplayNow), 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                Else
                    buffvssnum = VBEVSSBuffStr1(vbecommadnum(3, vbecommadtotplayNow) - 54)
                    If CollectionExists(�H�����`���A�C��(uscom, vbecommadnum(7, vbecommadtotplayNow)), buffvssnum) = True Then
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���`���A uscom, vbecommadnum(7, vbecommadtotplayNow), buffvssnum, 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                    End If
                End If
                '=======================
                ���涥�q�t��_�ŧi�}�l�ε��� 2
                vbecommadnum(2, vbecommadtotplayNow) = 4 '==���ݮɶ�
            End If
        Case 4
            If Vss_AtkingStartPlayNum(3) = 1 Then
                FormMainMode.Enabled = True
                GoTo VssCommadExit
            End If
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingStartPlay", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ޯ�ʵe����_�v��ʵe(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim piclist As Collection, i As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) > 2 Or commadtype <> 1 Or atkingnum = 9 Or (vbecommadnum(4, vbecommadtotplayNow) = 13 Or vbecommadnum(4, vbecommadtotplayNow) = 33) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If Formsetting.chkAtkingAnimateDisable.Value = 1 Then
                Dim tmpstr1 As String, tmpstr2 As String
                Select Case UBound(commadstr3)
                    Case 0
                        tmpstr1 = commadstr3(0) & "4.png"
                        tmpstr2 = commadstr3(0) & "12.png"
                    Case 1
                        tmpstr1 = commadstr3(0) & commadstr3(1) & ".png"
                        tmpstr2 = ""
                    Case 2
                        tmpstr1 = commadstr3(0) & commadstr3(1) & ".png"
                        tmpstr2 = commadstr3(0) & commadstr3(2) & ".png"
                End Select
                ������O��.Sub_�ޯ�ʵe����_�R�A tmpstr1, tmpstr2, uscom
            Else
                Set piclist = New Collection
                For i = 1 To 24
                    Dim filestr As String
                    filestr = App.Path & commadstr3(0) & i & ".png"
                    If Dir(filestr) <> "" Then
                        piclist.Add filestr
                    Else
                        If i <= 16 Then
                            GoTo vss_cmdlocalerr
                        Else
                            Exit For
                        End If
                    End If
                Next
                FormMainMode.PEAFAnimateInterface.AnimatePictureList = piclist
                FormMainMode.PEAFAnimateInterface.uscom = uscom
                '=======================
                Erase Vss_AtkingStartPlayNum
                FormMainMode.PEAFAnimateInterface.ZOrder
                FormMainMode.Enabled = False
                FormMainMode.PEAFAnimateInterface.AnimateStart
            End If
            vbecommadnum(2, vbecommadtotplayNow) = 3 '==���ݮɶ�
        Case 3
            If Vss_AtkingStartPlayNum(2) = 1 Then
                Dim vbecommadnumSecond As Integer '���h���涥�q�s����
                '=======================
                vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
                '=======================
                Dim VBEStageNumMainSec(1 To 1) As Integer
                Dim buffvssnum As String
                If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_�H���D�ʧޯ� uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_�H���Q�ʧޯ� uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_�H����ڪ��A uscom, vbecommadnum(7, vbecommadtotplayNow), 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                Else
                    buffvssnum = VBEVSSBuffStr1(vbecommadnum(3, vbecommadtotplayNow) - 54)
                    If CollectionExists(�H�����`���A�C��(uscom, vbecommadnum(7, vbecommadtotplayNow)), buffvssnum) = True Then
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���`���A uscom, vbecommadnum(7, vbecommadtotplayNow), buffvssnum, 61, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
                    End If
                End If
                '=======================
                ���涥�q�t��_�ŧi�}�l�ε��� 2
                vbecommadnum(2, vbecommadtotplayNow) = 4 '==���ݮɶ�
            End If
        Case 4
            If Vss_AtkingStartPlayNum(3) = 1 Then
                FormMainMode.Enabled = True
                GoTo VssCommadExit
            End If
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingStartPlayAnimate", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub Sub_�ޯ�ʵe����_�R�A(ByVal commadstr3_1 As String, ByVal commadstr3_2 As String, ByVal uscom As Integer)
    Dim piclist As Collection
    Set piclist = New Collection
    piclist.Add App.Path & commadstr3_1
    If commadstr3_2 <> "" Then
        piclist.Add App.Path & commadstr3_2
    End If
    FormMainMode.PEAFAnimateInterface.AnimatePictureList = piclist
    FormMainMode.PEAFAnimateInterface.uscom = uscom
    '=======================
    Erase Vss_AtkingStartPlayNum
    FormMainMode.PEAFAnimateInterface.ZOrder
    FormMainMode.Enabled = False
    FormMainMode.PEAFAnimateInterface.AnimateStart
End Sub
Sub ������O_�ܨ����d�P(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String, tmpflag As Boolean, tmpcollectionIndex As Integer
    Dim tmpcard As clsActionCard
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (commadtype <> 1 And commadtype <> 3) Or vbecommadnum(4, vbecommadtotplayNow) = 101 Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case uscom
         Case 1
               uscomt = 2
         Case 2
               uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            tmpcollectionIndex = �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(1)))
            If (Val(commadstr3(0)) = 1 And ((uscomt = 1 And tmpcollectionIndex <> 5) Or (uscomt = 2 And tmpcollectionIndex <> 7))) Or _
                (Val(commadstr3(0)) = 2 And ((uscomt = 1 And tmpcollectionIndex <> 6) Or (uscomt = 2 And tmpcollectionIndex <> 8))) Or _
                (Val(commadstr3(0)) <> 1 And Val(commadstr3(0)) <> 2) Then
                GoTo VssCommadExit
            End If
            Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(1))))(CStr(�԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(1)))))
            
            If tmpcard.Location <> Val(commadstr3(0)) Then GoTo VssCommadExit
            '===============================�[�J�Ӷ��q������T
            Dim stageInfoListObj As New clsVSStageObj
            stageInfoListObj.StageNum = vbecommadtotplayNow
            stageInfoListObj.CommandStr = "AtkingSeizeEnemyCards"
            stageInfoListObj.Value = "0"
            ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
            '===============================
            ReDim VBEStageNum(0 To 7) As Integer
            VBEStageNum(0) = 101
            VBEStageNum(1) = -uscom 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = Val(commadstr3(1)) '���v�T���d�P�s������
            VBEStageNum(3) = ���涥�q�t����.���涥�q�t��_�ǳ��ܼƲΦX_pagecardnum_type(tmpcard.UpperType) '���v�T���d�P��������
            VBEStageNum(4) = tmpcard.UpperNum '���v�T���d�P�����ƭ�
            VBEStageNum(5) = ���涥�q�t����.���涥�q�t��_�ǳ��ܼƲΦX_pagecardnum_type(tmpcard.LowerType) '���v�T���d�P�ϭ�����
            VBEStageNum(6) = tmpcard.LowerNum '���v�T���d�P�ϭ��ƭ�
            VBEStageNum(7) = tmpcard.Location
            '===========================���涥�q���J�I(101)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 101, 1
            '============================
            tmpflag = False
            If stageInfoListObj.CommandStr = "AtkingSeizeEnemyCards" Then
                If stageInfoListObj.Value = "OFF" Then
                    tmpflag = True
                End If
            End If
            ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
            
            If tmpflag = True Then GoTo VssCommadExit
            '=======================
            Select Case Val(commadstr3(0))
                Case 1  '==��P
                    Select Case uscomt
                        Case 1
                            If tmpcard.Location = 1 And tmpcard.Owner = 1 Then
                                �ثe��(20) = tmpcard.CardNum
                                �ثe��(21) = 2
                                FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
                                vbecommadnum(2, vbecommadtotplayNow) = 0
                            Else
                                GoTo VssCommadExit
                            End If
                        Case 2
                            If tmpcard.Location = 1 And tmpcard.Owner = 2 Then
                                �ثe��(16) = tmpcard.CardNum
                                FormMainMode.tr�q���P_½�P.Enabled = True
                                vbecommadnum(2, vbecommadtotplayNow) = 0
                            Else
                                GoTo VssCommadExit
                            End If
                    End Select
                Case 2  '==�X�P
                    Select Case uscomt
                        Case 1
                            If tmpcard.Location = 2 And tmpcard.Owner = 1 Then
                                turnpageoninatking = 1
                                FormMainMode.card_CardClick tmpcard.CardNum
                                vbecommadnum(2, vbecommadtotplayNow) = 0
                            Else
                                GoTo VssCommadExit
                            End If
                        Case 2
                            If tmpcard.Location = 2 And tmpcard.Owner = 2 Then
                                �԰��t����.�q���P_�������P_�~ tmpcard.CardNum
                                vbecommadnum(2, vbecommadtotplayNow) = 0
                            Else
                                GoTo VssCommadExit
                            End If
                    End Select
            End Select
        Case 2
            Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(1))))(CStr(�԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(1)))))
            Select Case Val(commadstr3(0))
                 Case 1  '==��P
                    Select Case uscomt
                        Case 1
                            GoTo VssCommadExit
                        Case 2
                            �ثe��(17) = 3
                            FormMainMode.tr�q���P_���P.Enabled = True
                            vbecommadnum(2, vbecommadtotplayNow) = 0
                    End Select
                 Case 2  '==�X�P
                    Select Case uscomt
                        Case 1
                            �ثe��(21) = 1
                            Vss_AtkingSeizeEnemyCardsNum = �ثe��(5)
                            '=========�N�y�Ы��w�ܹq����P
                            �԰��t����.�y�Эp��_�q����P
                            �԰��t����.����ʧ@_�ϥΪ̵P_���P_�q�� tmpcard.CardNum
                            �ثe��(5) = Vss_AtkingSeizeEnemyCardsNum
                            �ثe��(15) = 23
                            turnpageoninatking = 0
                            vbecommadnum(2, vbecommadtotplayNow) = 0
                        Case 2
                            �ثe��(17) = 2
                            Vss_AtkingSeizeEnemyCardsNum = �ثe��(9)
                            '=========�N�y�Ы��w�ܨϥΪ̤�P
                            �԰��t����.�y�Эp��_�ϥΪ̤�P
                            �԰��t����.����ʧ@_�q���P_���P_�ϥΪ� tmpcard.CardNum
                            �԰��t����.���εP�^�_���� tmpcard.CardNum
                            �ثe��(9) = Vss_AtkingSeizeEnemyCardsNum
                            �ثe��(15) = 23
                            vbecommadnum(2, vbecommadtotplayNow) = 0
                    End Select
            End Select
        Case 3
            Select Case Val(commadstr3(0))
                Case 1  '==��P
                      GoTo VssCommadExit
                Case 2  '==�X�P
                      GoTo VssCommadExit
            End Select
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingSeizeEnemyCards", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ܨ����d�P_�L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 101 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "AtkingSeizeEnemyCards" Then
                    stageInfoListObj.Value = "OFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventAtkingSeizeEnemyCardsActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�P���P(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String, ay() As String
    Dim tmpcard As clsActionCard
    Dim stageInfoListObj As clsVSStageObj
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    Dim tn As Integer, tmpflag As Boolean '�Ȯ��ܼ�
    If UBound(commadstr3) <> 2 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            '===============================�[�J�Ӷ��q������T
            Set stageInfoListObj = New clsVSStageObj
            stageInfoListObj.StageNum = vbecommadtotplayNow
            stageInfoListObj.CommandStr = "AtkingDrawCards"
            stageInfoListObj.Value = "0"
            ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
            '===============================
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 102
            VBEStageNum(1) = -uscom 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            '===========================���涥�q���J�I(102)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 102, 1
            '============================
            tmpflag = False
            If stageInfoListObj.CommandStr = "AtkingDrawCards" Then
                If stageInfoListObj.Value = "OFF%" Then
                    tmpflag = True
                End If
            End If
            ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
            
            If tmpflag = True Then GoTo VssCommadExit
            '=======================
            Vss_AtkingDrawCardsNum(1) = 0
            Vss_AtkingDrawCardsNum(2) = Val(commadstr3(2))
            Vss_AtkingDrawCardsNum(3) = 0
            vbecommadnum(2, vbecommadtotplayNow) = 2
        Case 2
             If Vss_AtkingDrawCardsNum(1) = 0 Then
                 If BattleCardNum < Val(commadstr3(2)) Then
                   �԰��t����.����ʧ@_�~�P
                End If
             End If
             If Vss_AtkingDrawCardsNum(1) < Vss_AtkingDrawCardsNum(2) And Vss_AtkingDrawCardsNum(1) < BattleCardNum Then
                Vss_AtkingDrawCardsNum(1) = Vss_AtkingDrawCardsNum(1) + 1
                vbecommadnum(2, vbecommadtotplayNow) = 0
                Select Case Val(commadstr3(1))
                    Case 1  '==���εP
                        �ثe��(15) = 21
                        '===============================�[�J�Ӷ��q������T
                        Set stageInfoListObj = New clsVSStageObj
                        stageInfoListObj.StageNum = vbecommadtotplayNow
                        stageInfoListObj.CommandStr = "AtkingDrawCardsEvent"
                        stageInfoListObj.Value = ""
                        ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
                        '=========================
                        �԰��t����.����ʧ@_��P_���εP uscomt, Vss_AtkingDrawCardsNum(3) + 1, True
                        '=========================
                        ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
                     Case 2  '==�ƥ�d
                        Select Case uscomt
                            Case 1
                                If �԰��t����.CardDeckCollection(3).Count > 0 Then
                                    Set tmpcard = �԰��t����.CardDeckCollection(3)(1)
                                    �ثe��(16) = tmpcard.CardNum
                                    BattleTurn = BattleTurn + 1
                                    FormMainMode.PEAFInterface.turn = BattleTurn
                                    �ثe��(15) = 21
                                    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 3, 5
                                    FormMainMode.tr�P��_�^�P_�ϥΪ�.Enabled = True
                                Else
                                    GoTo VssCommadExit
                                End If
                            Case 2
                                If �԰��t����.CardDeckCollection(4).Count > 0 Then
                                    Set tmpcard = �԰��t����.CardDeckCollection(4)(1)
                                    �ثe��(16) = tmpcard.CardNum
                                    BattleTurn = BattleTurn + 1
                                    FormMainMode.PEAFInterface.turn = BattleTurn
                                    �ثe��(15) = 21
                                    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 4, 7
                                    FormMainMode.tr�P��_�^�P_�q��.Enabled = True
                                Else
                                    GoTo VssCommadExit
                                End If
                        End Select
                End Select
             Else
                GoTo VssCommadExit
             End If
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingDrawCards", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�P���P_�L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or (Val(vbecommadnum(4, vbecommadtotplayNow)) <> 102 And Val(vbecommadnum(4, vbecommadtotplayNow)) <> 103) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And (stageInfoListObj.CommandStr = "AtkingDrawCards" Or stageInfoListObj.CommandStr = "AtkingDrawCardsEvent") Then
                    stageInfoListObj.Value = "OFF%"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventAtkingDrawCardsActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�P���P_�ƶq�W�[_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 103 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "AtkingDrawCardsEvent" Then
                    stageInfoListObj.Value = stageInfoListObj.Value + "AddOnce%"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventAtkingDrawCardsAddOnce", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�P���P_���L_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 103 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "AtkingDrawCardsEvent" Then
                    stageInfoListObj.Value = stageInfoListObj.Value + "Continue%"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventAtkingDrawCardsContinue", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�t�αj��~�P(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            �԰��t����.����ʧ@_�~�P
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BattleDeckShuffle", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�t�Φ^�X�Ʊ���(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(0))
                 Case 1
                       If BattleTurn + Val(commadstr3(1)) <= 18 Then
                          BattleTurn = BattleTurn + Val(commadstr3(1))
                       End If
                 Case 2
                       If BattleTurn - Val(commadstr3(1)) >= 1 Then
                          BattleTurn = BattleTurn - Val(commadstr3(1))
                       End If
            End Select
            FormMainMode.PEAFInterface.turn = BattleTurn
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BattleTurnControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�֦��d�P��P(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case uscomt
                Case 1
                    If �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(1))) = 5 Then
                        �ثe��(20) = �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(1)))
                        �ثe��(21) = 4
                        FormMainMode.tr�ϥΪ�_��P.Enabled = True
                        vbecommadnum(2, vbecommadtotplayNow) = 0
                    Else
                        GoTo VssCommadExit
                    End If
                Case 2
                    If �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(1))) = 7 Then
                        �ثe��(16) = �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(1)))
                        FormMainMode.tr�q���P_½�P.Enabled = True
                        vbecommadnum(2, vbecommadtotplayNow) = 0
                    Else
                        GoTo VssCommadExit
                    End If
            End Select
        Case 2
            Select Case uscomt
                Case 1
                    GoTo VssCommadExit
                Case 2
                    FormMainMode.tr�q���P_��P.Enabled = True
                    �ثe��(17) = 4
                    vbecommadnum(2, vbecommadtotplayNow) = 0
            End Select
        Case 3
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingDestroyCards", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�e�P�d�P(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case uscom
                 Case 1 '==�ϥΪ̤�
                    If �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(0))) = 5 Then
                        �ثe��(20) = �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(0)))
                        �ثe��(21) = 5
                        FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
                        vbecommadnum(2, vbecommadtotplayNow) = 0
                    Else
                        GoTo VssCommadExit
                    End If
                 Case 2 '==�q����
                    If �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(0))) = 7 Then
                        �ثe��(16) = �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(0)))
                        FormMainMode.tr�q���P_½�P.Enabled = True
                        vbecommadnum(2, vbecommadtotplayNow) = 0
                    Else
                        GoTo VssCommadExit
                    End If
            End Select
        Case 2
            Select Case uscom
                 Case 1 '==�ϥΪ̤�
                      GoTo VssCommadExit
                 Case 2 '==�q����
                       �ثe��(17) = 5
                        FormMainMode.tr�q���P_���P.Enabled = True
                        vbecommadnum(2, vbecommadtotplayNow) = 0
            End Select
        Case 3
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingGiveCards", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�Ӧa�P�^�P(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim tmpcard As clsActionCard, tmpcollectionIndex As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            tmpcollectionIndex = �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(0)))
            If tmpcollectionIndex <> 2 And tmpcollectionIndex <> 9 Then GoTo VssCommadExit
            
            Set tmpcard = �԰��t����.CardDeckCollection(tmpcollectionIndex)(CStr(�԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(0)))))
            Select Case uscom
                Case 1
                    If tmpcard.Location = 3 Then
                        �ثe��(16) = tmpcard.CardNum
                        �ثe��(15) = 22
                        FormMainMode.tr�P��_�^�P_�ϥΪ�.Enabled = True
                        vbecommadnum(2, vbecommadtotplayNow) = 0
                    Else
                        GoTo VssCommadExit
                    End If
                Case 2
                    If tmpcard.Location = 3 Then
                       �ثe��(16) = tmpcard.CardNum
                       �ثe��(15) = 22
                       FormMainMode.tr�P��_�^�P_�q��.Enabled = True
                       vbecommadnum(2, vbecommadtotplayNow) = 0
                    Else
                       GoTo VssCommadExit
                    End If
            End Select
        Case 2
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingGetUsedCards", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ǰe�T��(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            �԰��t����.�s���T�� commadstr3(0)
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BattleSendMessage", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�T�����H���D�ʧާޯ�_����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim i As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (commadtype <> 1 And commadtype <> 3) Or atkingnum <= 8 Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(2))
                 Case 1
                        For i = 1 To 4
                            Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), i) = 1
                        Next
                 Case 2
                        For i = 1 To 4
                            Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), i) = 0
                        Next
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonAtkingOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�T�����H���Q�ʧާޯ�_����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim i As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (commadtype <> 1 And commadtype <> 3) Or atkingnum <= 8 Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(2))
                 Case 1
                        For i = 5 To 8
                            Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), i) = 1
                        Next
                 Case 2
                        For i = 5 To 8
                            Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), i) = 0
                        Next
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonPassiveOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�T�����H���D�ʧާޯ�_���(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim uscomt As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 3 Or (commadtype <> 1 And commadtype <> 3) Or atkingnum <= 8 Then GoTo VssCommadExit
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(2))
                 Case 1
                        Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), Val(commadstr3(3))) = 1
                 Case 2
                        Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), Val(commadstr3(3))) = 0
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonAtkingOffSelect", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�T�����H���Q�ʧާޯ�_���(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim uscomt As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 3 Or (commadtype <> 1 And commadtype <> 3) Or atkingnum <= 8 Then GoTo VssCommadExit
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(2))
                 Case 1
                        Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), Val(commadstr3(3)) + 4) = 1
                 Case 2
                        Vss_PersonAtkingOffNum(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), Val(commadstr3(3)) + 4) = 0
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonPassiveOffSelect", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����_�ˮ`�L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 46 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If (stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "PersonBloodControl") Or (stageInfoListObj.StageNum = 0 And stageInfoListObj.CommandStr = "@SystemBloodAction") Then
                    stageInfoListObj.Value = "BLOODOFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventBloodActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����_�ˮ`�ĪG�ܧ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 46 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count = 0 Then GoTo VssCommadExit
            
            Dim stageInfoListObj As clsVSStageObj
            Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
            
            Dim tmparg() As String, tmpNum As Integer
            tmparg = Split(stageInfoListObj.Argument, "%")
            
            If (stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "PersonBloodControl") Or (stageInfoListObj.StageNum = 0 And stageInfoListObj.CommandStr = "@SystemBloodAction") Then
                tmpNum = Val(tmparg(0))
                Select Case Val(commadstr3(0))
                    Case 1
                        If Val(tmparg(3)) < 3 Then '�ư��ߧY���`
                            tmpNum = Val(tmparg(0)) + Val(commadstr3(1))
                        End If
                    Case 2
                        If Val(tmparg(3)) < 3 Then '�ư��ߧY���`
                            tmpNum = Val(tmparg(0)) - Val(commadstr3(1))
                        End If
                    Case 3
                        tmpNum = Val(commadstr3(1))
                End Select
            End If
            stageInfoListObj.Value = "BLOODCHANGE%" + str(tmpNum)
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventBloodActionChange", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����_�ˮ`�ĪG�Ϯg_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim uscomt As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (Val(vbecommadnum(4, vbecommadtotplayNow)) <> 46 And Val(vbecommadnum(4, vbecommadtotplayNow)) <> 48) Then GoTo VssCommadExit
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case uscomt
                Case 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� commadstr3(2), commadstr3(1), False
                Case 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� commadstr3(2), commadstr3(1), False
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventBloodReflection", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����_�^�_�ĪG�Ϯg_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim uscomt As Integer, statusnum As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (Val(vbecommadnum(4, vbecommadtotplayNow)) <> 46 And Val(vbecommadnum(4, vbecommadtotplayNow)) <> 48) Then GoTo VssCommadExit
    Select Case Val(commadstr3(0))
         Case 1
            uscomt = uscom
         Case 2
            If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case atkingnum
        Case Is <= 4
            statusnum = 1
        Case Is <= 8
            statusnum = 2
        Case 9
            statusnum = 3
        Case 10
            statusnum = 4
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case uscomt
                Case 1
                    �԰��t����.�^�_����_�ϥΪ� commadstr3(2), commadstr3(1), statusnum, False, False
                Case 2
                    �԰��t����.�^�_����_�q�� commadstr3(2), commadstr3(1), statusnum, False, False
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventHPLReflection", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����_�^�_�L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 48 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If (stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "PersonBloodControl") Or (stageInfoListObj.StageNum = 0 And stageInfoListObj.CommandStr = "@SystemHPLAction") Then
                    stageInfoListObj.Value = "HPLOFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventHPLActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����q����_�^�_�ĪG�ܧ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim tmpNum As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 48 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If (stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "PersonBloodControl") Or (stageInfoListObj.StageNum = 0 And stageInfoListObj.CommandStr = "@SystemHPLAction") Then
                    Select Case Val(commadstr3(0))
                        Case 1
                            tmpNum = stageInfoListObj.Argument + Val(commadstr3(1))
                        Case 2
                            tmpNum = stageInfoListObj.Argument - Val(commadstr3(1))
                        Case 3
                            tmpNum = Val(commadstr3(1))
                    End Select
                End If
                stageInfoListObj.Value = "HPLCHANGE%" + str(tmpNum)
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventHPLActionChange", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���a�Z������_�L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 47 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If (stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "BattleMoveControl") Or (stageInfoListObj.StageNum = 0 And stageInfoListObj.CommandStr = "@SystemBattleMove") Then
                    stageInfoListObj.Value = "BMCOFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventMoveActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H������_��_�L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 49 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "PersonResurrect" Then
                    stageInfoListObj.Value = "OFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventPersonResurrectActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���`���A����_�[�J(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim uscomt As Integer, k As Integer
    Dim vsstr As String
    Dim personStatus As clsStatus
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 4 Or atkingnum = 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 73 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If Val(commadstr3(4)) <= 0 Then GoTo vss_cmdlocalerr '==���O�ѼƦ^�X�Ƥ����T
            '==========================================
            If ((uscomt = 1 And liveus(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0) Or _
               (uscomt = 2 And livecom(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0)) Then
               GoTo VssCommadExit
            End If
            '===========================================������N�J�������`���A���
            If CollectionExists(�H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))), commadstr3(2)) = True Then
                Set personStatus = �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))))(commadstr3(2))
                personStatus.Value = Val(commadstr3(3))
                personStatus.Total = Val(commadstr3(4))
                '=======================
                vbecommadnum(2, vbecommadtotplayNow) = 2
                Exit Sub
            End If
            '===========================================�s�W���`���A���
            For k = 1 To UBound(VBEVSSBuffStr1)
                If VBEVSSBuffStr1(k) = commadstr3(2) Then
                    vsstr = FormMainMode.PEAFvssc(k + 54).Run("main", 4)
                    Set personStatus = New clsStatus
                    With personStatus
                        .Identifier = commadstr3(2)
                        .Value = Val(commadstr3(3))
                        .Total = Val(commadstr3(4))
                        .ImagePath = App.Path & vsstr
                    End With
                    �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))).Add personStatus, commadstr3(2)
                    '===================
                    vbecommadnum(2, vbecommadtotplayNow) = 2
                    Exit Sub
                End If
            Next
            '===============����첧�`���A���
            GoTo VssCommadExit
        Case 2
            Dim vbecommadnumSecond As Integer '���h���涥�q�s����
            '=======================
            vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
            '=======================
            Dim VBEStageNumMainSec(1 To 1) As Integer
            VBEStageNumMainSec(1) = Val(commadstr3(3))
            If CollectionExists(�H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))), commadstr3(2)) = True Then
                ���涥�q�t���`�D�n�{��_���`���A uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), commadstr3(2), 72, Val(commadstr3(1)), VBEStageNumMainSec, vbecommadnumSecond
            End If
            '=======================
            ���涥�q�t��_�ŧi�}�l�ε��� 2
            vbecommadnum(2, vbecommadtotplayNow) = 3
        Case 3
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 76
            VBEStageNum(1) = uscomt 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 1 '�[�J���A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            VBEStage7xAtkingInformation = commadstr3(2)
            '===========================���涥�q���J�I(76)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 76, 1
            '============================
            �԰��t����.���`���A��ܧ�s uscomt
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonAddBuff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_������Ʊ���(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (vbecommadnum(4, vbecommadtotplayNow) < 20 And vbecommadnum(4, vbecommadtotplayNow) > 29) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case Val(commadstr3(0))
                Case 1
                    �Y���淾�q�Ȯ��ܼ�(2) = �Y���淾�q�Ȯ��ܼ�(2) + Val(commadstr3(1))
                Case 2
                    �Y���淾�q�Ȯ��ܼ�(2) = �Y���淾�q�Ȯ��ܼ�(2) - Val(commadstr3(1))
                Case 3
                    �Y���淾�q�Ȯ��ܼ�(2) = Val(commadstr3(1))
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingTrueDiceControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���`���A����_��^�X����_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim buffvssnum As String
    Dim vsstr As String
    Dim personStatus As clsStatus
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 And atkingnum <> 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 73 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            buffvssnum = VBEVSSBuffStr1(vbecommadnum(3, vbecommadtotplayNow) - 54)
            VBEStage7xAtkingInformation = buffvssnum
            '===========================================������N�J�������`���A���
            If CollectionExists(�H�����`���A�C��(uscom, vbecommadnum(7, vbecommadtotplayNow)), buffvssnum) = True Then
                Set personStatus = �H�����`���A�C��(uscom, vbecommadnum(7, vbecommadtotplayNow))(buffvssnum)
                personStatus.Total = personStatus.Total - 1
                '=======================
                If personStatus.Total <= 0 Then
                    ���涥�q�t����.���涥�q73_���O_���`���A����_�D�ʲM�� uscom, vbecommadnum(6, vbecommadtotplayNow), buffvssnum
                    �H�����`���A�C��(uscom, vbecommadnum(7, vbecommadtotplayNow)).Remove buffvssnum
                    vbecommadnum(2, vbecommadtotplayNow) = 2
                    Exit Sub
                Else
                    �԰��t����.���`���A��ܧ�s uscom
                    GoTo VssCommadExit
                End If
            Else
                GoTo VssCommadExit
            End If
        Case 2
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 77
            VBEStageNum(1) = uscom 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 1 '�Ѱ����A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            '===========================���涥�q���J�I(77)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscom, 77, 1
            '============================
             �԰��t����.���`���A��ܧ�s uscom
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BuffTurnEnd", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���`���A����_�����M��_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim buffvssnum As String
    Dim vsstr As String
    Dim tempnum As Integer, uscomt As Integer, i As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or atkingnum = 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 73 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            '===========================================
            If ((uscomt = 1 And liveus(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0) Or _
               (uscomt = 2 And livecom(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0)) Then
               GoTo VssCommadExit
            End If
            If �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))).Count = 0 Then GoTo VssCommadExit
            '===========================================
            ���涥�q�t����.���涥�q73_���O_���`���A����_�����M�� uscomt, Val(commadstr3(1))
            tempnum = 1
            For i = 1 To �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))).Count
                If VBEStageRemoveBuffAllNum(i) = False Then
                    �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))).Remove tempnum
                Else
                    tempnum = tempnum + 1
                End If
            Next
            vbecommadnum(2, vbecommadtotplayNow) = 2
        Case 2
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 77
            VBEStageNum(1) = uscomt 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 1 '�Ѱ����A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            VBEStage7xAtkingInformation = ""
            '===========================���涥�q���J�I(77)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 77, 1
            '============================
            �԰��t����.���`���A��ܧ�s uscomt
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonRemoveBuffAll", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���`���A����_�S�w�M��_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim buffvssnum As String
    Dim vsstr As String
    Dim uscomt As Integer, tmpflag As Boolean
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or atkingnum = 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 73 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            '===========================================
            If ((uscomt = 1 And liveus(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0) Or _
               (uscomt = 2 And livecom(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0)) Then
               GoTo VssCommadExit
            End If
            '===========================================
            If CollectionExists(�H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))), commadstr3(2)) = True Then
                '===============================�[�J�Ӷ��q������T
                Dim stageInfoListObj As New clsVSStageObj
                stageInfoListObj.StageNum = vbecommadtotplayNow
                stageInfoListObj.CommandStr = "PersonRemoveBuffSelect"
                stageInfoListObj.Value = "0"
                ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
                '===============================
                tmpflag = False
                ���涥�q�t����.���涥�q73_���O_���`���A����_�S�w�M�� uscomt, Val(commadstr3(1)), commadstr3(2)
                If stageInfoListObj.CommandStr = "PersonRemoveBuffSelect" Then
                    If stageInfoListObj.Value = "OFF" Then
                        tmpflag = True
                    End If
                End If
                If tmpflag = False Then
                   �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))).Remove commadstr3(2)
                End If
                ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
                '===============================
                If tmpflag = True Then GoTo VssCommadExit
                '===============================
                VBEStage7xAtkingInformation = commadstr3(2)
                vbecommadnum(2, vbecommadtotplayNow) = 2
                Exit Sub
            Else
                GoTo VssCommadExit '�����Ӳ��`���A
            End If
        Case 2
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 77
            VBEStageNum(1) = uscomt 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 1 '�Ѱ����A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            '===========================���涥�q���J�I(77)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 77, 1
            '============================
            �԰��t����.���`���A��ܧ�s uscomt
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonRemoveBuffSelect", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���`���A����_�ܧ�^�X��(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim buffvssnum As String
    Dim vsstr As String
    Dim uscomt As Integer
    Dim personStatus As clsStatus
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 4 Or atkingnum = 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 73 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            '===========================================
            If ((uscomt = 1 And liveus(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0) Or _
               (uscomt = 2 And livecom(����ݾ��H��������(uscomt, Val(commadstr3(1)))) <= 0)) Then
               GoTo VssCommadExit
            End If
            '===========================================
            If CollectionExists(�H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))), commadstr3(2)) = True Then
                Set personStatus = �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))))(commadstr3(2))
                Select Case Val(commadstr3(3))
                    Case 1
                       personStatus.Total = personStatus.Total + Val(commadstr3(4))
                    Case 2
                       personStatus.Total = personStatus.Total - Val(commadstr3(4))
                    Case 3
                       personStatus.Total = Val(commadstr3(4))
                End Select
                '=======================
                If personStatus.Total <= 0 Then
                    ���涥�q�t����.���涥�q73_���O_���`���A����_�D�ʲM�� uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), commadstr3(2)
                    �H�����`���A�C��(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1)))).Remove commadstr3(2)
                    VBEStage7xAtkingInformation = commadstr3(2)
                    vbecommadnum(2, vbecommadtotplayNow) = 2
                    Exit Sub
                Else
                    �԰��t����.���`���A��ܧ�s uscom
                    GoTo VssCommadExit
                End If
            Else
                GoTo VssCommadExit '�����Ӳ��`���A
            End If
        Case 2
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 77
            VBEStageNum(1) = uscomt 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 1 '�Ѱ����A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            '===========================���涥�q���J�I(77)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 77, 1
            '============================
            �԰��t����.���`���A��ܧ�s uscomt
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonBuffTurnChange", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub

Sub ������O_���椧���`���A�����L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 73 Or atkingnum <> 9 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And (stageInfoListObj.CommandStr = "PersonRemoveBuffSelect" Or stageInfoListObj.CommandStr = "PersonRemoveBuffAll" Or stageInfoListObj.CommandStr = "@HPWEvent") Then
                    stageInfoListObj.Value = "OFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventRemoveBuffActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�֦����d�P����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim tmpcard As clsActionCard, tmpcollectionIndex As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or commadtype <> 1 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
            Select Case vbecommadnum(4, vbecommadtotplayNow)
                Case 2, 3, 4, 70, 10, 11, 12, 17, 30, 31, 32, 37
                Case Else
                    GoTo VssCommadExit
            End Select
        Case Else
            GoTo VssCommadExit
    End Select
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            tmpcollectionIndex = �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(Val(commadstr3(2)))
            Select Case uscomt
                Case 1
                     Select Case Val(commadstr3(1))
                         Case 1 '==��P�X�P
                            If tmpcollectionIndex = 5 Then
                                FormMainMode.card_CardClick �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            End If
                         Case 2 '==�X�P�^�P
                            If tmpcollectionIndex = 6 Then
                                FormMainMode.card_CardClick �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            End If
                         Case 3 '==��P
                            If tmpcollectionIndex = 5 Then
                                FormMainMode.card_CardButtonClickin �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            ElseIf tmpcollectionIndex = 6 Then
                                FormMainMode.card_CardButtonClickout �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            End If
                    End Select
                Case 2
                    Select Case Val(commadstr3(1))
                         Case 1 '==��P�X�P
                            If tmpcollectionIndex = 7 Then
                                �԰��t����.�q���P_�������P �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            End If
                         Case 2 '==�X�P�^�P
                            If tmpcollectionIndex = 8 Then
                                �԰��t����.�q���P_�������P_�~ �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            End If
                         Case 3 '==��P
                            If tmpcollectionIndex = 7 Then
                                Set tmpcard = �԰��t����.CardDeckCollection(7)(CStr(�԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))))
                                Call tmpcard.Reverse
                            ElseIf tmpcollectionIndex = 8 Then
                                �԰��t����.�q���P_������P_�~ �԰��t����.�d�P�P�ﶰ�X����_CardNum(Val(commadstr3(2)))
                            End If
                    End Select
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingOneSelfCardControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�����Y��l(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or (vbecommadnum(4, vbecommadtotplayNow) <> 13 And vbecommadnum(4, vbecommadtotplayNow) <> 33 And vbecommadnum(4, vbecommadtotplayNow) < 20 And vbecommadnum(4, vbecommadtotplayNow) > 29) Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Erase Vss_BattleStartDiceNum
            �Y���淾�q�Ȯ��ܼ�(9) = �������m��l�`��(1)
            �Y���淾�q�Ȯ��ܼ�(10) = �������m��l�`��(2)
            Vss_BattleStartDiceNum(0) = 62
            Vss_BattleStartDiceNum(1) = �������m��l�`��(1)
            Vss_BattleStartDiceNum(2) = �������m��l�`��(2)
            �O�_�t�Τ��� = False
            �԰��t����.�Y�������
            ���ݮɶ���C(2).Add 24
            FormMainMode.���ݮɶ�_2.Enabled = True
            vbecommadnum(2, vbecommadtotplayNow) = 0 '==���ݮɶ�
        Case 2
            Dim vbecommadnumSecond As Integer '���h���涥�q�s����
            '=======================
            vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
            '=======================
            �԰��t����.�Y�����P�_
            Dim buffvssnum As String
            If vbecommadnum(3, vbecommadtotplayNow) <= 24 Then
                ���涥�q�t����.���涥�q�t���`�D�n�{��_�H���D�ʧޯ� uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 62, vbecommadnum(6, vbecommadtotplayNow), Vss_BattleStartDiceNum, vbecommadnumSecond
            ElseIf vbecommadnum(3, vbecommadtotplayNow) > 24 And vbecommadnum(3, vbecommadtotplayNow) <= 48 Then
                ���涥�q�t����.���涥�q�t���`�D�n�{��_�H���Q�ʧޯ� uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 62, vbecommadnum(6, vbecommadtotplayNow), Vss_BattleStartDiceNum, vbecommadnumSecond
            ElseIf vbecommadnum(3, vbecommadtotplayNow) > 48 And vbecommadnum(3, vbecommadtotplayNow) <= 54 Then
                ���涥�q�t����.���涥�q�t���`�D�n�{��_�H����ڪ��A uscom, vbecommadnum(7, vbecommadtotplayNow), 62, vbecommadnum(6, vbecommadtotplayNow), Vss_BattleStartDiceNum, vbecommadnumSecond
            Else
                buffvssnum = VBEVSSBuffStr1(vbecommadnum(3, vbecommadtotplayNow) - 54)
                If CollectionExists(�H�����`���A�C��(uscom, vbecommadnum(7, vbecommadtotplayNow)), buffvssnum) = True Then
                    ���涥�q�t����.���涥�q�t���`�D�n�{��_���`���A uscom, vbecommadnum(7, vbecommadtotplayNow), buffvssnum, 62, vbecommadnum(6, vbecommadtotplayNow), Vss_BattleStartDiceNum, vbecommadnumSecond
                End If
            End If
            '=======================
            ���涥�q�t��_�ŧi�}�l�ε��� 2
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BattleStartDice", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H���̤j�d��Ʊ���(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case uscomt
                Case 1
                    Select Case Val(commadstr3(1))
                        Case 1
                            �P�`���q��(1) = �P�`���q��(1) + Val(commadstr3(2))
                        Case 2
                            �P�`���q��(1) = �P�`���q��(1) - Val(commadstr3(2))
                    End Select
                Case 2
                    Select Case Val(commadstr3(1))
                        Case 1
                            �P�`���q��(2) = �P�`���q��(2) + Val(commadstr3(2))
                        Case 2
                            �P�`���q��(2) = �P�`���q��(2) - Val(commadstr3(2))
                    End Select
            End Select
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonMaxCardsNumControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���J�ƥ�d(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim i As Integer
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (commadtype <> 1 And commadtype <> 3) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If Val(commadstr3(1)) <= �԰��t����.BattleTurn Or �@��t����.�ƥ�d��Ʈw(commadstr3(2), 1) = 99 Then
                GoTo VssCommadExit
            End If
            
            �԰��t����.�o��d�P_�ƥ�d uscomt, commadstr3(2), �@��t����.�ƥ�d��Ʈw(commadstr3(2), 2), Val(commadstr3(1)) - �԰��t����.BattleTurn
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "BattleInsertEventCard", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����ڪ��A����_�[�J(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim uscomt As Integer, k As Integer
    Dim vsstr As String, textlinea As String, str As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 3 Or atkingnum >= 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 75 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    '=========================
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If Val(commadstr3(3)) <= 0 Then GoTo vss_cmdlocalerr '==���O�ѼƦ^�X�Ƥ����T
            '===========================================�M�ŬJ�����H����ڪ��A���
            If �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 1) <> "" Then
                For k = 1 To 9
                     �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), k) = ""
                Next
            End If
            '===========================================�s�W�H����ڪ��A���
            For k = 1 To UBound(VBEVSSActualStatusStr1)
                If VBEVSSActualStatusStr1(k) = commadstr3(2) Then
                    Open VBEVSSActualStatusStr2(k) For Input As #1
                    Do Until EOF(1)
                       Line Input #1, textlinea
                       str = str & textlinea & vbCrLf
                    Loop
                    Close
                    If str <> "" Then
                        FormMainMode.PEAFvssc((uscomt - 1) * 3 + ����ݾ��H��������(uscomt, Val(commadstr3(1))) + 48).AddCode str
                        If �@��t����.ProgramIsOnWine = True Then ���涥�q�t����.���涥�q�t��_�[�JWine�{���i�J�I (uscomt - 1) * 3 + ����ݾ��H��������(uscomt, Val(commadstr3(1))) + 48
                    End If
                    vsstr = FormMainMode.PEAFvssc((uscomt - 1) * 3 + ����ݾ��H��������(uscomt, Val(commadstr3(1))) + 48).Run("main", 1)
                    If vsstr = commadstr3(2) Then
                        �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 1) = commadstr3(2)
                        �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 9) = FormMainMode.PEAFvssc((uscomt - 1) * 3 + ����ݾ��H��������(uscomt, Val(commadstr3(1))) + 48).Run("main", 3)
                        vbecommadnum(2, vbecommadtotplayNow) = 2
                        Exit Sub
                    End If
                End If
            Next
            '===============�����ŦX���H����ڪ��A�}�����
            GoTo VssCommadExit
        Case 2
            Dim vbecommadnumSecond As Integer '���h���涥�q�s����
            '=======================
            vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
            '=======================
            Dim VBEStageNumMainSec(1 To 1) As Integer
            ���涥�q�t����.���涥�q�t���`�D�n�{��_�H����ڪ��A uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 74, Val(commadstr3(1)), VBEStageNumMainSec, vbecommadnumSecond
            '=======================
            ���涥�q�t��_�ŧi�}�l�ε��� 2
            vbecommadnum(2, vbecommadtotplayNow) = 3
        Case 3
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 76
            VBEStageNum(1) = uscomt 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 2 '�[�J���A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            VBEStage7xAtkingInformation = commadstr3(2)
            '===========================���涥�q���J�I(76)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 76, 1
            '============================
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonAddActualStatus", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����ڪ��A�[�J���_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim personnum As Integer, i As Integer, p As Integer
    Dim strfalse As Boolean
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 7 And atkingnum <> 10 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) = 72 Or _
                vbecommadnum(4, vbecommadtotplayNow) = 73 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    personnum = vbecommadnum(7, vbecommadtotplayNow)
    '==========
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            For i = 3 To 6
                If commadstr3(i - 3) = "" Then
                    strfalse = True
                Else
                    �H����ڪ��A��Ʈw(uscom, personnum, i) = App.Path & commadstr3(i - 3)
                End If
            Next
            p = (uscom - 1) * 2 + 4
            For i = 7 To 8
                 �H����ڪ��A��Ʈw(uscom, personnum, i) = Val(commadstr3(p))
                 p = p + 1
            Next
            If strfalse = False Then �H����ڪ��A��Ʈw(uscom, personnum, 2) = 1 Else �H����ڪ��A��Ʈw(uscom, personnum, 2) = 0
            '===================
            If ����H����ԤH��(uscom, 2) = personnum And �H����ڪ��A��Ʈw(uscom, personnum, 2) = 1 Then
                Select Case uscom
                    Case 1
                        FormMainMode.personusminijpg.�p�H������ = True
                    Case 2
                        FormMainMode.personcomminijpg.�p�H������ = True
                End Select
                vbecommadnum(2, vbecommadtotplayNow) = 2
            Else
                GoTo VssCommadExit
            End If
            '===================
        Case 2
            If FormMainMode.personusminijpg.�p�H������ = False And FormMainMode.personcomminijpg.�p�H������ = False Then
                '==================
                Select Case uscom
                    Case 1
                        FormMainMode.personusminijpg.�p�H���Ϥ� = �H����ڪ��A��Ʈw(uscom, personnum, 4)
                        FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = �H����ڪ��A��Ʈw(uscom, personnum, 5)
                        FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = �H����ڪ��A��Ʈw(uscom, personnum, 6)
                        FormMainMode.personusminijpg.�p�H���v�lLeft = Val(�H����ڪ��A��Ʈw(uscom, personnum, 7))
                        FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(�H����ڪ��A��Ʈw(uscom, personnum, 8))
                        �԰��Y�뤶���H����ø�ϸ��|������(1) = �H����ڪ��A��Ʈw(uscom, personnum, 3)
                        FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -(FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width)
                        �԰��t����.����ʧ@_�Z���ܧ� movecp, False, True
                        FormMainMode.personusminijpg.�p�H����{ = True
                    Case 2
                        FormMainMode.personcomminijpg.�p�H���Ϥ� = �H����ڪ��A��Ʈw(uscom, personnum, 4)
                        FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = �H����ڪ��A��Ʈw(uscom, personnum, 5)
                        FormMainMode.��ܦC1.�q����p�H���Ϥ� = �H����ڪ��A��Ʈw(uscom, personnum, 6)
                        FormMainMode.personcomminijpg.�p�H���v�lLeft = Val(�H����ڪ��A��Ʈw(uscom, personnum, 7))
                        FormMainMode.personcomminijpg.�p�H���v�ltop�t = Val(�H����ڪ��A��Ʈw(uscom, personnum, 8))
                        �԰��Y�뤶���H����ø�ϸ��|������(2) = �H����ڪ��A��Ʈw(uscom, personnum, 3)
                        FormMainMode.��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
                        �԰��t����.����ʧ@_�Z���ܧ� movecp, False, True
                        FormMainMode.personcomminijpg.�p�H����{ = True
                End Select
                vbecommadnum(2, vbecommadtotplayNow) = 3
                '==================
            End If
        Case 3
            If FormMainMode.personusminijpg.�p�H����{ = False And FormMainMode.personcomminijpg.�p�H����{ = False Then
                GoTo VssCommadExit
            End If
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventAddActualStatusData", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����ڪ��A����_�ŧi����_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim personnum As Integer, i As Integer
    Dim vsstr As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 And atkingnum <> 10 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 75 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    personnum = vbecommadnum(7, vbecommadtotplayNow)
    '===========
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Dim vbecommadnumSecond As Integer '���h���涥�q�s����
            '=======================
            vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
            '=======================
            Dim VBEStageNumMainSec(0 To 1) As Integer
            VBEStageNumMainSec(0) = 75
            VBEStageNumMainSec(1) = 0
            ���涥�q�t����.���涥�q�t���`�D�n�{��_�H����ڪ��A uscom, vbecommadnum(7, vbecommadtotplayNow), 75, vbecommadnum(6, vbecommadtotplayNow), VBEStageNumMainSec, vbecommadnumSecond
            '=======================
            ���涥�q�t��_�ŧi�}�l�ε��� 2
            '=======================
            If ����H����ԤH��(uscom, 2) = personnum Then
                Select Case uscom
                    Case 1
                        FormMainMode.personusminijpg.�p�H������ = True
                    Case 2
                        FormMainMode.personcomminijpg.�p�H������ = True
                End Select
                vbecommadnum(2, vbecommadtotplayNow) = 2
            Else
                vbecommadnum(2, vbecommadtotplayNow) = 3
            End If
        Case 2
            If FormMainMode.personusminijpg.�p�H������ = False And FormMainMode.personcomminijpg.�p�H������ = False Then
                '==================
                Select Case uscom
                    Case 1
                        FormMainMode.personusminijpg.�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 1)
                        FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 2)
                        FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 4)
                        FormMainMode.personusminijpg.�p�H���v�lLeft = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 5))
                        FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 6))
                        �԰��Y�뤶���H����ø�ϸ��|������(1) = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 3)
                        FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -(FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width)
                        �԰��t����.����ʧ@_�Z���ܧ� movecp, False, True
                        FormMainMode.personusminijpg.�p�H����{ = True
                    Case 2
                        FormMainMode.personcomminijpg.�p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 1)
                        FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 2)
                        FormMainMode.��ܦC1.�q����p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 4)
                        FormMainMode.personcomminijpg.�p�H���v�lLeft = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 5)
                        FormMainMode.personcomminijpg.�p�H���v�ltop�t = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 6)
                        �԰��Y�뤶���H����ø�ϸ��|������(2) = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 3)
                        FormMainMode.��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
                        �԰��t����.����ʧ@_�Z���ܧ� movecp, False, True
                        FormMainMode.personcomminijpg.�p�H����{ = True
                End Select
                VBEStage7xAtkingInformation = �H����ڪ��A��Ʈw(uscom, personnum, 1)
                vbecommadnum(2, vbecommadtotplayNow) = 3
                '==================
            End If
        Case 3
            If FormMainMode.personusminijpg.�p�H����{ = False And FormMainMode.personcomminijpg.�p�H����{ = False Then
                For i = 1 To UBound(�H����ڪ��A��Ʈw, 3)
                     �H����ڪ��A��Ʈw(uscom, personnum, i) = ""
                Next
                FormMainMode.PEAFvssc(vbecommadnum(3, vbecommadtotplayNow)).Reset
                vbecommadnum(2, vbecommadtotplayNow) = 4
            End If
        Case 4
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 77
            VBEStageNum(1) = uscom 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 2 '�Ѱ����A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            '===========================���涥�q���J�I(77)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscom, 77, 1
            '============================
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "ActualStatusEnd", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H����ڪ��A����_�S�w�Ѱ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    Dim buffvssnum As String
    Dim vsstr As String
    Dim uscomt As Integer, i As Integer, tmpflag As Boolean
    Dim stageInfoListObj As clsVSStageObj
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or atkingnum >= 9 Then GoTo VssCommadExit
    Select Case commadtype
        Case 1
        Case 3
            If vbecommadnum(4, vbecommadtotplayNow) >= 72 And _
                vbecommadnum(4, vbecommadtotplayNow) <= 75 Then GoTo VssCommadExit
        Case Else
            GoTo VssCommadExit
    End Select
    '=======================
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 1) <> "" Then
                '===============================�[�J�Ӷ��q������T
                Set stageInfoListObj = New clsVSStageObj
                stageInfoListObj.StageNum = vbecommadtotplayNow
                stageInfoListObj.CommandStr = "PersonRemoveActualStatus"
                stageInfoListObj.Value = "0"
                ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
                '===============================
                Dim vbecommadnumSecond As Integer '���h���涥�q�s����
                vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
                '=======================
                Dim VBEStageNumMainSec(0 To 1) As Integer
                VBEStageNumMainSec(0) = 75
                VBEStageNumMainSec(1) = 1
                ���涥�q�t����.���涥�q�t���`�D�n�{��_�H����ڪ��A uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 75, Val(commadstr3(1)), VBEStageNumMainSec, vbecommadnumSecond
                '=======================
                ���涥�q�t��_�ŧi�}�l�ε��� 2
                '=======================
                vbecommadnum(2, vbecommadtotplayNow) = 2
            Else
                GoTo VssCommadExit
            End If
            '=================
        Case 2
            Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
            tmpflag = False
            If stageInfoListObj.CommandStr = "PersonRemoveActualStatus" Then
                If stageInfoListObj.Value = "OFF" Then
                    tmpflag = True
                End If
            End If
            
            ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
            
            If tmpflag = False Then
               If Val(commadstr3(1)) = 1 Then
                    Select Case uscomt
                        Case 1
                            FormMainMode.personusminijpg.�p�H������ = True
                        Case 2
                            FormMainMode.personcomminijpg.�p�H������ = True
                    End Select
                    vbecommadnum(2, vbecommadtotplayNow) = 3
                Else
                    vbecommadnum(2, vbecommadtotplayNow) = 4
                End If
            Else
                GoTo VssCommadExit
            End If
        Case 3
            If FormMainMode.personusminijpg.�p�H������ = False And FormMainMode.personcomminijpg.�p�H������ = False Then
                '==================
                Select Case uscomt
                    Case 1
                        FormMainMode.personusminijpg.�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 1)
                        FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 2)
                        FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 4)
                        FormMainMode.personusminijpg.�p�H���v�lLeft = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 5))
                        FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 6))
                        �԰��Y�뤶���H����ø�ϸ��|������(1) = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 3)
                        FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -(FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width)
                        �԰��t����.����ʧ@_�Z���ܧ� movecp, False, True
                        FormMainMode.personusminijpg.�p�H����{ = True
                    Case 2
                        FormMainMode.personcomminijpg.�p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 1)
                        FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 2)
                        FormMainMode.��ܦC1.�q����p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 4)
                        FormMainMode.personcomminijpg.�p�H���v�lLeft = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 5)
                        FormMainMode.personcomminijpg.�p�H���v�ltop�t = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 6)
                        �԰��Y�뤶���H����ø�ϸ��|������(2) = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 3)
                        FormMainMode.��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
                        �԰��t����.����ʧ@_�Z���ܧ� movecp, False, True
                        FormMainMode.personcomminijpg.�p�H����{ = True
                End Select
                VBEStage7xAtkingInformation = �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), 1)
                vbecommadnum(2, vbecommadtotplayNow) = 4
                '==================
            End If
        Case 4
            If FormMainMode.personusminijpg.�p�H����{ = False And FormMainMode.personcomminijpg.�p�H����{ = False Then
                For i = 1 To UBound(�H����ڪ��A��Ʈw, 3)
                     �H����ڪ��A��Ʈw(uscomt, ����ݾ��H��������(uscomt, Val(commadstr3(1))), i) = ""
                Next
                FormMainMode.PEAFvssc((uscomt - 1) * 3 + ����ݾ��H��������(uscomt, Val(commadstr3(1))) + 48).Reset
                vbecommadnum(2, vbecommadtotplayNow) = 5
            End If
        Case 5
            ReDim VBEStageNum(0 To 3) As Integer
            VBEStageNum(0) = 77
            VBEStageNum(1) = uscomt 'Ĳ�o�ƥ��(1.�ϥΪ�/2.�q��)
            VBEStageNum(2) = 2 '�Ѱ����A���O(1.���`���A/2.�H����ڪ��A)
            VBEStageNum(3) = 0 '�ޯ�ߤ@�ѧO�X�\���
            '===========================���涥�q���J�I(77)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomt, 77, 1
            '============================
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonRemoveActualStatus", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���椧�H����ڪ��A�����L�Ĥ�_�M(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or Val(vbecommadnum(4, vbecommadtotplayNow)) <> 75 Or atkingnum <> 10 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If ���涥�q�t����.VBEVSStageInfoList.Count > 0 Then
                Dim stageInfoListObj As clsVSStageObj
                Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
                If stageInfoListObj.StageNum = vbecommadtotplayNow - 1 And stageInfoListObj.CommandStr = "PersonRemoveActualStatus" Then
                    stageInfoListObj.Value = "OFF"
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventRemoveActualStatusActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�T��a�i��Ҧ��ާ@(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or (Val(vbecommadnum(4, vbecommadtotplayNow)) <> 1 And Val(vbecommadnum(4, vbecommadtotplayNow)) <> 17 And Val(vbecommadnum(4, vbecommadtotplayNow)) <> 37) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Vss_EventPlayerAllActionOffNum(uscomt) = 1
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventPlayerAllActionOff", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H�����Ⲿ�ʶ��q��ʱ���(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or (vbecommadnum(4, vbecommadtotplayNow) <> 2 And vbecommadnum(4, vbecommadtotplayNow) <> 3 And vbecommadnum(4, vbecommadtotplayNow) <> 4 And vbecommadnum(4, vbecommadtotplayNow) <> 70) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            If Val(commadstr3(1)) >= 1 And Val(commadstr3(1)) <= 5 Then
                If ����H����ԤH��(uscomt, 1) = 1 And Val(commadstr3(1)) = 4 Then
                    Vss_PersonMoveActionChangeNum(uscomt, 1) = 0
                Else
                    Vss_PersonMoveActionChangeNum(uscomt, 1) = 1
                End If
                If Val(commadstr3(1)) = 5 Then
                    Vss_PersonMoveActionChangeNum(uscomt, 2) = 0
                Else
                    Vss_PersonMoveActionChangeNum(uscomt, 2) = Val(commadstr3(1))
                End If
            End If
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonMoveActionChange", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_���ʫe�`���ʶq����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or (vbecommadnum(4, vbecommadtotplayNow) <> 2 And vbecommadnum(4, vbecommadtotplayNow) <> 3 And vbecommadnum(4, vbecommadtotplayNow) <> 4 And vbecommadnum(4, vbecommadtotplayNow) <> 70) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case commadstr3(1)
                Case 1
                     If Vss_PersonMoveControlNum(uscomt, 2) = 0 Then
                        Vss_PersonMoveControlNum(uscomt, 1) = Vss_PersonMoveControlNum(uscomt, 1) + Val(commadstr3(2))
                     End If
                Case 2
                     If Vss_PersonMoveControlNum(uscomt, 2) = 0 Then
                        Vss_PersonMoveControlNum(uscomt, 1) = Vss_PersonMoveControlNum(uscomt, 1) - Val(commadstr3(2))
                     End If
                Case 3
                     Vss_PersonMoveControlNum(uscomt, 1) = Val(commadstr3(2))
                     Vss_PersonMoveControlNum(uscomt, 2) = 1
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonMoveControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�H�������u����������(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or (vbecommadnum(4, vbecommadtotplayNow) <> 2 And vbecommadnum(4, vbecommadtotplayNow) <> 3 And vbecommadnum(4, vbecommadtotplayNow) <> 4 And vbecommadnum(4, vbecommadtotplayNow) <> 70 And vbecommadnum(4, vbecommadtotplayNow) <> 71) Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Vss_PersonAttackFirstControlNum = uscomt
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonAttackFirstControl", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ޯ���O�Ƶ��r��(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 0 Or vbecommadnum(3, vbecommadtotplayNow) > 48 Then GoTo VssCommadExit
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case vbecommadnum(3, vbecommadtotplayNow)
                Case Is <= 24 '==�D�ʧ�
                        Vss_AtkingInformationRecordStr(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum) = commadstr3(0)
                Case Is <= 48 '==�Q�ʧ�
                        Vss_AtkingInformationRecordStr(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum) = commadstr3(0)
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "AtkingInformationRecord", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub

Sub ������O_���O�����аO(ByVal vbecommadtotplayNow As Integer)
    vbecommadnum(1, vbecommadtotplayNow) = vbecommadnum(1, vbecommadtotplayNow) + 1
End Sub
Sub ������O��_���~�T���q��(ByVal name As String, ByVal cmdturn As Integer, ByVal systurn As Integer)
MsgBox "���涥�q���~(04-" & systurn & "-" & name & "-" & cmdturn & ")�G" & Chr(10) & "���O�����ɵo�Ϳ��~�C" & Chr(10) & Chr(10) & "(" & Err.Number & "):" & Err.Description, vbCritical
End
End Sub
Function ������O��_��������(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer) As Boolean
If vbecommadnum(3, vbecommadtotplayNow) <= 48 Then  '==�ȥD�ʧޯ�/�Q�ʧޯ�ݶi��Ұ�����
    Select Case vbecommadnum(3, vbecommadtotplayNow)
         Case Is <= 24
             If atkingck(uscom, ����H����ԤH��(uscom, 2), atkingnum, 1) = 1 Then
                 ������O��_�������� = True
             Else
                 ������O��_�������� = False
             End If
         Case Is <= 48
             If atkingck(uscom, vbecommadnum(7, vbecommadtotplayNow), atkingnum, 1) = 1 Then
                 ������O��_�������� = True
             Else
                 ������O��_�������� = False
             End If
    End Select
Else
    ������O��_�������� = True
End If
End Function
Sub ������O_�H������խȹ����ܤƶq����(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 2 Or vbecommadnum(4, vbecommadtotplayNow) <> 45 Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            Select Case commadstr3(1)
                Case 1
                     If Vss_EventPersonAbilityDiceChangeNum(uscomt, 2) = 0 Then
                        Vss_EventPersonAbilityDiceChangeNum(uscomt, 1) = Vss_EventPersonAbilityDiceChangeNum(uscomt, 1) + Val(commadstr3(2))
                     End If
                Case 2
                     If Vss_EventPersonAbilityDiceChangeNum(uscomt, 2) = 0 Then
                        Vss_EventPersonAbilityDiceChangeNum(uscomt, 1) = Vss_EventPersonAbilityDiceChangeNum(uscomt, 1) - Val(commadstr3(2))
                     End If
                Case 3
                     Vss_EventPersonAbilityDiceChangeNum(uscomt, 1) = Val(commadstr3(2))
                     Vss_EventPersonAbilityDiceChangeNum(uscomt, 2) = 1
            End Select
            GoTo VssCommadExit
    End Select
        '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "EventPersonAbilityDiceChange", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
Sub ������O_�ܧ�H���԰���ø(ByVal uscom As Integer, ByVal commadtype As Integer, ByVal atkingnum As Integer, ByVal vbecommadtotplayNow As Integer)
    If Formsetting.checktest.Value = 0 Then On Error GoTo vss_cmdlocalerr
    Dim commadstr3() As String
    
    commadstr3 = Split(vbecommadstr(3, vbecommadtotplayNow), ",")
    If UBound(commadstr3) <> 1 Or atkingnum > 8 Or commadtype = 2 Then GoTo VssCommadExit
    Dim uscomt As Integer
    Select Case Val(commadstr3(0))
         Case 1
               uscomt = uscom
         Case 2
               If uscom = 1 Then uscomt = 2 Else uscomt = 1
    End Select
    Select Case vbecommadnum(2, vbecommadtotplayNow)
        Case 1
            �԰��Y�뤶���H����ø�ϸ��|������(uscomt) = App.Path & commadstr3(1)
            GoTo VssCommadExit
    End Select
    '============================
    Exit Sub
VssCommadExit:
    ������O��.������O_���O�����аO vbecommadtotplayNow
    '============================
'=============================
Exit Sub
vss_cmdlocalerr:
������O��.������O��_���~�T���q�� "PersonChangeBattleImage", vbecommadnum(2, vbecommadtotplayNow), vbecommadnum(4, vbecommadtotplayNow)
End Sub
