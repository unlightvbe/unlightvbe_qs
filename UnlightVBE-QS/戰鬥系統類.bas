Attribute VB_Name = "�԰��t����"
Option Explicit
Public Const a1a As String = "ATK-�C"
Public Const a2a As String = "DEF"
Public Const a3a As String = "MOV"
Public Const a4a As String = "SPE"
Public Const a5a As String = "ATK-�j"
Public Const a6a As String = "DRAW"
Public Const a7a As String = "BRK"
Public Const a8a As String = "HPL"
Public Const a9a As String = "HPW"
Public Const b1b As Integer = 1
Public Const b2b As Integer = 2
Public Const b3b As Integer = 3
Public Const b4b As Integer = 4
Public Const b5b As Integer = 5
Public Const b6b As Integer = 6
Public Const b7b As Integer = 7
Public Const b8b As Integer = 8
Public Const b9b As Integer = 9

Public goicheck(1 To 2) As Integer   '����/���m�Ҧ��[��ƭ��ˬd�X
Public pageonin(1 To 999) As Integer  '�P�i���ϭ��ˬd�X
Public liveus(1 To 3) As Integer, livecom(1 To 3) As Integer, liveusmax(1 To 3) As Integer, livecommax(1 To 3) As Integer
Public BattleTurn As Integer, BattleCardNum As Integer, atkus(1 To 3) As Integer, atkcom(1 To 3) As Integer, defus(1 To 3) As Integer, defcom(1 To 3) As Integer, pagecheckus As Integer, pagecheckcom As Integer, pagegive As Integer, goidefus As Integer, movecom As Integer, moveus As Integer, movecp As Integer, chkcomck As Integer, uslevel(1 To 3) As Integer, comlevel(1 To 3) As Integer, liveus41(1 To 3) As Integer, livecom41(1 To 3) As Integer, movecheckcom As Integer, movecheckus As Integer
Public nameus(1 To 3) As String, namecom(1 To 3) As String
Public moveturn As Integer  '���������m�Ҧ������ˬd�X(1.�ϥΪ̥���/2.�q������)
Public atkinghelpxy(1 To 2, 1 To 4, 1 To 2) As Integer '�ޯ໡����y�Ы��w���(1.�ϥΪ̤�/2.�q����,��1~4�ӧޯ�,1.Left/2.Top(�y��))
Public pageusleadmax(0 To 1) As Integer   '�ϥΪ̵P���ǭp�ƪ�(0.��P/1.�X�P)
Public pagecomleadmax(0 To 1) As Integer   '�q���P���ǭp�ƪ�(0.��P/1.�X�P)
Public pageqlead(1 To 2) As Integer   '�X�P�p���ܼ�(1.�ϥΪ�/2.�q��)
Public pageglead(1 To 2) As Integer   '��P�p���ܼ�(1.�ϥΪ�/2.�q��)
Public movedsus As Integer   '�ϥΪ̲��ʶ��q�M�w���ܼ�
Public turnpageonin As Integer  '���q�O�_�i�X�P�ܼ�(�@��)
Public turnpageoninatking As Integer  '���q�O�_�i�X�P�ܼ�(�ޯ�ϥ�)
Public goickus As Integer '�P�Ȥ@���ˬd�X
Public atkingck(1 To 2, 1 To 3, 1 To 8, 1 To 3) As Integer '�ޯඥ�q�ҰʽX(1.�ϥΪ�/2.�q��,1~3.�H���s��/1~4�H���ۨ��ޯඵ��;5~8�H���ۨ��Q�ʧ޶���,1.�ޯ�ҰʼаO/2.�o�^�X���Ұʦ���(�D�ʧ�->�ʵe����)/3.�o���԰����Ұʦ���(�D�ʧ�->�ʵe����))
Public atkingckdice(1 To 2, 1 To 2, 1 To 4) As String '�H���ޯ��l�v�T�����Ȯ��ܼ�(1.�ϥΪ�/2.�q��,1.��ϥΪ�/2.��q��,1.�D�ʧ�/2.�Q�ʧ�/3.���`���A/4.�H����ڪ��A,���`��Ƥ��v�T�q�ܤƦ�)
Public atkingtrn(1 To 4) As Integer '�ޯ�p�ƾ��Ȯ��x�s�ܼ�(1.�ϥΪ�(�{)/2.�q��(�{)/3.�ϥΪ�(�ƥ�)/4.�q��(�ƥ�))
Public akhpnm As Integer  '�ޯ໡���Ȯ��ܼ�
Public turnatk As Integer  '���������m���q�ܼ�(1.�ϥΪ̧����B�q�����m,2.�ϥΪ̨��m�B�q������,3.�o�P�B����)
Public trend�Ȯ��ܼ� As Integer '�������q�p�ƾ��Ȯ��ܼ�
Public HP�ˬd�ܼ� As Boolean 'HP�ˬd���q�O�_�w�ˬd�ܼ�
Public HP�ˬd���q�� As Integer 'HP�ˬd���q�ܼ�(1.���ʶ��q��,2.����/���m���q�e,3.��/���m���q��)
Public �Z�����(1 To 2, 1 To 2, 1 To 2) As Integer  '�Z�����Ȯ��x�s���(1.HP���/2.�P����,1.�ϥΪ�/2.�q��,1.Left���/2.Top���)
Public personminixy(1 To 2, 1 To 3, 1 To 3, 1 To 2) As Integer '�p�H���Ϥ��y�Ы��w���(1.�ϥΪ�/2.�q��,��n��,1.��Z��/2.���Z��/3.���Z��,1.Left/2.Top(�y��))
Public ���`���A�ˬd��(1 To 40, 1 To 2) As Integer '���`���A�ҰʽX(x.���`���A�s��,1.���A���涥�q/2.���A�Ұ��ˬd��)
Public �ޯ�ʵe��ܶ��q�� As Integer '�ޯ�ʵe�p�ƾ����q�X(1.����/���m���q-���q,2.���ʶ��q-���q/3.�o�P���q��B���ʶ��q�e/4.���ʶ��q��/5.�������q��/6.���m���q��/7.�^�X������)
Public �������m��l�`��(1 To 4) As Integer '����/���m�Ҧ���l�ƶq���(1.�ϥΪ�(�`)/2.�q��(�`)/3.�ϥΪ�(��)/4.�q��(��))
Public atkingpagetot(1 To 2, 1 To 5) As Integer  '�C���q�X�P�����μƭȲέp���(1.�ϥΪ�/2.�q��,1.�C/2.��/3.��/4.�S/5.�j)
Public ��ƹs�ˬd��(1 To 2) As Boolean '��e���q��l�ƶq�O�_���s�ˬd��(1.�ϥΪ�/2.�q��)
Public pagecardnum(1 To 999, 1 To 11) As String '���εP���(��x�s��,1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.(1)�ϥΪ�-(2)�q��/6.(1)��P-(2)�X�P-(3)�õP-(4)�P��/7.�X�P����/8.�Ϥ��s��/9.�ثeLeft(�y��)/10.�ثeTop(�y��)/11.(1)�q����X�P(��)-(2)�q���o�X�P(�~))
Public �P�`���q��(1 To 3) As Integer '�P�֦��`���q��(1.�ϥΪ�/2.�q��/3.�`�p)
Public �P���ʼȮ��ܼ�(1 To 3) As Long '�P���ʭp�ƾ��Ȯ��ܼ�(1.Left���/2.Top���/3.�P�i�s��)
Public �ثe��(1 To 33) As Integer '�`�Ȯ��ܼ�
Public �X�P���ǲέp�Ȯ��ܼ�(1 To 4, 1 To 999, 1 To 2) As Integer '�X�P���ǲέp�`�Ȯɸ��(1.�ϥΪ̥X�P/2.�ϥΪ̤�P/3.�q���X�P/4.�q����P,��x����,1.�ثe�P�X�P����/2.�P�i�s��)
Public �Z�����_���P�Ȯɼ�(1 To 999, 1 To 3) As Integer  '���P�ӧO�Z�����Ȯ��x�s�ܼ�(��x����,1.Left���/2.Top���/3.�P�i�s��)
Public ���q���A�� As Integer '�C���q�}�l�������A�ˬd��(1.�}�l���q(�ϥΪ�)/2.�������q(�ϥΪ�)/3.�}�l���q(�q��)/4.�������q(�q��)/5.�洫����)
Public �p�H���Y�����ʤ�V��(1 To 2) As Integer '�p�H���Y�����ʤ�V���A��(1.�ϥΪ�/2.�q��[1.�V��,2.�V�~])
Public ��q�p�ƾ��ʵe�Ȯ��ܼ�(1 To 2, 1 To 2) As Integer '�}�l��l���q-��q�ʵe�p�ƾ��Ȯ��ܼ�(1.�ϥΪ̦��/2.�q�����,1.�C�����ʶq/2.�O�_�w����)
Public �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1 To 4, 1 To 3) As Integer '�ɶ��b�i���C���ܤƶ��q�����Ȯ��ܼ�(1~3(1)����ܤƶq(1(1).�ɶ��b(���~))/2.�ثe�֭p�q/3.�ثe�C��(R,G,B),4.(1)�ɶ��b(�~)���q��-(1)���ܬ�-(2)���ܶ�/2.�ثe�֭p�q/3.�ثe�C��(R))
Public �}�l�d�����ʰʵe������(1 To 2, 1 To 4) As Integer   '�}�l�ɨC�i�d�����ʰʵe����������(1.�ϥΪ�/2.�q��,1~3.�d��/4.�ثe�ĴX�i)
Public �洫��������Ȯ��ܼ�(1 To 4) As Integer '�洫������������Ȯɼ�(1.�ϥΪ�/2.�q��/3.�O�_��U����/4.�洫���⧹���涥�q��)
Public pageeventnum(1 To 2, 1 To 18, 1 To 2) As String '�ƥ�d�ƦC�������(1.�ϥΪ�/2.�q��,1~18-�s��,1.�ƥ�d�W��/2.�ƥ�d�ɮצW��)
Public �԰��Ҧ��ӱѬ����� As Integer '�԰��t�η�e�ӱѬ����Ȯ��ܼ�(1.�ϥΪ̤�ӧQ/2.�ϥΪ̤�ѥ_/3.����)
Public �q���貾�ʶ��q��ܼ� As Integer '���ʶ��q�q�����ܤ���ʼȮ��ܼ�
Public �q����ƥ�d�O�_�X����ܼ� As Boolean '�q������X�ƥ�d�O�_�X���Ȯɬ���
Public �H���d���I���s��������(1 To 7) As Integer '�H���d���I���ޯ໡���H���s���Ȯ��ܼ�(1.(1).�ϥΪ�/(2).�q��,2.��n��,3.�ثe�ϥΪ̤�ϥΤH���s��/4.�ثe��ܤ��ޯ�s��(�ϥΪ̤�ϥΤH��)/5.�ثe��ܤ��ޯ�s��(��L)/6~7.�ثe��ܤ��ޯ�s��(�洫����)
Public �Y���淾�q�Ȯ��ܼ�(1 To 10) As Integer '�Y�뤶�����q�Ȯ��ܼ�(1.�@�^�X������P�_(1.�e/2.��),2.�Y��ᦳ�Ķˮ`��,3.�Y���ˮ`��H(1.�ϥΪ�/2.�q��),4.(1.�ϥΪ̥���/2.�q������)/5.��e���(�ϥΪ�)/6.��e���(�q��)/7.�t�Τ��λ��(�ϥΪ�)/8.�t�Τ��λ��(�q��)/9.�Y��e���-�`��(�ϥΪ�)/10.�Y��e���-�`��(�q��))
Public �H�������ˬd�Ȯ��ܼ�(1 To 3) As Integer '�H�������ˬd�p�ƾ������Ȯ��ܼ�(1.�ثe�p��/2.�ϥΪ̼аO/3.�q���аO)
Public ���εP�U�P����������(0 To 31, 1 To 2) As Integer '�U�������εP�P���������Ȯ��ܼ�(0.(1)�ثe�w�o�P�`�ƶq/(2)�ثe�����P�`�ƶq,1~31.(1)�ثe�w�ϥΤ��P��/(2)�ӵP����ϥΤ��`�ƶq)
Public �d���H����T�ɮ�Ū�����Ѭ����� As String '�d���H����T�ɮ�Ū�����Ѯ��ɮצW�����Ȯ��ܼ�
Public ���εP����d�����j������(1 To 5) As Integer '�԰��t�ι���P����������(1.�`�@�P��/2.���P�P��/3.�ϥΪ̨ƥ�d�̩��s��/4.�q���ƥ�d�̩��s��/5.�ۥѤ��t����P�}�l�s��)
Public ��ܦC����ƭ���w������(1 To 2) As Boolean '�԰��t����ܦC����ƭ���w��ܬ����ܼ�(1.�ϥΪ̤�/2.�q����)
Public �O�_�t�Τ��� As Boolean '�O�_���t�Τ��������
Public �԰��Y�뤶���H����ø�ϸ��|������(1 To 2) As String '�԰��t���Y�뤶������H����ø�ϸ��|������(1.�ϥΪ̤�/2.�q����)
Public �H����ڪ��A��Ʈw(1 To 2, 1 To 3, 1 To 9) As String '�H����ڪ��A���
Public �t����ܬɭ������� As Integer '�԰��t����ܤ����]�w������(1.�ª�/2.�s��)
Public ���ݮɶ���C(1 To 2) As New Collection '�԰��t�ε��ݮɶ��p�ƾ��u�@��C
Public �H�����`���A�C��(1 To 2, 1 To 3) As Collection '���`���A�C��(1.�ϥΪ�/2.�q��,��n��)
Public ActiveSkillObj(1 To 2, 1 To 4) As clsPersonActiveSkill '�԰��t�ΥD�ʧޯ໡������(1.�ϥΪ̤�/2.�q����,��n��)
Public PersonCardShowOnMode(1 To 2, 1 To 3) As Boolean '�԰��t�ΤH���d����T�O�_�i��(1.�ϥΪ̤�/2.�q����,��n��)
Sub �H���ޯ���O�}��(ByVal isOn As Boolean, ByVal num As Integer)
FormMainMode.PEAFInterface.ActiveSkillLight 1, num, isOn
End Sub
Function ����ʧ@_���|�ϥηs�����`���A�Ϯ�(ByVal ph As String) As String
Dim i As Integer
For i = 1 To Len(ph)
    If Mid(ph, i, 1) = "." Then
        ph = Mid(ph, 1, i - 1) & "new" & Right(ph, 4)
        Exit For
    End If
Next
����ʧ@_���|�ϥηs�����`���A�Ϯ� = ph
End Function
Sub �ˮ`����_�ޯઽ��_�ϥΪ�(ByVal tot As Integer, ByVal num As Integer, ByVal isEvent As Boolean)
If tot <= 0 Then Exit Sub
If isEvent = True Then
'===============================
    Vss_EventBloodActionOffNum = 0
    VBEStageNum(0) = 46
    VBEStageNum(1) = -1 '����ˮ`��(1.�ϥΪ�/2.�q��)
    VBEStageNum(2) = num '����ˮ`�H���s��
    VBEStageNum(3) = 2 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    VBEStageNum(4) = tot '����ˮ`���ƭ�
    Vss_EventBloodActionChangeNum(0) = 0
    Vss_EventBloodActionChangeNum(1) = 1 '����ˮ`��(1.�ϥΪ�/2.�q��)
    Vss_EventBloodActionChangeNum(2) = num '����ˮ`�H���s��
    Vss_EventBloodActionChangeNum(3) = 2 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    Vss_EventBloodActionChangeNum(4) = tot  '����ˮ`���ƭ�
    '===========================���涥�q���J�I(46)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 46, 1
    '============================
    If Vss_EventBloodActionChangeNum(0) = 1 Then tot = Vss_EventBloodActionChangeNum(4)
    If Vss_EventBloodActionOffNum = 1 Then Exit Sub
End If
Select Case num
   Case 1
      If tot > 0 And liveus(����H����ԤH��(1, 2)) > 0 Then
          If tot >= liveus(����H����ԤH��(1, 2)) Then
             �԰��t����.�s���T�� "�z����F" & liveus(����H����ԤH��(1, 2)) & "�I�ˮ`�C"
             FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = 0
             liveus(����H����ԤH��(1, 2)) = 0
             FormMainMode.bloodnumus1.Caption = 0
             FormMainMode.bloodlineout1.Width = 0
             �P�`���q��(1) = �P�`���q��(1) + 1
          Else
             FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = liveus(����H����ԤH��(1, 2)) - tot
             liveus(����H����ԤH��(1, 2)) = liveus(����H����ԤH��(1, 2)) - tot
             FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
             FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (�Z�����(1, 1, 1) * tot)
             �԰��t����.�s���T�� "�z����F" & tot & "�I�ˮ`�C"
          End If
          FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).CurrentHP = liveus(����H����ԤH��(1, 2))
          �԰��t����.����ˮ`����
       End If
   Case Is > 1
       If tot > 0 And liveus(����ݾ��H��������(1, num)) > 0 Then
          If tot >= liveus(����ݾ��H��������(1, num)) Then
             liveus(����ݾ��H��������(1, num)) = 0
             FormMainMode.cardus(����ݾ��H��������(1, num)).CardMain_����HP = 0
             �P�`���q��(1) = �P�`���q��(1) + 1
          Else
             liveus(����ݾ��H��������(1, num)) = liveus(����ݾ��H��������(1, num)) - tot
             FormMainMode.cardus(����ݾ��H��������(1, num)).CardMain_����HP = liveus(����ݾ��H��������(1, num))
          End If
          FormMainMode.PEAFpersoncardus(����ݾ��H��������(1, num)).CurrentHP = liveus(����ݾ��H��������(1, num))
       End If
End Select

End Sub
Sub ��q��s���()
�������m��l�`��(1) = 0
�������m��l�`��(2) = 0
Erase ��ܦC����ƭ���w������
Erase atkingckdice
Erase Vss_EventPersonAbilityDiceChangeNum
Dim uscom As Integer
'===========================���涥�q���J�I(45)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 45, 1
'============================
For uscom = 1 To 2
    Select Case uscom
        Case 1
            If turnatk = 1 Then
                If atkingpagetot(1, 1) > 0 And movecp = 1 Then
                    If Vss_EventPersonAbilityDiceChangeNum(1, 2) = 0 Then
                        �������m��l�`��(1) = �������m��l�`��(1) + atkus(����H����ԤH��(1, 2))
                    End If
                    �������m��l�`��(1) = �������m��l�`��(1) + Vss_EventPersonAbilityDiceChangeNum(1, 1)
                    �������m��l�`��(1) = �������m��l�`��(1) + atkingpagetot(1, 1)
                ElseIf atkingpagetot(1, 5) > 0 And movecp > 1 Then
                    If Vss_EventPersonAbilityDiceChangeNum(1, 2) = 0 Then
                        �������m��l�`��(1) = �������m��l�`��(1) + atkus(����H����ԤH��(1, 2))
                    End If
                    �������m��l�`��(1) = �������m��l�`��(1) + Vss_EventPersonAbilityDiceChangeNum(1, 1)
                    �������m��l�`��(1) = �������m��l�`��(1) + atkingpagetot(1, 5)
                End If
            ElseIf turnatk = 2 Then
                If Vss_EventPersonAbilityDiceChangeNum(1, 2) = 0 Then
                    �������m��l�`��(1) = �������m��l�`��(1) + defus(����H����ԤH��(1, 2))
                End If
                �������m��l�`��(1) = �������m��l�`��(1) + Vss_EventPersonAbilityDiceChangeNum(1, 1)
                �������m��l�`��(1) = �������m��l�`��(1) + atkingpagetot(1, 2)
            End If
            '=======�D�ʧ�
            �ѪR��q�ܤ� atkingckdice(1, 1, 1), 1
            '=======�Q�ʧ�
            �ѪR��q�ܤ� atkingckdice(1, 1, 2), 1
            '=======���`���A
            �ѪR��q�ܤ� atkingckdice(1, 1, 3), 1
            '=======�H����ڪ��A
            �ѪR��q�ܤ� atkingckdice(1, 1, 4), 1
            '=================================���
            '=======�D�ʧ�
            �ѪR��q�ܤ� atkingckdice(2, 1, 1), 1
            '=======�Q�ʧ�
            �ѪR��q�ܤ� atkingckdice(2, 1, 2), 1
            '=======���`���A
            �ѪR��q�ܤ� atkingckdice(2, 1, 3), 1
            '=======�H����ڪ��A
            �ѪR��q�ܤ� atkingckdice(2, 1, 4), 1
            '===================================
'            FormMainMode.trgoi1_Timer
        Case 2
            If turnatk = 2 Then
                If atkingpagetot(2, 1) > 0 And movecp = 1 Then
                    If Vss_EventPersonAbilityDiceChangeNum(2, 2) = 0 Then
                        �������m��l�`��(2) = �������m��l�`��(2) + atkcom(����H����ԤH��(2, 2))
                    End If
                    �������m��l�`��(2) = �������m��l�`��(2) + Vss_EventPersonAbilityDiceChangeNum(2, 1)
                    �������m��l�`��(2) = �������m��l�`��(2) + atkingpagetot(2, 1)
                ElseIf atkingpagetot(2, 5) > 0 And movecp > 1 Then
                    If Vss_EventPersonAbilityDiceChangeNum(2, 2) = 0 Then
                        �������m��l�`��(2) = �������m��l�`��(2) + atkcom(����H����ԤH��(2, 2))
                    End If
                    �������m��l�`��(2) = �������m��l�`��(2) + Vss_EventPersonAbilityDiceChangeNum(2, 1)
                    �������m��l�`��(2) = �������m��l�`��(2) + atkingpagetot(2, 5)
                End If
            ElseIf turnatk = 1 Then
                If Vss_EventPersonAbilityDiceChangeNum(2, 2) = 0 Then
                    �������m��l�`��(2) = �������m��l�`��(2) + defcom(����H����ԤH��(2, 2))
                End If
                �������m��l�`��(2) = �������m��l�`��(2) + Vss_EventPersonAbilityDiceChangeNum(2, 1)
                �������m��l�`��(2) = �������m��l�`��(2) + atkingpagetot(2, 2)
            End If
            '=======�D�ʧ�
            �ѪR��q�ܤ� atkingckdice(2, 2, 1), 2
            '=======�Q�ʧ�
            �ѪR��q�ܤ� atkingckdice(2, 2, 2), 2
            '=======���`���A
            �ѪR��q�ܤ� atkingckdice(2, 2, 3), 2
            '=======�H����ڪ��A
            �ѪR��q�ܤ� atkingckdice(2, 2, 4), 2
            '=================================���
            '=======�D�ʧ�
            �ѪR��q�ܤ� atkingckdice(1, 2, 1), 2
            '=======�Q�ʧ�
            �ѪR��q�ܤ� atkingckdice(1, 2, 2), 2
            '=======���`���A
            �ѪR��q�ܤ� atkingckdice(1, 2, 3), 2
            '=======�H����ڪ��A
            �ѪR��q�ܤ� atkingckdice(1, 2, 4), 2
            '===================================
    End Select
Next
End Sub

Sub ����ˮ`����()
Select Case movecp
    Case 1
        �@��t����.���ļ��� 2
    Case Is >= 2
        �@��t����.���ļ��� 8
End Select
End Sub
Sub �^�_����_�ϥΪ�(ByVal tot As Integer, ByVal num As Integer, ByVal statusfrom As Integer, ByVal isEvent As Boolean)
If isEvent = True Then
    '===============================
    If statusfrom = 0 Then
        ReDim VBEStageNum(0 To 5) As Integer
        VBEStageNum(4) = 0 'Ĳ�o�ƥ��
        VBEStageNum(5) = 0 'Ĳ�o�ƥ���t
    End If
    Vss_EventHPLActionOffNum = 0
    VBEStageNum(0) = 48
    VBEStageNum(1) = -1 '�^�_��(1.�ϥΪ�/2.�q��)
    VBEStageNum(2) = num '�^�_�H���s��
    VBEStageNum(3) = tot '�^�_���ƭ�
    Vss_EventHPLActionChangeNum(0) = 0
    Vss_EventHPLActionChangeNum(1) = tot  '�^�_���ƭ�
    '===========================���涥�q���J�I(48)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 48, 1
    '============================
    If Vss_EventHPLActionChangeNum(0) = 1 Then tot = Vss_EventHPLActionChangeNum(1)
    If Vss_EventHPLActionOffNum = 1 Then Exit Sub
End If

Select Case num
   Case 1
         If liveus(����H����ԤH��(1, 2)) > 0 And tot > 0 Then
            If liveusmax(����H����ԤH��(1, 2)) - liveus(����H����ԤH��(1, 2)) >= tot Then
                �԰��t����.�s���T�� "�A��HP��_�F" & tot & "�I�C"
                FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + �Z�����(1, 1, 1) * tot
                liveus(����H����ԤH��(1, 2)) = Val(liveus(����H����ԤH��(1, 2))) + tot
            ElseIf liveusmax(����H����ԤH��(1, 2)) - liveus(����H����ԤH��(1, 2)) < tot Then
                If liveusmax(����H����ԤH��(1, 2)) - liveus(����H����ԤH��(1, 2)) > 0 Then
                   �԰��t����.�s���T�� "�A��HP��_�F" & Val(liveusmax(����H����ԤH��(1, 2))) - Val(liveus(����H����ԤH��(1, 2))) & "�I�C"
                   FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + �Z�����(1, 1, 1) * (Val(liveusmax(����H����ԤH��(1, 2))) - Val(liveus(����H����ԤH��(1, 2))))
                   liveus(����H����ԤH��(1, 2)) = Val(liveusmax(����H����ԤH��(1, 2)))
                End If
            End If
            FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = liveus(����H����ԤH��(1, 2))
            FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).CurrentHP = liveus(����H����ԤH��(1, 2))
            FormMainMode.bloodnumus1.Caption = liveus(����H����ԤH��(1, 2))
        End If
   Case Is > 1
        If liveus(����ݾ��H��������(1, num)) > 0 And tot > 0 Then
            If liveusmax(����ݾ��H��������(1, num)) - liveus(����ݾ��H��������(1, num)) >= tot Then
                liveus(����ݾ��H��������(1, num)) = Val(liveus(����ݾ��H��������(1, num))) + tot
            ElseIf liveusmax(����ݾ��H��������(1, num)) - liveus(����ݾ��H��������(1, num)) < tot Then
                liveus(����ݾ��H��������(1, num)) = Val(liveusmax(����ݾ��H��������(1, num)))
            End If
            FormMainMode.cardus(����ݾ��H��������(1, num)).CardMain_����HP = liveus(����ݾ��H��������(1, num))
            FormMainMode.PEAFpersoncardus(����ݾ��H��������(1, num)).CurrentHP = liveus(����ݾ��H��������(1, num))
        End If
End Select
End Sub
Sub �^�_����_�q��(ByVal tot As Integer, ByVal num As Integer, ByVal statusfrom As Integer, ByVal isEvent As Boolean)
If isEvent = True Then
    '===============================
    If statusfrom = 0 Then
        ReDim VBEStageNum(0 To 5) As Integer
        VBEStageNum(4) = 0 'Ĳ�o�ƥ��
        VBEStageNum(5) = 0 'Ĳ�o�ƥ���t
    End If
    Vss_EventHPLActionOffNum = 0
    VBEStageNum(0) = 48
    VBEStageNum(1) = -2 '�^�_��(�t�ΥN��)
    VBEStageNum(2) = num '�^�_�H���s��
    VBEStageNum(3) = tot '�^�_���ƭ�
    Vss_EventHPLActionChangeNum(0) = 0
    Vss_EventHPLActionChangeNum(1) = tot  '�^�_���ƭ�
    '===========================���涥�q���J�I(48)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 48, 1
    '============================
    If Vss_EventHPLActionChangeNum(0) = 1 Then tot = Vss_EventHPLActionChangeNum(1)
    If Vss_EventHPLActionOffNum = 1 Then Exit Sub
End If

Select Case num
   Case 1
         If livecom(����H����ԤH��(2, 2)) > 0 And tot > 0 Then
            If livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2)) >= tot Then
                �԰��t����.�s���T�� "��誺HP��_�F" & tot & "�I�C"
                FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - �Z�����(1, 2, 1) * tot
                livecom(����H����ԤH��(2, 2)) = Val(livecom(����H����ԤH��(2, 2))) + tot
            ElseIf livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2)) < tot Then
                If livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2)) > 0 Then
                   �԰��t����.�s���T�� "��誺HP��_�F" & Val(livecommax(����H����ԤH��(2, 2))) - Val(livecom(����H����ԤH��(2, 2))) & "�I�C"
                   FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - �Z�����(1, 2, 1) * (Val(livecommax(����H����ԤH��(2, 2))) - Val(livecom(����H����ԤH��(2, 2))))
                   livecom(����H����ԤH��(2, 2)) = Val(livecommax(����H����ԤH��(2, 2)))
                End If
            End If
            FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).CurrentHP = livecom(����H����ԤH��(2, 2))
            FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = livecom(����H����ԤH��(2, 2))
            FormMainMode.bloodnumcom1.Caption = livecom(����H����ԤH��(2, 2))
        End If
   Case Is > 1
        If livecom(����ݾ��H��������(2, num)) > 0 And tot > 0 Then
            If livecommax(����ݾ��H��������(2, num)) - livecom(����ݾ��H��������(2, num)) >= tot Then
                livecom(����ݾ��H��������(2, num)) = Val(livecom(����ݾ��H��������(2, num))) + tot
            ElseIf livecommax(����ݾ��H��������(2, num)) - livecom(����ݾ��H��������(2, num)) < tot Then
                livecom(����ݾ��H��������(2, num)) = Val(livecommax(����ݾ��H��������(2, num)))
            End If
            FormMainMode.cardcom(����ݾ��H��������(2, num)).CardMain_����HP = livecom(����ݾ��H��������(2, num))
            FormMainMode.PEAFpersoncardcom(����ݾ��H��������(2, num)).CurrentHP = livecom(����ݾ��H��������(2, num))
        End If
End Select
End Sub
Sub �ˮ`����_�ϥΪ�(ByVal tot As Integer)
If tot <= 0 Then Exit Sub
'===============================
ReDim VBEStageNum(0 To 6) As Integer
Vss_EventBloodActionOffNum = 0
VBEStageNum(0) = 46
VBEStageNum(1) = -1 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = 1 '����ˮ`�H���s��
VBEStageNum(3) = 1 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = tot '����ˮ`���ƭ�
VBEStageNum(5) = 0 '�Ӧۨt�Ϊ��ˮ`
VBEStageNum(6) = 0 '�Ӧۨt�Ϊ��ˮ`
Vss_EventBloodActionChangeNum(0) = 0
Vss_EventBloodActionChangeNum(1) = 1 '����ˮ`��(1.�ϥΪ�/2.�q��)
Vss_EventBloodActionChangeNum(2) = 1 '����ˮ`�H���s��
Vss_EventBloodActionChangeNum(3) = 1 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
Vss_EventBloodActionChangeNum(4) = tot  '����ˮ`���ƭ�
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 46, 1
'============================
If Vss_EventBloodActionChangeNum(0) = 1 Then tot = Vss_EventBloodActionChangeNum(4)
If Vss_EventBloodActionOffNum = 0 Then
    If tot > 0 And liveus(����H����ԤH��(1, 2)) > 0 Then
          If tot >= liveus(����H����ԤH��(1, 2)) Then
             �԰��t����.�s���T�� "�z����F" & liveus(����H����ԤH��(1, 2)) & "�I�ˮ`�C"
             FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = 0
             liveus(����H����ԤH��(1, 2)) = 0
             FormMainMode.bloodnumus1.Caption = 0
             FormMainMode.bloodlineout1.Width = 0
             �P�`���q��(1) = �P�`���q��(1) + 1
          Else
             FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = liveus(����H����ԤH��(1, 2)) - tot
             liveus(����H����ԤH��(1, 2)) = liveus(����H����ԤH��(1, 2)) - tot
             FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
             FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (�Z�����(1, 1, 1) * tot)
             �԰��t����.�s���T�� "�z����F" & tot & "�I�ˮ`�C"
          End If
          FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).CurrentHP = liveus(����H����ԤH��(1, 2))
    �԰��t����.����ˮ`����
    End If
End If
End Sub
Sub �ˮ`����_�ޯઽ��_�q��(ByVal tot As Integer, ByVal num As Integer, ByVal isEvent As Boolean)
If tot <= 0 Then Exit Sub
If isEvent = True Then
    '===============================
    Vss_EventBloodActionOffNum = 0
    VBEStageNum(0) = 46
    VBEStageNum(1) = -2 '����ˮ`��(1.�ϥΪ�/2.�q��)
    VBEStageNum(2) = num '����ˮ`�H���s��
    VBEStageNum(3) = 2 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    VBEStageNum(4) = tot '����ˮ`���ƭ�
    Vss_EventBloodActionChangeNum(0) = 0
    Vss_EventBloodActionChangeNum(1) = 2 '����ˮ`��(1.�ϥΪ�/2.�q��)
    Vss_EventBloodActionChangeNum(2) = num '����ˮ`�H���s��
    Vss_EventBloodActionChangeNum(3) = 2 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    Vss_EventBloodActionChangeNum(4) = tot  '����ˮ`���ƭ�
    '===========================���涥�q���J�I(46)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 46, 1
    '============================
    If Vss_EventBloodActionChangeNum(0) = 1 Then tot = Vss_EventBloodActionChangeNum(4)
    If Vss_EventBloodActionOffNum = 1 Then Exit Sub
End If
Select Case num
    Case 1
       If tot > 0 And livecom(����H����ԤH��(2, 2)) > 0 Then
            If tot >= livecom(����H����ԤH��(2, 2)) Then
               �԰��t����.�s���T�� "������F" & livecom(����H����ԤH��(2, 2)) & "�I�ˮ`�C"
               FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = 0
               FormMainMode.bloodnumcom1.Caption = 0
               livecom(����H����ԤH��(2, 2)) = 0
               FormMainMode.bloodlineout2.Left = 11580
               �P�`���q��(2) = �P�`���q��(2) + 1
            Else
               �԰��t����.�s���T�� "������F" & Val(tot) & "�I�ˮ`�C"
               FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = livecom(����H����ԤH��(2, 2)) - tot
               FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
               livecom(����H����ԤH��(2, 2)) = livecom(����H����ԤH��(2, 2)) - tot
               FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (�Z�����(1, 2, 1) * tot)
            End If
            FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).CurrentHP = livecom(����H����ԤH��(2, 2))
            �԰��t����.����ˮ`����
        End If
    Case Is > 1
       If tot > 0 And livecom(����ݾ��H��������(2, num)) > 0 Then
                If tot >= livecom(����ݾ��H��������(2, num)) Then
                    livecom(����ݾ��H��������(2, num)) = 0
                    FormMainMode.cardcom(����ݾ��H��������(2, num)).CardMain_����HP = 0
                    �P�`���q��(2) = �P�`���q��(2) + 1
                Else
                    livecom(����ݾ��H��������(2, num)) = livecom(����ݾ��H��������(2, num)) - tot
                    FormMainMode.cardcom(����ݾ��H��������(2, num)).CardMain_����HP = livecom(����ݾ��H��������(2, num))
                End If
                FormMainMode.PEAFpersoncardcom(����ݾ��H��������(2, num)).CurrentHP = livecom(����ݾ��H��������(2, num))
        End If
End Select
End Sub
Sub �ˮ`����_�q��(ByVal tot As Integer)
If tot <= 0 Then Exit Sub
'===============================
ReDim VBEStageNum(0 To 6) As Integer
Vss_EventBloodActionOffNum = 0
VBEStageNum(0) = 46
VBEStageNum(1) = -2 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = 1 '����ˮ`�H���s��
VBEStageNum(3) = 1 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = tot '����ˮ`���ƭ�
VBEStageNum(5) = 0 '�Ӧۨt�Ϊ��ˮ`
VBEStageNum(6) = 0 '�Ӧۨt�Ϊ��ˮ`
Vss_EventBloodActionChangeNum(0) = 0
Vss_EventBloodActionChangeNum(1) = 2 '����ˮ`��(1.�ϥΪ�/2.�q��)
Vss_EventBloodActionChangeNum(2) = 1 '����ˮ`�H���s��
Vss_EventBloodActionChangeNum(3) = 1 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
Vss_EventBloodActionChangeNum(4) = tot  '����ˮ`���ƭ�
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 46, 1
'============================
If Vss_EventBloodActionChangeNum(0) = 1 Then tot = Vss_EventBloodActionChangeNum(4)
If Vss_EventBloodActionOffNum = 0 Then
    If tot > 0 And livecom(����H����ԤH��(2, 2)) > 0 Then
        If tot >= livecom(����H����ԤH��(2, 2)) Then
           �԰��t����.�s���T�� "������F" & livecom(����H����ԤH��(2, 2)) & "�I�ˮ`�C"
           FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = 0
           FormMainMode.bloodnumcom1.Caption = 0
           livecom(����H����ԤH��(2, 2)) = 0
           FormMainMode.bloodlineout2.Left = 11580
           �P�`���q��(2) = �P�`���q��(2) + 1
        Else
           �԰��t����.�s���T�� "������F" & Val(tot) & "�I�ˮ`�C"
           FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = livecom(����H����ԤH��(2, 2)) - tot
           FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
           livecom(����H����ԤH��(2, 2)) = livecom(����H����ԤH��(2, 2)) - tot
           FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (�Z�����(1, 2, 1) * tot)
        End If
        FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).CurrentHP = livecom(����H����ԤH��(2, 2))
        �԰��t����.����ˮ`����
    End If
End If
End Sub
Sub ����ʧ@_�ϥΪ�_��P(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) - 1
    �ثe��(5) = pagecardnum(n, 7)
    pagecardnum(n, 6) = 3
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �ثe��(15) = 4
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�P��_�^�P_�ϥΪ�(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    pagecardnum(n, 5) = 1
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.���εP�^�_���� n
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�ϥΪ� n
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�q���P_���P_�ϥΪ�(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    �ثe��(9) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 1
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�ϥΪ� n
    �ثe��(15) = 2
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�ϥΪ̵P_���P_�q��(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pageusglead = Val(FormMainMode.pageusglead) - 1
    �ثe��(5) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 2
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�q�� n
    �ثe��(15) = 20
    �԰��t����.���εP�ܭI��
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�P��_�^�P_�q��(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
'    �ثe��(5) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 2
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�q�� n
    �԰��t����.���εP�ܭI��
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_½�P(ByVal n As Integer)
    FormMainMode.card(n).Width = 810
    FormMainMode.card(n).Height = 1260
'    FormMainMode.card(n).Picture = LoadPicture(app_path & "card\" & pagecardnum(n, 8) & "-" & pageonin(n) & ".bmp")
    FormMainMode.card(n).LocationType = 1
    FormMainMode.card(n).CardEventType = False
    FormMainMode.card(n).Visible = True
    �@��t����.���ļ��� 4
End Sub
Sub �y�Эp��_�q���X�P()
Dim xy As Long  '�Ȯ��ܼ�(���PLeft)
If pageqlead(2) = 1 Then
    �P���ʼȮ��ܼ�(1) = 5260
    �P���ʼȮ��ܼ�(2) = 1120
ElseIf pageqlead(2) > 1 Then
    xy = (pageqlead(2) - 1) * 460
    �P���ʼȮ��ܼ�(1) = (Val(5260) - xy) + ((pageqlead(2) - 1) * Val(960))
    �P���ʼȮ��ܼ�(2) = 1120
End If

End Sub
Sub �y�Эp��_�q����P()
�P���ʼȮ��ܼ�(1) = 10560 - 240 * (Val(FormMainMode.pagecomglead) - 1) '�p��Left�y��
�P���ʼȮ��ܼ�(2) = -600 '���wTop�y��
End Sub
Sub �y�Эp��_�ϥΪ̥X�P()
Dim xy As Long   '�Ȯ��ܼ�(���PLeft)
If pageqlead(1) = 1 Then
    �P���ʼȮ��ܼ�(1) = 5260
    �P���ʼȮ��ܼ�(2) = 4840
ElseIf pageqlead(1) > 1 Then
    xy = (pageqlead(1) - 1) * 460
    �P���ʼȮ��ܼ�(1) = (Val(5260) - xy) + ((pageqlead(1) - 1) * Val(960))
    �P���ʼȮ��ܼ�(2) = 4840
End If

End Sub
Sub �y�Эp��_�ϥΪ̤�P()
If Val(FormMainMode.pageusglead) <= 9 Then
    �P���ʼȮ��ܼ�(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 1) '�p��Left�y��
Else
   �P���ʼȮ��ܼ�(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 10)
End If

If Val(FormMainMode.pageusglead) <= 9 Then
   �P���ʼȮ��ܼ�(2) = 6700 '���wTop�y��
Else
   �P���ʼȮ��ܼ�(2) = 7980 '���wTop�y��
End If
End Sub
Sub �P���ǼW�[_�X�P_�q��(ByRef m As Integer)
pagecardnum(m, 7) = pagecomleadmax(1) + 1
pagecomleadmax(1) = pagecomleadmax(1) + 1
End Sub
Sub �P���ǼW�[_��P_�q��(ByRef m As Integer)
pagecardnum(m, 7) = pagecomleadmax(0) + 1
pagecomleadmax(0) = pagecomleadmax(0) + 1
End Sub
Sub �P���ǼW�[_��P_�ϥΪ�(ByVal m As Integer)
pagecardnum(m, 7) = pageusleadmax(0) + 1
pageusleadmax(0) = pageusleadmax(0) + 1
End Sub
Sub �P���ǼW�[_�X�P_�ϥΪ�(ByRef m As Integer)
pagecardnum(m, 7) = pageusleadmax(1) + 1
pageusleadmax(1) = pageusleadmax(1) + 1
End Sub
Sub ����ʧ@_�q��_��P(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) - 1
    �ثe��(9) = pagecardnum(n, 7)
    pagecardnum(n, 6) = 3
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �ثe��(15) = 5
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�~�P()
Dim g As Integer
For g = 1 To ���εP����d�����j������(2)
     If pagecardnum(g, 6) = 3 Then
         ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) - 1
         pagecardnum(g, 6) = 4
         Select Case pagecardnum(g, 8)
            Case "021"  '==��1�j1��
                 ���εP�U�P����������(1, 1) = Val(���εP�U�P����������(1, 1)) - 1
            Case "019"  '==��1�j2��
                 ���εP�U�P����������(2, 1) = Val(���εP�U�P����������(2, 1)) - 1
            Case "017"  '==��1�j3��
                 ���εP�U�P����������(3, 1) = Val(���εP�U�P����������(3, 1)) - 1
            Case "025"  '==��1��1��
                 ���εP�U�P����������(4, 1) = Val(���εP�U�P����������(4, 1)) - 1
            Case "024"  '==��1��2��
                 ���εP�U�P����������(5, 1) = Val(���εP�U�P����������(5, 1)) - 1
            Case "023"  '==��1��3��
                 ���εP�U�P����������(6, 1) = Val(���εP�U�P����������(6, 1)) - 1
            Case "026"  '==��2�S3��
                 ���εP�U�P����������(7, 1) = Val(���εP�U�P����������(7, 1)) - 1
            Case "027"  '==��3��3��
                 ���εP�U�P����������(8, 1) = Val(���εP�U�P����������(8, 1)) - 1
            Case "001"  '==�C6�C6��
                 ���εP�U�P����������(9, 1) = Val(���εP�U�P����������(9, 1)) - 1
            Case "011"  '==�C1�j1��
                 ���εP�U�P����������(10, 1) = Val(���εP�U�P����������(10, 1)) - 1
            Case "007"  '==�C2�j1��
                 ���εP�U�P����������(11, 1) = Val(���εP�U�P����������(11, 1)) - 1
            Case "006"  '==�C2�j2��
                 ���εP�U�P����������(12, 1) = Val(���εP�U�P����������(12, 1)) - 1
            Case "004"  '==�C3�j3��
                 ���εP�U�P����������(13, 1) = Val(���εP�U�P����������(13, 1)) - 1
            Case "028"  '==�C5�j5��
                 ���εP�U�P����������(14, 1) = Val(���εP�U�P����������(14, 1)) - 1
            Case "012"  '==�C1��1��
                 ���εP�U�P����������(15, 1) = Val(���εP�U�P����������(15, 1)) - 1
            Case "009"  '==�C2��1��
                 ���εP�U�P����������(16, 1) = Val(���εP�U�P����������(16, 1)) - 1
            Case "008"  '==�C2��2��
                 ���εP�U�P����������(17, 1) = Val(���εP�U�P����������(17, 1)) - 1
            Case "005"  '==�C3��3��
                 ���εP�U�P����������(18, 1) = Val(���εP�U�P����������(18, 1)) - 1
            Case "013"  '==�C1�S1��
                 ���εP�U�P����������(19, 1) = Val(���εP�U�P����������(19, 1)) - 1
            Case "010"  '==�C2�S1��
                 ���εP�U�P����������(20, 1) = Val(���εP�U�P����������(20, 1)) - 1
            Case "003"  '==�C4�S1��
                 ���εP�U�P����������(21, 1) = Val(���εP�U�P����������(21, 1)) - 1
            Case "002"  '==�C5�S2��
                 ���εP�U�P����������(22, 1) = Val(���εP�U�P����������(22, 1)) - 1
            Case "015"  '==�j4�j4��
                 ���εP�U�P����������(23, 1) = Val(���εP�U�P����������(23, 1)) - 1
            Case "020"  '==�j2�S1��
                 ���εP�U�P����������(24, 1) = Val(���εP�U�P����������(24, 1)) - 1
            Case "018"  '==�j3�S2��
                 ���εP�U�P����������(25, 1) = Val(���εP�U�P����������(25, 1)) - 1
            Case "016"  '==�j4�S1��
                 ���εP�U�P����������(26, 1) = Val(���εP�U�P����������(26, 1)) - 1
            Case "014"  '==�j5�S2��
                 ���εP�U�P����������(27, 1) = Val(���εP�U�P����������(27, 1)) - 1
            Case "022"  '==��5��5��
                 ���εP�U�P����������(28, 1) = Val(���εP�U�P����������(28, 1)) - 1
            Case "029"  '==��3�S5��
                 ���εP�U�P����������(29, 1) = Val(���εP�U�P����������(29, 1)) - 1
         End Select
     End If
Next
BattleCardNum = Val(���εP�U�P����������(0, 2)) - Val(���εP�U�P����������(0, 1))
�԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
End Sub
Sub ����ʧ@_�M���Ҧ����`���A_�t��(ByVal uscom As Integer, ByVal num As Integer)
If �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num)).Count > 0 Then
    '==================
    ���涥�q�t����.���涥�q73_���O_���`���A����_�����M�� uscom, num, True
    '==================
    Dim tempnum As Integer, i As Integer
    tempnum = 1
    For i = 1 To �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num)).Count
        If VBEStageRemoveBuffAllNum(i) = False Then
            �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num)).Remove tempnum
        Else
            tempnum = tempnum + 1
        End If
    Next
    �԰��t����.���`���A��ܧ�s uscom
End If
End Sub
Sub ����ʧ@_�Z���ܧ�(ByVal m As Integer, ByVal isEvent As Boolean)
'===========================���涥�q���J�I(47)
If isEvent = True Then
    Vss_EventMoveActionOffNum = 0
    ReDim VBEStageNum(0 To 2) As Integer
    VBEStageNum(0) = 47
    VBEStageNum(1) = movecp '�ܧ�e�Z��
    VBEStageNum(2) = m  '�ܧ��Z��
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 47, 1
    '=====================
    If Vss_EventMoveActionOffNum = 1 Then Exit Sub
End If
'============================
Dim anw(1 To 2) As Integer
Dim anh(1 To 2) As Integer
anw(1) = Val(FormMainMode.personusminijpg.�p�H���Ϥ�width) / 2
anw(2) = Val(FormMainMode.personcomminijpg.�p�H���Ϥ�width) / 2
anh(1) = Val(FormMainMode.personusminijpg.�p�H���Ϥ�height)
anh(2) = Val(FormMainMode.personcomminijpg.�p�H���Ϥ�height)
Select Case m
  Case 1
    FormMainMode.PEAFMoveRange.LoadImage_FromFile app_path & "\gif\system\short.png"
    FormMainMode.PEAFMoveRange.Left = 4440
    FormMainMode.PEAFMoveRange.Top = 2520
    FormMainMode.personusminijpg.Left = 4320 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 7080 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
  Case 2
    FormMainMode.PEAFMoveRange.LoadImage_FromFile app_path & "\gif\system\middle.png"
    FormMainMode.PEAFMoveRange.Left = 2880
    FormMainMode.PEAFMoveRange.Top = 2000
    FormMainMode.personusminijpg.Left = 2640 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 8680 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
  Case 3
    FormMainMode.PEAFMoveRange.LoadImage_FromFile app_path & "\gif\system\long.png"
    FormMainMode.PEAFMoveRange.Left = 1080
    FormMainMode.PEAFMoveRange.Top = 2360
    FormMainMode.personusminijpg.Left = 1040 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 10320 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
End Select

movecp = m
End Sub
Sub �p��P���ʶZ�����()
If �P���ʼȮ��ܼ�(1) >= pagecardnum(�P���ʼȮ��ܼ�(3), 9) Then
   �Z�����(2, 1, 1) = (�P���ʼȮ��ܼ�(1) - pagecardnum(�P���ʼȮ��ܼ�(3), 9)) \ 12
Else
   �Z�����(2, 1, 1) = -((pagecardnum(�P���ʼȮ��ܼ�(3), 9) - �P���ʼȮ��ܼ�(1)) \ 12)
End If

If �P���ʼȮ��ܼ�(2) >= pagecardnum(�P���ʼȮ��ܼ�(3), 10) Then
   �Z�����(2, 1, 2) = (�P���ʼȮ��ܼ�(2) - pagecardnum(�P���ʼȮ��ܼ�(3), 10)) \ 12
Else
   �Z�����(2, 1, 2) = -((pagecardnum(�P���ʼȮ��ܼ�(3), 10) - �P���ʼȮ��ܼ�(2)) \ 12)
End If
End Sub
Sub ���`���A��ܧ�s(ByVal uscom As Integer)
Dim numNow As Integer, obj As clsStatus
Dim i As Integer, k As Integer

For i = 1 To ����H����ԤH��(uscom, 1)
    numNow = 1
    For Each obj In �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, i))
        Select Case uscom
            Case 1
                FormMainMode.cardus(����ݾ��H��������(1, i)).��ﲧ�`���A��� numNow, obj.ImagePath, obj.Value, obj.Total, True
            Case 2
                FormMainMode.cardcom(����ݾ��H��������(2, i)).��ﲧ�`���A��� numNow, obj.ImagePath, obj.Value, obj.Total, True
        End Select
        numNow = numNow + 1
        If numNow > 14 Then Exit For
    Next
    If numNow <= 14 Then
        For k = numNow To 14
            Select Case uscom
                Case 1
                    FormMainMode.cardus(����ݾ��H��������(1, i)).��ﲧ�`���A��� k, 0, 0, 0, False
                Case 2
                    FormMainMode.cardcom(����ݾ��H��������(2, i)).��ﲧ�`���A��� k, 0, 0, 0, False
            End Select
        Next
    End If
Next

End Sub

Sub �����g�J��ܦC�ƭ�(ByVal n As Integer, ByVal num As Integer)
If num < 0 Then num = 0
Select Case n
    Case 1
        FormMainMode.��ܦC1.goi1 = num
    Case 2
        FormMainMode.��ܦC1.goi2 = num
End Select
End Sub
Sub �p�H���Y�����槹�P�__�ϥΪ�()
Dim ckl As Integer

If turnatk = 1 Or turnatk = 2 Then
   turnpageonin = 1
    If Vss_EventPlayerAllActionOffNum(1) = 1 Then
        For ckl = 1 To ���εP����d�����j������(1)
            FormMainMode.card(ckl).CardEnabledType = False
        Next
        FormMainMode.PEAFInterface.BnOKEnabled False
        ���ݮɶ���C(2).Add 47
        FormMainMode.���ݮɶ�_2.Enabled = True
    ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
        For ckl = 1 To ���εP����d�����j������(1)
            FormMainMode.card(ckl).CardEnabledType = False
        Next
        FormMainMode.PEAFInterface.BnOKEnabled False
        ���ݮɶ���C(2).Add 45
        FormMainMode.���ݮɶ�_2.Enabled = True
    End If
End If
If turnatk = 3 Then
    FormMainMode.trtimeline.Enabled = True
End If
End Sub
Sub �p�H���Y�����槹�P�__�q��()
If turnatk = 1 Or turnatk = 2 Or turnatk = 3 Then
    If Vss_EventPlayerAllActionOffNum(2) = 0 Then
        ���q���A�� = 3
        FormMainMode.�q���X�P.Enabled = True
    Else
        ���ݮɶ���C(2).Add 48
        FormMainMode.���ݮɶ�_2.Enabled = True
    End If
End If
End Sub
Sub ���εP�ܭI��()
FormMainMode.card(�P���ʼȮ��ܼ�(3)).Width = 720
FormMainMode.card(�P���ʼȮ��ܼ�(3)).Height = 990
FormMainMode.card(�P���ʼȮ��ܼ�(3)).LocationType = 3
End Sub
Sub ���εP�^�_����(ByVal num As Integer)
FormMainMode.card(num).Width = 810
FormMainMode.card(num).Height = 1260
FormMainMode.card(num).LocationType = 1
FormMainMode.card(num).CardEventType = False
End Sub
Sub �X�P���ǭp��_�ϥΪ�_��P()
Dim pagegustot As Integer '�Ȯ��ܼ�
Dim i As Integer, j As Integer, o As Integer
Dim g As Integer, h As Integer

For i = 1 To 999
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(2, i, j) = 0
   Next
Next

For i = 1 To 999
   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
    pagegustot = Val(pagegustot) + 1
    �X�P���ǲέp�Ȯ��ܼ�(2, pagegustot, 1) = Val(pagecardnum(i, 7))
    �X�P���ǲέp�Ȯ��ܼ�(2, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(2, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(2, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(2, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(2, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(2, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(2, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(2, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(2, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(2, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(2, o, 2) = h
   End If
  Next
Next
'MsgBox 123
End Sub
Sub �X�P���ǭp��_�ϥΪ�_�X�P()
Dim pagegustot As Integer '�Ȯ��ܼ�
Dim i As Integer, j As Integer, o As Integer
Dim g As Integer, h As Integer

For i = 1 To 999
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(1, i, j) = 0
   Next
Next

For i = 1 To 999
   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
    pagegustot = Val(pagegustot) + 1
    �X�P���ǲέp�Ȯ��ܼ�(1, pagegustot, 1) = Val(pagecardnum(i, 7))
    �X�P���ǲέp�Ȯ��ܼ�(1, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(1, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(1, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(1, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(1, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(1, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(1, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(1, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(1, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(1, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(1, o, 2) = h
   End If
  Next
Next

End Sub
Sub �X�P���ǭp��_�q��_��P()
Dim pagegustot As Integer '�Ȯ��ܼ�
Dim i As Integer, j As Integer, o As Integer
Dim g As Integer, h As Integer

For i = 1 To 999
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(4, i, j) = 0
   Next
Next

For i = 1 To 999
   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
       pagegustot = Val(pagegustot) + 1
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 1) = Val(pagecardnum(i, 7))
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 2) = i
   ElseIf Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 1 Then
       pagegustot = Val(pagegustot) + 1
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 1) = Val(pagecardnum(i, 7))
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(4, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(4, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(4, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(4, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(4, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(4, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(4, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(4, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(4, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(4, o, 2) = h
   End If
  Next
Next
End Sub
Sub �X�P���ǭp��_�q��_�X�P()
Dim pagegustot As Integer '�Ȯ��ܼ�
Dim i As Integer, j As Integer, o As Integer
Dim g As Integer, h As Integer

For i = 1 To 999
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(3, i, j) = 0
   Next
Next

For i = 1 To 999
   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 2 Then
       pagegustot = Val(pagegustot) + 1
       �X�P���ǲέp�Ȯ��ܼ�(3, pagegustot, 1) = Val(pagecardnum(i, 7))
       �X�P���ǲέp�Ȯ��ܼ�(3, pagegustot, 2) = i
    End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(3, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(3, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(3, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(3, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(3, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(3, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(3, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(3, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(3, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(3, o, 2) = h
   End If
  Next
Next
End Sub
Sub ���P�p��Z�����_�ϥΪ�()
Dim i As Integer

For i = 1 To 999
    �Z�����_���P�Ȯɼ�(i, 1) = 0
    �Z�����_���P�Ȯɼ�(i, 2) = 0
Next

�԰��t����.�X�P���ǭp��_�ϥΪ�_�X�P
For i = 1 To pageqlead(1)
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = �X�P���ǲέp�Ȯ��ܼ�(1, i, 2)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2), 9) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2)).Left  '���w�ثeLeft(�y��)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2), 10) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2)).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �Z�����_���P�Ȯɼ�(i, 1) = �Z�����(2, 1, 1)
    �Z�����_���P�Ȯɼ�(i, 2) = �Z�����(2, 1, 2)
    �Z�����_���P�Ȯɼ�(i, 3) = �X�P���ǲέp�Ȯ��ܼ�(1, i, 2)
Next
End Sub
Sub ���P�p��Z�����_�q��()
Dim i As Integer

For i = 1 To 999
    �Z�����_���P�Ȯɼ�(i, 1) = 0
    �Z�����_���P�Ȯɼ�(i, 2) = 0
Next

�԰��t����.�X�P���ǭp��_�q��_�X�P
For i = 1 To pageqlead(2)
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = �X�P���ǲέp�Ȯ��ܼ�(3, i, 2)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2), 9) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2)).Left  '���w�ثeLeft(�y��)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2), 10) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2)).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �Z�����_���P�Ȯɼ�(i, 1) = �Z�����(2, 1, 1)
    �Z�����_���P�Ȯɼ�(i, 2) = �Z�����(2, 1, 2)
    �Z�����_���P�Ȯɼ�(i, 3) = �X�P���ǲέp�Ȯ��ܼ�(3, i, 2)
Next
End Sub
Sub �ޯ໡�����J_�ϥΪ�()
Dim i As Integer, ahmt As String, n As Integer
Dim tmpobj As clsPersonActiveSkill

For n = 1 To 4
    Set tmpobj = New clsPersonActiveSkill
    If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.ActiveDescription 1, n, tmpobj
       FormMainMode.PEAFInterface.ActiveSkillVisable 1, n, False
    Else
        tmpobj.name = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 1)
        
        If VBEPerson(1, ����H����ԤH��(1, 2), 2, 3, 5) = 1 Then
            tmpobj.NameFontSize = 12
        Else
            tmpobj.NameFontSize = VBEPerson(1, ����H����ԤH��(1, 2), 2, 3, n)
        End If
        
        tmpobj.Stage = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 2)
        tmpobj.Distance = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 3)
        tmpobj.card = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 4)
        ahmt = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 5)
        For i = 1 To Len(ahmt)
            If Mid(ahmt, i, 1) = "&" Then
                Mid(ahmt, i, 1) = Chr(10)
            End If
        Next
        tmpobj.Effect = ahmt
        If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 6) <> "" Then
            tmpobj.cardFontSize = Val(VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 6))
        Else
            tmpobj.cardFontSize = 10
        End If
        If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 7) <> "" Then
            tmpobj.EffectFontSize = Val(VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 7))
        Else
            tmpobj.EffectFontSize = 10
        End If
        
        FormMainMode.PEAFInterface.ActiveDescription 1, n, tmpobj
        Set �԰��t����.ActiveSkillObj(1, n) = tmpobj
        FormMainMode.PEAFInterface.ActiveSkillVisable 1, n, True
        If atkingck(1, ����H����ԤH��(1, 2), n, 1) = 1 Then
            �԰��t����.�H���ޯ���O�}�� True, n
        End If
    End If
Next
End Sub
Sub �ޯ໡�����J_�q��()
Dim i As Integer, ahmt As String, n As Integer
Dim tmpobj As clsPersonActiveSkill

For n = 1 To 4
    Set tmpobj = New clsPersonActiveSkill
    If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 1) = "" Then
        FormMainMode.PEAFInterface.ActiveDescription 2, n, tmpobj
        FormMainMode.PEAFInterface.ActiveSkillVisable 2, n, False
    Else
        tmpobj.name = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 1)
        
        If VBEPerson(2, ����H����ԤH��(2, 2), 2, 3, 5) = 1 Then
            tmpobj.NameFontSize = 12
        Else
            tmpobj.NameFontSize = VBEPerson(2, ����H����ԤH��(2, 2), 2, 3, n)
        End If
    
        tmpobj.Stage = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 2)
        tmpobj.Distance = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 3)
        tmpobj.card = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 4)
        ahmt = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 5)
        For i = 1 To Len(ahmt)
            If Mid(ahmt, i, 1) = "&" Then
                Mid(ahmt, i, 1) = Chr(10)
            End If
        Next
        tmpobj.Effect = ahmt
        If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 6) <> "" Then
            tmpobj.cardFontSize = Val(VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 6))
        Else
            tmpobj.cardFontSize = 10
        End If
        If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 7) <> "" Then
            tmpobj.EffectFontSize = Val(VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 7))
        Else
            tmpobj.EffectFontSize = 10
        End If
       
        FormMainMode.PEAFInterface.ActiveDescription 2, n, tmpobj
        Set �԰��t����.ActiveSkillObj(2, n) = tmpobj
        FormMainMode.PEAFInterface.ActiveSkillVisable 2, n, True
    End If
Next
End Sub
Sub ���q�R���ո`�]�w()
Dim i As Integer

If Formsetting.cksemute.Value = 1 Then
    For i = 1 To FormMainMode.cMusicPlayer.UBound
        FormMainMode.cMusicPlayer(i).Mute = True
    Next
Else
    For i = 1 To FormMainMode.cMusicPlayer.UBound
        FormMainMode.cMusicPlayer(i).Mute = False
    Next
End If
End Sub
Sub �ɶ��b_���]()
FormMainMode.timelineout1.X1 = 0
FormMainMode.timelineout2.X2 = 11310
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 1) = 23
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) = 77
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 3) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) = 111
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = 251
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) = 50
FormMainMode.timelineout1.BorderColor = RGB(111, 251, 50)
FormMainMode.timelineout2.BorderColor = RGB(111, 251, 50)
End Sub
Sub �ɶ��b_����()
FormMainMode.trtimeline.Enabled = False
FormMainMode.timelinein1.BorderColor = RGB(0, 0, 0)
FormMainMode.timelinein2.BorderColor = RGB(0, 0, 0)
End Sub
Sub �ɶ��b_����()
FormMainMode.timeup.Visible = False
FormMainMode.timelinein1.Visible = False
FormMainMode.timelinein2.Visible = False
FormMainMode.timelineout1.Visible = False
FormMainMode.timelineout2.Visible = False
End Sub
Sub �ɶ��b_���()
FormMainMode.timeup.Visible = True
FormMainMode.timelinein1.Visible = True
FormMainMode.timelinein2.Visible = True
FormMainMode.timelineout1.Visible = True
FormMainMode.timelineout2.Visible = True
End Sub
Sub ���q����P�_()
If Val(�Y���淾�q�Ȯ��ܼ�(4)) = 1 Then
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
    Case 1
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
           �Y���淾�q�Ȯ��ܼ�(1) = 2
           ���ݮɶ���C(1).Add 14
           FormMainMode.���ݮɶ�.Enabled = True
       Else
           ���ݮɶ���C(1).Add 15
           FormMainMode.���ݮɶ�.Enabled = True
       End If
    Case 2
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
          ���ݮɶ���C(1).Add 15
          FormMainMode.���ݮɶ�.Enabled = True
       Else
          �Y���淾�q�Ȯ��ܼ�(1) = 2
          ���ݮɶ���C(1).Add 13
          FormMainMode.���ݮɶ�.Enabled = True
       End If
    End Select
Else
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
    Case 1
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
          ���ݮɶ���C(1).Add 15
          FormMainMode.���ݮɶ�.Enabled = True
       Else
          �Y���淾�q�Ȯ��ܼ�(1) = 2
          ���ݮɶ���C(1).Add 13
          FormMainMode.���ݮɶ�.Enabled = True
       End If
    Case 2
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
           �Y���淾�q�Ȯ��ܼ�(1) = 2
           ���ݮɶ���C(1).Add 14
           FormMainMode.���ݮɶ�.Enabled = True
       Else
           ���ݮɶ���C(1).Add 15
           FormMainMode.���ݮɶ�.Enabled = True
       End If
    End Select
  End If
End Sub
Sub �q���P_�������P(ByVal Index As Integer)
If pagecardnum(Index, 6) = 1 And pagecardnum(Index, 5) = 2 Then
   pagecardnum(Index, 6) = 2
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp = 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp > 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + defcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
         �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + Val(pagecardnum(Index, 2))
   End If
   '===================
    �ثe��(9) = pagecardnum(Index, 7)
    pagecardnum(Index, 7) = Val(pagecomleadmax(1)) + 1
    pagecomleadmax(1) = Val(pagecomleadmax(1)) + 1
    pageqlead(2) = Val(pageqlead(2)) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) + 1
    pagecardnum(Index, 11) = 2
   '===================�H�U�O�X�P���
    �ثe��(7) = 0
    �԰��t����.�X�P���ǭp��_�q��_�X�P
    FormMainMode.�q���X�P_�X�P���_�a��.Enabled = True
   '=============�H�U�O�P����(�X�P)(�q��)
    �԰��t����.�y�Эp��_�q���X�P
    �P���ʼȮ��ܼ�(3) = Index
    pagecardnum(Index, 9) = FormMainMode.card(Index).Left  '���w�ثeLeft(�y��)
    pagecardnum(Index, 10) = FormMainMode.card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �ثe��(15) = 0
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
   '================�H�U�O��P���
   �ثe��(8) = 0
   �ثe��(17) = 1
   '===================�H�U�O�ƥ�d�ˬd�αҰ�
   If pagecardnum(Index, 1) = a6a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.���|_�q�� Index, pagecardnum(Index, 2)
   End If
   If turnatk = 1 Or turnatk = 2 Then
        If pagecardnum(Index, 1) = a7a Then
            �ƥ�d�O���Ȯɼ�(2, 3) = 1
            �ƥ�d.�A�G�N_�q�� Index, pagecardnum(Index, 2)
        End If
   End If
   If pagecardnum(Index, 1) = a8a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.HP�^�__�q�� Index, pagecardnum(Index, 2)
   End If
   If pagecardnum(Index, 1) = a9a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.�t��_�q�� Index, pagecardnum(Index, 2)
   End If
    '==============================================
    Select Case turnatk
        Case 1
            '===========================���涥�q���J�I(ATK-42/DEF-43)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 42
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 42, 4
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 43
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 43, 4
            '============================
        Case 2
            '===========================���涥�q���J�I(ATK-42/DEF-43)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 42
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 42, 4
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 43
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 43, 4
            '============================
        Case 3
            '===========================���涥�q���J�I(44)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 44
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 44, 3
            '============================
    End Select
    �԰��t����.��q��s���
End If
End Sub
Sub �q���P_�������P_�~(ByVal Index As Integer)
If pagecardnum(Index, 6) = 2 And pagecardnum(Index, 5) = 2 Then
   pagecardnum(Index, 6) = 1
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 2))
      End If
      If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
          �������m��l�`��(4) = 0
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 2))
      End If
      If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
          �������m��l�`��(4) = 0
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - Val(pagecardnum(Index, 2))
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2))
         �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - Val(pagecardnum(Index, 2))
   End If
   '================
   �ثe��(9) = pagecardnum(Index, 7)
    pagecardnum(Index, 7) = Val(pagecomleadmax(0)) + 1
    pagecomleadmax(0) = Val(pagecomleadmax(0)) + 1
    pageqlead(2) = Val(pageqlead(2)) - 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
    pagecardnum(Index, 11) = 0
   '=============�H�U�O�P����(�^�P)(�q��)
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = Index
    pagecardnum(Index, 9) = FormMainMode.card(Index).Left  '���w�ثeLeft(�y��)
    pagecardnum(Index, 10) = FormMainMode.card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.���εP�ܭI��
    �ثe��(15) = 0
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
   '================�H�U�O�X�P���
   �ثe��(7) = 0
   �԰��t����.�X�P���ǭp��_�q��_�X�P
   FormMainMode.�q���X�P_�X�P���_�a�k.Enabled = True
   '=====================�H�U�O�ޯ��ˬd�αҰ�
    If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards") <> 0 Then
        vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards")) = 2 '(���q2)
    End If
    '==============================================
    Select Case turnatk
        Case 1
            '===========================���涥�q���J�I(ATK-42/DEF-43)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 42
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 42, 4
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 43
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 43, 4
            '============================
        Case 2
            '===========================���涥�q���J�I(ATK-42/DEF-43)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 42
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 42, 4
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 43
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 43, 4
            '============================
        Case 3
            '===========================���涥�q���J�I(44)
            ReDim VBEStageNum(0 To 1) As Integer
            VBEStageNum(0) = 44
            VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
            ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 44, 3
            '============================
    End Select
    �԰��t����.��q��s���
End If
End Sub
Sub �q���P_������P_�~(ByVal Index As Integer)
Dim uspce As String, uspme As String

uspce = pagecardnum(Index, 1)
uspme = pagecardnum(Index, 2)
pagecardnum(Index, 1) = pagecardnum(Index, 3)
pagecardnum(Index, 2) = pagecardnum(Index, 4)
pagecardnum(Index, 3) = uspce
pagecardnum(Index, 4) = uspme
�@��t����.���ļ��� 3
If pageonin(Index) = 1 Then
   pageonin(Index) = 2
Else
   pageonin(Index) = 1
End If
FormMainMode.card(Index).CardRotationType = pageonin(Index)

   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + pagecardnum(Index, 2)
      If turnatk = 2 And movecp = 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + pagecardnum(Index, 2)
      If turnatk = 2 And movecp > 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + pagecardnum(Index, 2)
      If turnatk = 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + defcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
         �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + pagecardnum(Index, 2)
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + pagecardnum(Index, 2)
   End If
'======================================
   If pagecardnum(Index, 3) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - pagecardnum(Index, 4)
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 4))
      End If
      If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
          �������m��l�`��(4) = 0
      End If
   End If
   If pagecardnum(Index, 3) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - pagecardnum(Index, 4)
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 4))
      End If
      If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
          �������m��l�`��(4) = 0
      End If
   End If
   If pagecardnum(Index, 3) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - pagecardnum(Index, 4)
      If turnatk = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - pagecardnum(Index, 4)
   End If
   If pagecardnum(Index, 3) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - pagecardnum(Index, 4)
   End If
'==============================================
Select Case turnatk
    Case 1
        '===========================���涥�q���J�I(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 43, 4
        '============================
    Case 2
        '===========================���涥�q���J�I(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 43, 4
        '============================
    Case 3
        '===========================���涥�q���J�I(44)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 44
        VBEStageNum(1) = -2 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 44, 3
        '============================
End Select
�԰��t����.��q��s���
End Sub
Sub ��ƹs����P�_()
    Dim ustruenum As Integer, comtruenum As Integer
    Dim p As Integer, i As Integer, j As Integer
    '==�L������ܡA�ݦۦ��Y��
    For p = 1 To �Y���淾�q�Ȯ��ܼ�(9)
       Randomize Timer
       i = Int(Rnd() * 6) + 1
       If i = 1 Or i = 6 Then ustruenum = ustruenum + 1
    Next
    For p = 1 To �Y���淾�q�Ȯ��ܼ�(10)
        Randomize Timer
        j = Int(Rnd() * 6) + 1
        If j = 1 Or j = 6 Then comtruenum = comtruenum + 1
    Next
    If �O�_�t�Τ��� = True Then
        �Y���淾�q�Ȯ��ܼ�(5) = ustruenum
        �Y���淾�q�Ȯ��ܼ�(6) = comtruenum
    Else
        Vss_BattleStartDiceNum(3) = ustruenum
        Vss_BattleStartDiceNum(4) = comtruenum
    End If
End Sub
Sub �Y�������()
If ��ƹs�ˬd��(1) = False And ��ƹs�ˬd��(2) = False Then
     If moveturn = 1 Then
       Select Case �Y���淾�q�Ȯ��ܼ�(1)
          Case 1
              FormMainMode.PEAFDiceInterface.DiceATKMode = 1
              FormMainMode.PEAFDiceInterface.DiceInputMode = 2
              FormMainMode.PEAFDiceInterface.diceusTotal = �Y���淾�q�Ȯ��ܼ�(9)
              FormMainMode.PEAFDiceInterface.dicecomTotal = �Y���淾�q�Ȯ��ܼ�(10)
              FormMainMode.PEAFDiceInterface.PersonImage = �԰��Y�뤶���H����ø�ϸ��|������(1)
              FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, ����H����ԤH��(1, 2), 2, 2, 5))
              FormMainMode.PEAFDiceInterface.ZOrder
              �԰��t����.�Y��ɦ�q�������h���
              FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
              FormMainMode.PEAFDiceInterface.DiceStart = True
          Case 2
              FormMainMode.PEAFDiceInterface.DiceATKMode = 2
              FormMainMode.PEAFDiceInterface.DiceInputMode = 2
              FormMainMode.PEAFDiceInterface.diceusTotal = �Y���淾�q�Ȯ��ܼ�(9)
              FormMainMode.PEAFDiceInterface.dicecomTotal = �Y���淾�q�Ȯ��ܼ�(10)
              FormMainMode.PEAFDiceInterface.PersonImage = �԰��Y�뤶���H����ø�ϸ��|������(2)
              FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, ����H����ԤH��(2, 2), 2, 2, 5))
              FormMainMode.PEAFDiceInterface.ZOrder
              �԰��t����.�Y��ɦ�q�������h���
              FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
              FormMainMode.PEAFDiceInterface.DiceStart = True
       End Select
     ElseIf moveturn = 2 Then
        Select Case �Y���淾�q�Ȯ��ܼ�(1)
           Case 1
              FormMainMode.PEAFDiceInterface.DiceATKMode = 2
              FormMainMode.PEAFDiceInterface.DiceInputMode = 2
              FormMainMode.PEAFDiceInterface.diceusTotal = �Y���淾�q�Ȯ��ܼ�(9)
              FormMainMode.PEAFDiceInterface.dicecomTotal = �Y���淾�q�Ȯ��ܼ�(10)
              FormMainMode.PEAFDiceInterface.PersonImage = �԰��Y�뤶���H����ø�ϸ��|������(2)
              FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, ����H����ԤH��(2, 2), 2, 2, 5))
              FormMainMode.PEAFDiceInterface.ZOrder
              �԰��t����.�Y��ɦ�q�������h���
              FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
              FormMainMode.PEAFDiceInterface.DiceStart = True
           Case 2
              FormMainMode.PEAFDiceInterface.DiceATKMode = 1
              FormMainMode.PEAFDiceInterface.DiceInputMode = 2
              FormMainMode.PEAFDiceInterface.diceusTotal = �Y���淾�q�Ȯ��ܼ�(9)
              FormMainMode.PEAFDiceInterface.dicecomTotal = �Y���淾�q�Ȯ��ܼ�(10)
              FormMainMode.PEAFDiceInterface.PersonImage = �԰��Y�뤶���H����ø�ϸ��|������(1)
              FormMainMode.PEAFDiceInterface.PersonImageLeftZero = CBool(VBEPerson(1, ����H����ԤH��(1, 2), 2, 2, 5))
              FormMainMode.PEAFDiceInterface.ZOrder
              �԰��t����.�Y��ɦ�q�������h���
              FormMainMode.PEAFDiceInterface.dicevoice = Formsetting.seve.Caption
              FormMainMode.PEAFDiceInterface.DiceStart = True
         End Select
     End If
Else
   '========================
     �ثe��(26) = 0
    '========================
    �԰��t����.��ƹs����P�_
 End If
End Sub
Sub �Y��ɦ�q�������h���()
FormMainMode.PEAFbloodbackimage1.ZOrder
FormMainMode.PEAFbloodbackimage2.ZOrder
FormMainMode.bloodnumus1.ZOrder
FormMainMode.bloodnumus2.ZOrder
FormMainMode.bloodnumcom1.ZOrder
FormMainMode.bloodnumcom2.ZOrder
End Sub
Sub �Y�����P�_()
If �O�_�t�Τ��� = True Then
    If ��ƹs�ˬd��(1) = False And ��ƹs�ˬd��(2) = False Then
        �Y���淾�q�Ȯ��ܼ�(5) = Val(FormMainMode.PEAFDiceInterface.diceusTrue)
        �Y���淾�q�Ȯ��ܼ�(6) = Val(FormMainMode.PEAFDiceInterface.dicecomTrue)
    End If
    FormMainMode.��l���槹�Ұ�.Enabled = True
Else
    If ��ƹs�ˬd��(1) = False And ��ƹs�ˬd��(2) = False Then
        Vss_BattleStartDiceNum(3) = Val(FormMainMode.PEAFDiceInterface.diceusTrue)
        Vss_BattleStartDiceNum(4) = Val(FormMainMode.PEAFDiceInterface.dicecomTrue)
    End If
End If
'=====================================================
If Val(�Y���淾�q�Ȯ��ܼ�(4)) = 1 Then
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
        Case 1
           GoTo usatkcom
        Case 2
           GoTo comatkus
    End Select
  Else
    Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
        Case 1
           GoTo comatkus
        Case 2
           GoTo usatkcom
     End Select
End If
'==========================================
Exit Sub
usatkcom:
    If �O�_�t�Τ��� = True Then
        �Y���淾�q�Ȯ��ܼ�(2) = �Y���淾�q�Ȯ��ܼ�(5) - �Y���淾�q�Ȯ��ܼ�(6)
        �Y���淾�q�Ȯ��ܼ�(3) = 2
    Else
        Vss_BattleStartDiceNum(5) = Vss_BattleStartDiceNum(3) - Vss_BattleStartDiceNum(4)
    End If
'==========================================
Exit Sub
comatkus:
    If �O�_�t�Τ��� = True Then
        �Y���淾�q�Ȯ��ܼ�(2) = �Y���淾�q�Ȯ��ܼ�(6) - �Y���淾�q�Ȯ��ܼ�(5)
        �Y���淾�q�Ȯ��ܼ�(3) = 1
    Else
        Vss_BattleStartDiceNum(5) = Vss_BattleStartDiceNum(4) - Vss_BattleStartDiceNum(3)
    End If
End Sub
Sub ����HP�ˬd()
Dim inp As Integer 'RND�Ȯ��ܼ�
Dim person(1 To 2) As Integer
Erase �H�������ˬd�Ȯ��ܼ�
If livecom(����H����ԤH��(2, 2)) <= 0 Then
   �H�������ˬd�Ȯ��ܼ�(3) = 1
   If livecom(����ݾ��H��������(2, 2)) > 0 Then
       person(2) = 2
       �洫��������Ȯ��ܼ�(2) = 1
   ElseIf livecom(����ݾ��H��������(2, 3)) > 0 Then
       �洫��������Ȯ��ܼ�(2) = 1
       person(2) = 2
   Else
       person(2) = 1
   End If
End If
If liveus(����H����ԤH��(1, 2)) <= 0 Then
   �H�������ˬd�Ȯ��ܼ�(2) = 1
   If liveus(����ݾ��H��������(1, 2)) > 0 Or liveus(����ݾ��H��������(1, 3)) > 0 Then
       person(1) = 2
       �洫��������Ȯ��ܼ�(1) = 1
   Else
       person(1) = 1
   End If
End If

If person(1) = 2 Or person(2) = 2 Then
   ���ݮɶ���C(1).Add 21
   FormMainMode.�H�������ˬd.Enabled = True
   Exit Sub
ElseIf person(1) = 0 And person(2) = 1 Then
   �԰��Ҧ��ӱѬ����� = 1
   ���ݮɶ���C(1).Add 36
   FormMainMode.�H�������ˬd.Enabled = True
ElseIf person(1) = 1 And person(2) = 0 Then
   ���ݮɶ���C(1).Add 36
   �԰��Ҧ��ӱѬ����� = 2
   FormMainMode.�H�������ˬd.Enabled = True
ElseIf person(1) = 1 And person(2) = 1 Then
   Randomize
   inp = Int(Rnd() * 2) + 1
   Select Case inp
       Case 1
           �԰��Ҧ��ӱѬ����� = 1
           ���ݮɶ���C(1).Add 36
           FormMainMode.�H�������ˬd.Enabled = True
       Case 2
           �԰��Ҧ��ӱѬ����� = 2
           ���ݮɶ���C(1).Add 36
           FormMainMode.�H�������ˬd.Enabled = True
    End Select
End If

If FormMainMode.�H�������ˬd.Enabled = False Then
  Select Case HP�ˬd���q��
     Case 1
       '----------�H�U�����q�~����]���ʶ��q3�^
        ���ݮɶ���C(1).Add 4
        FormMainMode.���ݮɶ�.Enabled = True
     Case 2
          ���ݮɶ���C(1).Add 11
          FormMainMode.���ݮɶ�.Enabled = True
     Case 3
        �԰��t����.���q����P�_
     Case 4
        FormMainMode.NextTurn_���q2.Enabled = True
  End Select
End If
End Sub
Function ����HP�ˬd_�����^�X�ˬd() As Boolean
Dim num(1 To 2) As Integer '��ܤH���Ȯ��ܼ�
Dim i As Integer

If BattleTurn >= Val(Formsetting.ckendturnnum.Text) And Formsetting.ckendturn.Value = 1 Then
        ����HP�ˬd_�����^�X�ˬd = True
        '==============
        For i = 1 To 3
            If liveus(����ݾ��H��������(1, i)) > 0 Then
                num(1) = Val(num(1)) + Val(liveus(����ݾ��H��������(1, i)))
            End If
            If livecom(����ݾ��H��������(2, i)) > 0 Then
                num(2) = Val(num(2)) + Val(livecom(����ݾ��H��������(2, i)))
            End If
         Next
        '==============
        If num(1) > num(2) Then
           �԰��Ҧ��ӱѬ����� = 1
           FormMainMode.trend.Enabled = True
        ElseIf num(1) < num(2) Then
           �԰��Ҧ��ӱѬ����� = 2
           FormMainMode.trend.Enabled = True
        ElseIf num(1) = num(2) Then
            '�L����ѥ_
            �԰��Ҧ��ӱѬ����� = 2
            FormMainMode.trend.Enabled = True
        End If
Else
     ����HP�ˬd_�����^�X�ˬd = False
End If
End Function

Sub checkpage()
Dim i As Integer

For i = 1 To �ثe��(11)
  If �ثe��(10) = 1 Then
   FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
   pageqlead(1) = Val(pageqlead(1)) - 1
  ElseIf �ثe��(10) = 2 Then
   FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
   pageqlead(2) = Val(pageqlead(2)) - 1
  End If
Next
End Sub
Sub chkcom()
If goicheck(2) = 0 Then
  If atkingpagetot(2, 1) > 0 And movecp = 1 Then
    �������m��l�`��(2) = �������m��l�`��(2) + atkcom(����H����ԤH��(2, 2))
    �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
    goicheck(2) = 1
  ElseIf atkingpagetot(2, 5) > 0 And movecp > 1 Then
    �������m��l�`��(2) = �������m��l�`��(2) + atkcom(����H����ԤH��(2, 2))
    �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
    goicheck(2) = 1
  End If
End If
End Sub
Sub chkdef()
If goidefus = 0 Then
 �������m��l�`��(1) = �������m��l�`��(1) + defus(����H����ԤH��(1, 2))
 �������m��l�`��(3) = �������m��l�`��(3) + defus(����H����ԤH��(1, 2))
 FormMainMode.��ܦC1.goi1 = Val(FormMainMode.��ܦC1.goi1) + defus(����H����ԤH��(1, 2))
 goidefus = 1
End If
End Sub
Sub chkdefcom()
If chkcomck = 0 Then
 �������m��l�`��(2) = �������m��l�`��(2) + defcom(����H����ԤH��(2, 2))
 �������m��l�`��(4) = �������m��l�`��(4) + defcom(����H����ԤH��(2, 2))
 FormMainMode.��ܦC1.goi2 = Val(FormMainMode.��ܦC1.goi2) + defcom(����H����ԤH��(2, 2))
 chkcomck = 1
End If
End Sub
Sub chkus1()
If goicheck(1) = 0 Then
 If atkingpagetot(1, 1) > 0 Then
   �������m��l�`��(1) = �������m��l�`��(1) + atkus(����H����ԤH��(1, 2))
   �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
   goicheck(1) = 1
  End If
End If
End Sub
Sub cleanatkingpagetot()
Dim i As Integer, j As Integer

For i = 1 To 2
     For j = 1 To 5
        atkingpagetot(i, j) = 0
     Next
Next
End Sub
Sub comatk1()
Dim a As Integer
Dim cspce As String, cspme As String

For a = 1 To ���εP����d�����j������(1)
  If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 Then
     If pagecardnum(a, 1) = a1a Then
       pagecardnum(a, 11) = 1
     ElseIf pagecardnum(a, 3) = a1a Then
       cspce = pagecardnum(a, 1)
       cspme = pagecardnum(a, 2)
       pagecardnum(a, 1) = pagecardnum(a, 3)
       pagecardnum(a, 2) = pagecardnum(a, 4)
       pagecardnum(a, 3) = cspce
       pagecardnum(a, 4) = cspme
       If pageonin(a) = 2 Then
          pageonin(a) = 1
       Else
          pageonin(a) = 2
       End If
       pagecardnum(a, 11) = 1
     End If
  End If
Next
End Sub
Sub comatk2()
Dim j As Integer
Dim cspce As String, cspme As String

For j = 1 To ���εP����d�����j������(1)
  If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
     If pagecardnum(j, 1) = a5a Then
       pagecardnum(j, 11) = 1
     ElseIf pagecardnum(j, 3) = a5a Then
       cspce = pagecardnum(j, 1)
       cspme = pagecardnum(j, 2)
       pagecardnum(j, 1) = pagecardnum(j, 3)
       pagecardnum(j, 2) = pagecardnum(j, 4)
       pagecardnum(j, 3) = cspce
       pagecardnum(j, 4) = cspme
       If pageonin(j) = 2 Then
          pageonin(j) = 1
       Else
          pageonin(j) = 2
       End If
       pagecardnum(j, 11) = 1
     End If
  End If
Next
End Sub
Sub comatk_���z��AI�޾ɵ{��_�W�X�P�i��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal choose As Integer)
Dim werstr As String, werbo As Boolean
Dim a As Integer, k As Integer
Dim cspce As String, cspme As String

If movecpre = 1 And turn = 1 Then
   werstr = a1a
ElseIf movecpre > 1 And turn = 1 Then
   werstr = a5a
ElseIf turn = 2 Then
   werstr = a2a
End If
'=================================
For a = 1 To ���εP����d�����j������(1)
    werbo = False
    For k = 1 To UBound(cardAInumOvertenrecord)
        If a = cardAInumOvertenrecord(k) Then
            werbo = True
        End If
    Next
    If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 And werbo = False Then
            If pagecardnum(a, 1) = werstr Then
              pagecardnum(a, 11) = 1
            ElseIf pagecardnum(a, 3) = werstr Then
              cspce = pagecardnum(a, 1)
              cspme = pagecardnum(a, 2)
              pagecardnum(a, 1) = pagecardnum(a, 3)
              pagecardnum(a, 2) = pagecardnum(a, 4)
              pagecardnum(a, 3) = cspce
              pagecardnum(a, 4) = cspme
              If pageonin(a) = 2 Then
                 pageonin(a) = 1
              Else
                 pageonin(a) = 2
              End If
              pagecardnum(a, 11) = 1
            End If
            If choose = 1 And pagecardnum(a, 11) = 0 Then
                pagecardnum(a, 11) = 1
            End If
    End If
Next
End Sub
Sub moveatkin()
Dim j As Integer
Dim cspce As String, cspme As String

Do
    For j = ���εP����d�����j������(2) + 1 To ���εP����d�����j������(4)
      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
         If pagecardnum(j, 1) = a3a And pagecardnum(j, 3) = a3a Then '���ʳ歱�ƥ�d�u��
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            �ثe��(25) = �ثe��(25) + Val(pagecardnum(j, 2))
         End If
         If �ثe��(25) >= 2 Then Exit Do
      End If
    Next
    For j = 1 To ���εP����d�����j������(1)
      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
         If pagecardnum(j, 1) = a3a Then
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            �ثe��(25) = �ثe��(25) + 1
         ElseIf pagecardnum(j, 3) = a3a Then
           cspce = pagecardnum(j, 1)
           cspme = pagecardnum(j, 2)
           pagecardnum(j, 1) = pagecardnum(j, 3)
           pagecardnum(j, 2) = pagecardnum(j, 4)
           pagecardnum(j, 3) = cspce
           pagecardnum(j, 4) = cspme
           If pageonin(j) = 2 Then
              pageonin(j) = 1
           Else
              pageonin(j) = 2
           End If
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            �ثe��(25) = �ثe��(25) + Val(pagecardnum(j, 2))
         End If
         If �ثe��(25) >= 2 Then Exit Do
      End If
    Next
    Exit Do
Loop
'movecheckcom = movecom
End Sub
Sub movetnus()
�԰��t����.�s���T�� "�A���D���v�C"
FormMainMode.move3.Picture = LoadPicture(app_path & "gif\system\atk1.gif")
FormMainMode.move4.Picture = LoadPicture(app_path & "gif\system\def1.gif")
FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\system\atk2.gif")
FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\system\def2.gif")
moveturn = 1
FormMainMode.cnmove2.Visible = False
�Y���淾�q�Ȯ��ܼ�(1) = 1
End Sub
Sub movetncom()
�԰��t����.�s���T�� "��観�D���v�C"
FormMainMode.move3.Picture = LoadPicture(app_path & "gif\system\def1.gif")
FormMainMode.move4.Picture = LoadPicture(app_path & "gif\system\atk1.gif")
FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\system\def2.gif")
FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\system\atk2.gif")
moveturn = 2
FormMainMode.cnmove2.Visible = False
�Y���淾�q�Ȯ��ܼ�(1) = 1
End Sub
Sub �H���洫_�ϥΪ�_���w�洫(ByVal num As Integer)
Dim ae As Integer, n As Integer, i As Integer, ahmt As String
Dim tmpobj As clsPersonActiveSkill
'=======================
ReDim VBEStageNum(0 To 3) As Integer
VBEStageNum(0) = 41
VBEStageNum(1) = -1 '����ĪG��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = 1 '�洫�e�H���s��
VBEStageNum(3) = num '�洫��H���s��
'===========================���涥�q���J�I(41)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 41, 1
'============================
FormMainMode.personusminijpg.�p�H������ = True
Do Until FormMainMode.personusminijpg.�p�H������ = False
    DoEvents
Loop
'=======================
ae = ����H����ԤH��(1, 2)
����H����ԤH��(1, 2) = ����ݾ��H��������(1, num)
����ݾ��H��������(1, 1) = ����H����ԤH��(1, 2)
����ݾ��H��������(1, num) = ae
FormMainMode.PEAFpersoncardus(����ݾ��H��������(1, num)).Left = 2520 * (num - 1)
FormMainMode.PEAFpersoncardus(����ݾ��H��������(1, num)).Visible = True
FormMainMode.cardus(����ݾ��H��������(1, num)).Visible = False

FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).Left = 0
FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).Visible = False
FormMainMode.cardus(����H����ԤH��(1, 2)).Left = 0
FormMainMode.cardus(����H����ԤH��(1, 2)).Top = 6240
FormMainMode.cardus(����H����ԤH��(1, 2)).ZOrder
FormMainMode.cardus(����H����ԤH��(1, 2)).Visible = True
'=======================
�԰��t����.�ޯ໡�����J_�ϥΪ�
FormMainMode.PEAFInterface.Passive_�ޯ�@������] 1
For n = 5 To 8
    If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ����� n - 4
    Else
       FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�W�� VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 1), n - 4
       FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ���� n - 4
       '=======================
       If atkingck(1, ����H����ԤH��(1, 2), n, 1) = 1 Then
           FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�O�o�G n - 4
       End If
    End If
Next
If �H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 1) <> "" And Val(�H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 2)) = 1 Then
    FormMainMode.personusminijpg.�p�H���Ϥ� = �H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 4)
    FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = �H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 5)
    FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = �H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 6)
    FormMainMode.personusminijpg.�p�H���v�lLeft = Val(�H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 7))
    FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(�H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 8))
    �԰��Y�뤶���H����ø�ϸ��|������(1) = �H����ڪ��A��Ʈw(1, ����H����ԤH��(1, 2), 3)
Else
    FormMainMode.personusminijpg.�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 1)
    FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 2)
    FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 4)
    FormMainMode.personusminijpg.�p�H���v�lLeft = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 5))
    FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 6))
    �԰��Y�뤶���H����ø�ϸ��|������(1) = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 3)
End If
FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -(FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width)
'--------------------------�p��s�Z�����(HP���)
�Z�����(1, 1, 1) = 5295 \ liveusmax(����H����ԤH��(1, 2))
FormMainMode.bloodlineout1.Width = (�Z�����(1, 1, 1) * liveus(����H����ԤH��(1, 2)))
FormMainMode.bloodnumus1.Caption = liveus(����H����ԤH��(1, 2))
FormMainMode.bloodnumus2.Caption = liveusmax(����H����ԤH��(1, 2))
'========================
����ʧ@_�Z���ܧ� movecp, False
'========================
For i = 1 To 4
    �԰��t����.�H���ޯ���O�}�� False, i
Next
'=============================
FormMainMode.personusminijpg.�p�H����{ = True
Do Until FormMainMode.personusminijpg.�p�H����{ = False
    DoEvents
Loop

End Sub

Sub �H���洫_�q��_���w�洫(ByVal num As Integer)
Dim ae As Integer, n As Integer
'=======================
ReDim VBEStageNum(0 To 3) As Integer
VBEStageNum(0) = 41
VBEStageNum(1) = -2 '����ĪG��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = 1 '�洫�e�H���s��
VBEStageNum(3) = num '�洫��H���s��
'===========================���涥�q���J�I(41)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 41, 1
'============================
FormMainMode.personcomminijpg.�p�H������ = True
Do Until FormMainMode.personcomminijpg.�p�H������ = False
    DoEvents
Loop
'=======================
ae = ����H����ԤH��(2, 2)
����H����ԤH��(2, 2) = ����ݾ��H��������(2, num)
����ݾ��H��������(2, num) = ae
����ݾ��H��������(2, 1) = ����H����ԤH��(2, 2)
FormMainMode.PEAFpersoncardcom(����ݾ��H��������(2, num)).Left = 2520 * (num - 1)
FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).Left = 0
'=======================
�԰��t����.�ޯ໡�����J_�q��
FormMainMode.PEAFInterface.Passive_�ޯ�@������] 2
For n = 5 To 8
    If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.Passive_�q��_�ޯ����� n - 4
    Else
       FormMainMode.PEAFInterface.Passive_�q��_�ޯ�W�� VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 1), n - 4
       FormMainMode.PEAFInterface.Passive_�q��_�ޯ���� n - 4
       '=======================
       If atkingck(2, ����H����ԤH��(2, 2), n, 1) = 1 Then
           FormMainMode.PEAFInterface.Passive_�q��_�ޯ�O�o�G n - 4
       End If
    End If
Next
'====================
If �H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 1) <> "" And Val(�H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 2)) = 1 Then
    FormMainMode.personcomminijpg.�p�H���Ϥ� = �H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 4)
    FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = �H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 5)
    FormMainMode.��ܦC1.�q����p�H���Ϥ� = �H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 6)
    FormMainMode.personcomminijpg.�p�H���v�lLeft = Val(�H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 7))
    FormMainMode.personcomminijpg.�p�H���v�ltop�t = Val(�H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 8))
    �԰��Y�뤶���H����ø�ϸ��|������(2) = �H����ڪ��A��Ʈw(2, ����H����ԤH��(2, 2), 3)
Else
    FormMainMode.personcomminijpg.�p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 1)
    FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 2)
    FormMainMode.��ܦC1.�q����p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 4)
    FormMainMode.personcomminijpg.�p�H���v�lLeft = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 5)
    FormMainMode.personcomminijpg.�p�H���v�ltop�t = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 6)
    �԰��Y�뤶���H����ø�ϸ��|������(2) = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 3)
End If
FormMainMode.��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
�԰��t����.��������p�d��T�g�J 2, ����H����ԤH��(2, 2)
�԰��t����.PersonCardShowOnMode(2, ����H����ԤH��(2, 2)) = True
FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).ShowOnMode = True
FormMainMode.cardcom(����H����ԤH��(2, 2)).ShowOnMode = True
'--------------------------�p��s�Z�����(HP���)
�Z�����(1, 2, 1) = (11340 - 6060) \ livecommax(����H����ԤH��(2, 2))
FormMainMode.bloodlineout2.Left = 11340 - (�Z�����(1, 2, 1) * livecom(����H����ԤH��(2, 2)))
FormMainMode.bloodnumcom1.Caption = livecom(����H����ԤH��(2, 2))
FormMainMode.bloodnumcom2.Caption = livecommax(����H����ԤH��(2, 2))
'==============================
����ʧ@_�Z���ܧ� movecp, False
'=============================
FormMainMode.personcomminijpg.�p�H����{ = True
Do Until FormMainMode.personcomminijpg.�p�H����{ = False
    DoEvents
Loop
'=======================
End Sub
Sub ����ʧ@_�洫�H������_�ϥΪ�_��l()
Dim i As Integer, k As Integer
Dim ne As Integer
Dim numNow As Integer, obj As clsStatus

For i = 2 To 3
   Formchangeperson.card(i - 1).���`���A�����]
   Formchangeperson.card(i - 1).CardBack�����]
   Formchangeperson.card(i - 1).CardMain_����Ϥ� = VBEPerson(1, ����ݾ��H��������(1, i), 1, 5, 5)
   Formchangeperson.card(i - 1).CardMain_����HP = liveus(����ݾ��H��������(1, i))
   Formchangeperson.card(i - 1).CardMain_����HPMAX = liveusmax(����ݾ��H��������(1, i))
   Formchangeperson.card(i - 1).CardMain_����ATK = atkus(����ݾ��H��������(1, i))
   Formchangeperson.card(i - 1).CardMain_����DEF = defus(����ݾ��H��������(1, i))
   Formchangeperson.card(i - 1).CardMain_�O�_���s�˦���T = CBool(Val(VBEPerson(1, ����ݾ��H��������(1, i), 1, 3, 5)) = 1)
Next
�԰��t����.�ޯ໡�����J_�H���d���I��_�洫����

ne = 1
For k = 2 To 3
    numNow = 1
    For Each obj In �H�����`���A�C��(1, ����ݾ��H��������(1, k))
        Formchangeperson.card(ne).��ﲧ�`���A��� numNow, obj.ImagePath, obj.Value, obj.Total, True
        numNow = numNow + 1
        If numNow > 14 Then Exit For
    Next
    ne = ne + 1
Next

�洫��������Ȯ��ܼ�(1) = 0
For i = 2 To 3
    Formchangeperson.card(i - 1).MusicPlayerObj = FormMainMode.cMusicPlayer(9)
    Formchangeperson.card(i - 1).ShowOnMode = True
Next
If Formsetting.chkusenewaipersonauto.Value = 1 Then
    Formchangeperson.�ϥΪ̤贼�z��AI_�۰ʱ����H.Enabled = True
End If
Formchangeperson.Left = FormMainMode.Left + 2430
Formchangeperson.Top = FormMainMode.Top + 1655
Formchangeperson.Show 1
End Sub
Sub ����ʧ@_�洫�H������_�q��_��l()
Select Case �洫��������Ȯ��ܼ�(2)
    Case 1
       �洫��������Ȯ��ܼ�(2) = 0
       ���ݮɶ���C(1).Add 18
       FormMainMode.���ݮɶ�.Enabled = True
    Case 0
       ���ݮɶ���C(1).Add 19
       FormMainMode.���ݮɶ�.Enabled = True
End Select

End Sub
Sub ����ʧ@_�洫�H������_�q��_�洫()
If livecom(����ݾ��H��������(2, 2)) > 0 Then
       �H���洫_�q��_���w�洫 2
ElseIf livecom(����ݾ��H��������(2, 3)) > 0 Then
       �H���洫_�q��_���w�洫 3
End If
����ʧ@_�洫�H������_��������
End Sub
Sub ����ʧ@_�洫�H������_��l()
If (�洫��������Ȯ��ܼ�(1) = 1 Or �洫��������Ȯ��ܼ�(2) = 1) And �洫��������Ȯ��ܼ�(3) = 0 Then
    turnatk = 6
    ���q���A�� = 5
    �԰��t����.�ɶ��b_���]
    FormMainMode.��ܦC1.��ܦC�Ϥ� = App.Path & "\gif\system\linechange.png"
    FormMainMode.��ܦC1.Visible = True
    FormMainMode.��ܦC1.goi1��� = False
    FormMainMode.��ܦC1.goi2��� = False
    �԰��t����.�ɶ��b_���
    FormMainMode.trtimeline.Enabled = True
    �p�H���Y�����ʤ�V��(1) = 2
    �p�H���Y�����ʤ�V��(2) = 2
    FormMainMode.�p�H���Y������_�ϥΪ�.Enabled = True
    FormMainMode.�p�H���Y������_�q��.Enabled = True
    �洫��������Ȯ��ܼ�(3) = 1
    FormMainMode.��ܦC1.���ʶ��q��ܭ� = 0
    FormMainMode.��ܦC1.���ʶ��q����� = False
End If
If �洫��������Ȯ��ܼ�(1) = 1 Then
    ����ʧ@_�洫�H������_�ϥΪ�_��l
ElseIf �洫��������Ȯ��ܼ�(2) = 1 Then
    ����ʧ@_�洫�H������_�q��_��l
End If
End Sub
Sub ����ʧ@_���ʶ��q��ܰ���()
'===========�洫������
If �洫��������Ȯ��ܼ�(1) = 1 Or �洫��������Ȯ��ܼ�(2) = 1 Then
    ����ʧ@_�洫�H������_��l
Else
    �洫��������Ȯ��ܼ�(3) = 0
    ���ݮɶ���C(1).Add 17
    FormMainMode.���ݮɶ�.Enabled = True
End If
End Sub
Sub ����ʧ@_�H�����`�洫���q��ܰ���()
If �洫��������Ȯ��ܼ�(1) = 1 Or �洫��������Ȯ��ܼ�(2) = 1 Then
    ����ʧ@_�洫�H������_��l
Else
    �洫��������Ȯ��ܼ�(3) = 0
    ���ݮɶ���C(1).Add 20
    FormMainMode.���ݮɶ�.Enabled = True
End If
End Sub
Sub ����ʧ@_�洫�H������_��������()
   Formchangeperson.Hide
   �԰��t����.�ɶ��b_����
   Select Case �洫��������Ȯ��ܼ�(4)
      Case 1
         ����ʧ@_���ʶ��q��ܰ���
      Case 2
         ����ʧ@_�H�����`�洫���q��ܰ���
    End Select
End Sub
Sub �ƥ�d�B�z_���w_�ϥΪ̤�()
Dim kp(1 To 18)  As Integer '�ƥ�d�аO�Ȯɼ�
Dim m As Integer, km As Integer, i As Integer
If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgreus.Value = 0 Then
    Do
        Randomize
        m = Int(Rnd() * 18) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(1, km, 1) = Formsetting.personus(m).Text
            pageeventnum(1, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personus(m).Text, 2)
        End If
    Loop Until km >= 18
ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
    Do
        Randomize
        m = Int(Rnd() * 6) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(1, km, 1) = Formsetting.personus(m).Text
            pageeventnum(1, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personus(m).Text, 2)
        End If
    Loop Until km >= 6
    For i = 7 To 12
        pageeventnum(1, i, 1) = Formsetting.personus(i).Text
        pageeventnum(1, i, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).Text, 2)
    Next
End If
End Sub
Sub �ƥ�d�B�z_���w_�q����()
Dim kp(1 To 18)  As Integer '�ƥ�d�аO�Ȯɼ�
Dim m As Integer, km As Integer, i As Integer

If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgrecom.Value = 0 Then
    Do
        Randomize
        m = Int(Rnd() * 18) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(2, km, 1) = Formsetting.personcom(m).Text
            pageeventnum(2, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(m).Text, 2)
        End If
    Loop Until km >= 18
ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
    Do
        Randomize
        m = Int(Rnd() * 6) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(2, km, 1) = Formsetting.personcom(m).Text
            pageeventnum(2, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(m).Text, 2)
        End If
    Loop Until km >= 6
    For i = 7 To 11
        pageeventnum(2, i, 1) = Formsetting.personcom(i).Text
        pageeventnum(2, i, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).Text, 2)
    Next
End If
End Sub
Sub �ƥ�d�B�z_��l_�ϥΪ̤�()
Dim ck As Boolean
Dim m As Integer, i As Integer, j As Integer, tmpfailed As Integer

If Formsetting.comboeventcarrdus.Text = "�L" Then '=====(�L)
    For i = 1 To 18
       Randomize
       m = Int(Rnd() * 3) + 1
       Select Case m
          Case 1
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "�C1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
          Case 2
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "�j1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
          Case 3
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "��1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
       End Select
    Next
ElseIf Formsetting.comboeventcarrdus.Text = "�ۭq" Then '=====�ۭq
   If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgreus.Value = 0 Then
        For i = 1 To 18
'            If Formsetting.personus(i).Text = "(�L)" Then
            If �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�C1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�j1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "��1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
         For i = 1 To 18
            If Formsetting.personus(i).Text = "(�L)" Or i >= 7 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�C1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�j1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "��1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    End If
ElseIf Formsetting.comboeventcarrdus.Text = "�̤j��" Then '===============��̤ܳj��
    If Formsetting.persontgreus.Value = 1 Then  '===��u�W�h
         For i = 1 To 18
             Select Case Formsetting.persontgus(i).Caption
                 Case 0
                      Randomize
                      m = Int(Rnd() * 8) + 1
                      Select Case m
                          Case 1
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C3/�j1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j3/�C1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 3
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 4
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 5
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 6
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 7
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 8
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�S2" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                     End Select
                 Case 1
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C5/�j3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C5/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C8" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 2
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j5/�C3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j5/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j8" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 3
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��5/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��7" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "HP�^�_3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 4
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��3/�S3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��5" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 5
                    For j = 0 To Formsetting.personus(i).ListCount - 1
                        If Formsetting.personus(i).List(j) = "���|5" Then
                            Formsetting.personus(i).ListIndex = j
                        End If
                     Next
                 Case 6
                    For j = 0 To Formsetting.personus(i).ListCount - 1
                        If Formsetting.personus(i).List(j) = "�A�G�N5" Then
                            Formsetting.personus(i).ListIndex = j
                        End If
                     Next
                 Case 7
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�S3/��3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�S5" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
             End Select
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�C1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�j1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "��1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else  '================================����u�W�h
        For i = 1 To 18
            Do
               Randomize
               m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
               '==============================
                    Select Case Formsetting.personus(i).List(m)
                        Case "�C8"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�j8"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��7"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "HP�^�_3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "���|5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�A�G�N5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�S5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�C5/�j3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�j5/�C3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��5/��1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�j5/��1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�C5/��1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��3/�S3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�S3/��3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                    End Select
            Loop
        Next
    End If
ElseIf Formsetting.comboeventcarrdus.Text = "�H��" Or Formsetting.comboeventcarrdus.Text = "�H��(���t�t��)" Then '=====�H��
    If Formsetting.persontgreus.Value = 1 Then '===��u�W�h
        For i = 1 To 18
             tmpfailed = 0
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).List(m), 1) = Formsetting.persontgus(i).Caption Or _
                   (tmpfailed > 10 And �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).List(m), 1) = 0) Then
                    If Formsetting.comboeventcarrdus.Text = "�H��(���t�t��)" And Formsetting.personus(i).List(m) = "�t��" Then
                    Else
                        Formsetting.personus(i).ListIndex = m
                        Exit Do
                    End If
                End If
                tmpfailed = tmpfailed + 1
             Loop
         Next
        If �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
            For i = 7 To 18
                   Randomize
                   m = Int(Rnd() * 3) + 1
                   Select Case m
                      Case 1
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "�C1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                      Case 2
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "�j1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                      Case 3
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "��1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                   End Select
            Next
        End If
    Else '=============================����u�W�h
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
            If Formsetting.comboeventcarrdus.Text = "�H��(���t�t��)" And Formsetting.personus(i).List(m) = "�t��" Then
                i = i - 1
            Else
                Formsetting.personus(i).ListIndex = m
            End If
         Next
    End If
End If
End Sub
Sub �ƥ�d�B�z_��l_�q����()
Dim m As Integer, i As Integer, j As Integer, tmpfailed As Integer
Dim ay() As String

If Formsetting.comboeventcarrdcom.Text = "�L" Then '=====(�L)
    For i = 1 To 18
       Randomize
       m = Int(Rnd() * 3) + 1
       Select Case m
          Case 1
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "�C1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
          Case 2
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "�j1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
          Case 3
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "��1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
       End Select
    Next
ElseIf Formsetting.comboeventcarrdcom.Text = "�ۭq" Then '=====�ۭq
   If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgrecom.Value = 0 Then
        For i = 1 To 18
'            If Formsetting.personcom(i).Text = "(�L)" Then
            If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
         For i = 1 To 18
            If Formsetting.personcom(i).Text = "(�L)" Or i >= 7 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    End If
ElseIf Formsetting.comboeventcarrdcom.Text = "�̤j��" Then '=====��̤ܳj��
    If Formsetting.persontgrecom.Value = 1 Then  '===��u�W�h
         For i = 1 To 18
             Select Case Formsetting.persontgcom(i).Caption
                 Case 0
                      Randomize
                      m = Int(Rnd() * 8) + 1
                      Select Case m
                          Case 1
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C3/�j1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j3/�C1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 3
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 4
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 5
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 6
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 7
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 8
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�S2" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                     End Select
                 Case 1
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C5/�j3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C5/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C8" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 2
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j5/�C3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j5/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j8" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 3
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��5/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��7" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "HP�^�_3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 4
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��3/�S3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��5" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 5
                    For j = 0 To Formsetting.personcom(i).ListCount - 1
                        If Formsetting.personcom(i).List(j) = "���|5" Then
                            Formsetting.personcom(i).ListIndex = j
                        End If
                     Next
                 Case 6
                    For j = 0 To Formsetting.personcom(i).ListCount - 1
                        If Formsetting.personcom(i).List(j) = "�A�G�N5" Then
                            Formsetting.personcom(i).ListIndex = j
                        End If
                     Next
                 Case 7
                        For j = 0 To Formsetting.personcom(i).ListCount - 1
                           If Formsetting.personcom(i).List(j) = "�S3/��3" Then
                               Formsetting.personcom(i).ListIndex = j
                           End If
                        Next
             End Select
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else  '================================����u�W�h
        For i = 1 To 18
            Do
               Randomize
               m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
               '==============================
                    Select Case Formsetting.personcom(i).List(m)
                        Case "�C8"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�j8"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��7"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "HP�^�_3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "���|5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�A�G�N5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�C5/�j3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�j5/�C3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��5/��1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�j5/��1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�C5/��1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��3/�S3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�S3/��3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                    End Select
            Loop
        Next
    End If
ElseIf Formsetting.comboeventcarrdcom.Text = "�H��" Or Formsetting.comboeventcarrdcom.Text = "�H��(���t�t��)" Then '=====�H��
    If Formsetting.persontgrecom.Value = 1 Then '===��u�W�h
        For i = 1 To 18
             tmpfailed = 0
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                   (tmpfailed > 10 And �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = 0) Then
                    If Formsetting.comboeventcarrdcom.Text = "�H��(���t�t��)" And Formsetting.personcom(i).List(m) = "�t��" Then
                    Else
                        Formsetting.personcom(i).ListIndex = m
                        Exit Do
                    End If
                End If
                tmpfailed = tmpfailed + 1
             Loop
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else '=============================����u�W�h
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
            If Formsetting.comboeventcarrdcom.Text = "�H��(���t�t��)" And Formsetting.personcom(i).List(m) = "�t��" Then
                i = i - 1
            Else
                Formsetting.personcom(i).ListIndex = m
            End If
         Next
    End If
End If
End Sub
Sub �ƥ�d�B�z_����_�ϥΪ̤�()
Dim tn As Integer
Dim ay() As String
tn = BattleTurn
If tn <= 18 Then
    If tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgreus.Value = 0 Then
        If pageeventnum(1, tn, 1) <> "" Then
            ay = Split(�@��t����.�ƥ�d��Ʈw(pageeventnum(1, tn, 1), 3), "=")
            pagecardnum(���εP����d�����j������(2) + tn, 1) = ay(0)
            pagecardnum(���εP����d�����j������(2) + tn, 2) = ay(1)
            pagecardnum(���εP����d�����j������(2) + tn, 3) = ay(2)
            pagecardnum(���εP����d�����j������(2) + tn, 4) = ay(3)
            pagecardnum(���εP����d�����j������(2) + tn, 5) = 1
            pagecardnum(���εP����d�����j������(2) + tn, 6) = 1
            pagecardnum(���εP����d�����j������(2) + tn, 8) = pageeventnum(1, tn, 2)
            pagecardnum(���εP����d�����j������(2) + tn, 11) = 0
            FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
            FormMainMode.card(���εP����d�����j������(2) + tn).cardImage = app_path & "card\" & pageeventnum(1, tn, 2) & ".png"
            FormMainMode.card(���εP����d�����j������(2) + tn).CardRotationType = 1
            pageonin(���εP����d�����j������(2) + tn) = 1
            �԰��t����.�y�Эp��_�ϥΪ̤�P
            �P���ʼȮ��ܼ�(3) = ���εP����d�����j������(2) + tn
            �԰��t����.�P���ǼW�[_��P_�ϥΪ� ���εP����d�����j������(2) + tn
            pagecardnum(���εP����d�����j������(2) + tn, 9) = �P���ʼȮ��ܼ�(1) '���w�ثeLeft(�y��)
            pagecardnum(���εP����d�����j������(2) + tn, 10) = �P���ʼȮ��ܼ�(2) '���w�ثeTop(�y��)
            FormMainMode.card(���εP����d�����j������(2) + tn).Left = �P���ʼȮ��ܼ�(1)
            FormMainMode.card(���εP����d�����j������(2) + tn).Top = �P���ʼȮ��ܼ�(2)
            FormMainMode.card(���εP����d�����j������(2) + tn).ZOrder
            FormMainMode.card(���εP����d�����j������(2) + tn).Visible = True
        End If
    End If
End If
End Sub
Sub �ƥ�d�B�z_����_�q����()
Dim tn As Integer, i As Integer
Dim ay() As String
tn = BattleTurn
If tn <= 18 Then
    If tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgrecom.Value = 0 Then
        If pageeventnum(2, tn, 1) <> "" Then
            ay = Split(�@��t����.�ƥ�d��Ʈw(pageeventnum(2, tn, 1), 3), "=")
            pagecardnum(���εP����d�����j������(3) + tn, 1) = ay(0)
            pagecardnum(���εP����d�����j������(3) + tn, 2) = ay(1)
            pagecardnum(���εP����d�����j������(3) + tn, 3) = ay(2)
            pagecardnum(���εP����d�����j������(3) + tn, 4) = ay(3)
            pagecardnum(���εP����d�����j������(3) + tn, 5) = 2
            pagecardnum(���εP����d�����j������(3) + tn, 6) = 1
            pagecardnum(���εP����d�����j������(3) + tn, 8) = pageeventnum(2, tn, 2)
            pagecardnum(���εP����d�����j������(3) + tn, 11) = 0
            FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
            FormMainMode.card(���εP����d�����j������(3) + tn).cardImage = app_path & "card\" & pageeventnum(2, tn, 2) & ".png"
            pageonin(���εP����d�����j������(3) + tn) = 1
            �԰��t����.�y�Эp��_�q����P
            �P���ʼȮ��ܼ�(3) = ���εP����d�����j������(3) + tn
            �԰��t����.���εP�ܭI��
            �԰��t����.�P���ǼW�[_��P_�q�� ���εP����d�����j������(3) + tn
            pagecardnum(���εP����d�����j������(3) + tn, 9) = �P���ʼȮ��ܼ�(1) '���w�ثeLeft(�y��)
            pagecardnum(���εP����d�����j������(3) + tn, 10) = �P���ʼȮ��ܼ�(2) '���w�ثeTop(�y��)
            FormMainMode.card(���εP����d�����j������(3) + tn).Left = �P���ʼȮ��ܼ�(1)
            FormMainMode.card(���εP����d�����j������(3) + tn).Top = �P���ʼȮ��ܼ�(2)
            FormMainMode.card(���εP����d�����j������(3) + tn).ZOrder
            FormMainMode.card(���εP����d�����j������(3) + tn).Visible = True
            For i = 1 To 3
                FormMainMode.PEAFpersoncardcom(i).ZOrder
            Next
        End If
    End If
End If
End Sub
Sub �ƥ�d�B�z_�p��i��()
If ����H����ԤH��(1, 1) > 1 Or ����H����ԤH��(2, 1) > 1 Then
    �ƥ�d�O���Ȯɼ�(0, 1) = 18
Else
    �ƥ�d�O���Ȯɼ�(0, 1) = 12
End If
End Sub
Sub ����ʧ@_���m���q�����ɧޯ�Ұ�()
'===========================���涥�q���J�I(ATK-14/DEF-34)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 14, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 34, 2
'============================
'===========================���涥�q���J�I(ATK-15/DEF-35)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 15, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 35, 2
'============================
'===========================���涥�q���J�I(ATK-16/DEF-36)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 16, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 36, 2
'============================
HP�ˬd���q�� = 3
�԰��t����.����HP�ˬd
End Sub
Sub ����ʧ@_�������q�����ɧޯ�Ұ�()
'===========================���涥�q���J�I(ATK-14/DEF-34)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 14, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 34, 2
'============================
'===========================���涥�q���J�I(ATK-15/DEF-35)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 15, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 35, 2
'============================
'===========================���涥�q���J�I(ATK-16/DEF-36)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 16, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 36, 2
'============================
HP�ˬd���q�� = 3
�԰��t����.����HP�ˬd
End Sub
Sub �ޯ໡�����J_�H���d���I��_�ϥΪ�(ByVal n As Integer)
Dim i As Integer
For i = 1 To 4
    '==============================�D�ʧ�
    FormMainMode.cardus(n).CardBack_�D�ʧ�_�ޯ�W�� i, VBEPerson(1, n, 3, i, 1)
    FormMainMode.cardus(n).CardBack_�D�ʧ�_���q�N�X i, Val(VBEPerson(1, n, 3, i, 8))
    FormMainMode.cardus(n).CardBack_�D�ʧ�_�Z���N�X i, VBEPerson(1, n, 3, i, 9)
    FormMainMode.cardus(n).CardBack_�D�ʧ�_�d���N�X i, VBEPerson(1, n, 3, i, 10)
    FormMainMode.cardus(n).CardBack_�D�ʧ�_�ޯ໡�� i, VBEPerson(1, n, 3, i, 5)
    '==============================�Q�ʧ�
    FormMainMode.cardus(n).CardBack_�Q�ʧ�_�ޯ�W�� i, VBEPerson(1, n, 3, i + 4, 1)
    FormMainMode.cardus(n).CardBack_�Q�ʧ�_�ޯ໡�� i, VBEPerson(1, n, 3, i + 4, 2)
Next
End Sub
Sub �ޯ໡�����J_�H���d���I��_�q��(ByVal n As Integer)
Dim i As Integer
For i = 1 To 4
    '==============================�D�ʧ�
    FormMainMode.cardcom(n).CardBack_�D�ʧ�_�ޯ�W�� i, VBEPerson(2, n, 3, i, 1)
    FormMainMode.cardcom(n).CardBack_�D�ʧ�_���q�N�X i, Val(VBEPerson(2, n, 3, i, 8))
    FormMainMode.cardcom(n).CardBack_�D�ʧ�_�Z���N�X i, VBEPerson(2, n, 3, i, 9)
    FormMainMode.cardcom(n).CardBack_�D�ʧ�_�d���N�X i, VBEPerson(2, n, 3, i, 10)
    FormMainMode.cardcom(n).CardBack_�D�ʧ�_�ޯ໡�� i, VBEPerson(2, n, 3, i, 5)
    '==============================�Q�ʧ�
    FormMainMode.cardcom(n).CardBack_�Q�ʧ�_�ޯ�W�� i, VBEPerson(2, n, 3, i + 4, 1)
    FormMainMode.cardcom(n).CardBack_�Q�ʧ�_�ޯ໡�� i, VBEPerson(2, n, 3, i + 4, 2)
Next
End Sub

Sub �ޯ໡�����J_�H���d���I��_�洫����()
Dim n As Integer, i As Integer
For n = 1 To 2
    For i = 1 To 4
        '==============================�D�ʧ�
        Formchangeperson.card(n).CardBack_�D�ʧ�_�ޯ�W�� i, VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 1)
        Formchangeperson.card(n).CardBack_�D�ʧ�_���q�N�X i, Val(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 8))
        Formchangeperson.card(n).CardBack_�D�ʧ�_�Z���N�X i, VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9)
        Formchangeperson.card(n).CardBack_�D�ʧ�_�d���N�X i, VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 10)
        Formchangeperson.card(n).CardBack_�D�ʧ�_�ޯ໡�� i, VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 5)
        '==============================�Q�ʧ�
        Formchangeperson.card(n).CardBack_�Q�ʧ�_�ޯ�W�� i, VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i + 4, 1)
        Formchangeperson.card(n).CardBack_�Q�ʧ�_�ޯ໡�� i, VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i + 4, 2)
    Next
Next
End Sub
Sub getpage(ByVal k As Integer, m As Integer)
Dim qwp As Integer, n As Integer, uspce As String, uspme As String, yne As Boolean
If Val(���εP�U�P����������(0, 1)) < Val(���εP�U�P����������(0, 2)) Then
    yne = False
    Do
            Randomize
            qwp = Int(Rnd() * 29) + 1
            Select Case qwp
                    Case 1  '==��1�j1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\021.png"
                            pagecardnum(m, 8) = "021"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 2  '==��1�j2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\019.png"
                            pagecardnum(m, 8) = "019"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 3  '==��1�j3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\017.png"
                            pagecardnum(m, 8) = "017"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 4  '==��1��1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\025.png"
                            pagecardnum(m, 8) = "025"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 5  '==��1��2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\024.png"
                            pagecardnum(m, 8) = "024"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 6  '==��1��3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\023.png"
                            pagecardnum(m, 8) = "023"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 7  '==��2�S3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\026.png"
                            pagecardnum(m, 8) = "026"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 8  '==��3��3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a3a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\027.png"
                            pagecardnum(m, 8) = "027"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 9  '==�C6�C6��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b6b
                            pagecardnum(m, 3) = a1a
                            pagecardnum(m, 4) = b6b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\001.png"
                            pagecardnum(m, 8) = "001"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 10  '==�C1�j1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\011.png"
                            pagecardnum(m, 8) = "011"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 11  '==�C2�j1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\007.png"
                            pagecardnum(m, 8) = "007"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 12  '==�C2�j2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\006.png"
                            pagecardnum(m, 8) = "006"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 13  '==�C3�j3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\004.png"
                            pagecardnum(m, 8) = "004"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 14  '==�C5�j5��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b5b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\028.png"
                            pagecardnum(m, 8) = "028"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 15  '==�C1��1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\012.png"
                            pagecardnum(m, 8) = "012"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 16  '==�C2��1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\009.png"
                            pagecardnum(m, 8) = "009"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 17  '==�C2��2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\008.png"
                            pagecardnum(m, 8) = "008"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 18  '==�C3��3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\005.png"
                            pagecardnum(m, 8) = "005"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 19  '==�C1�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\013.png"
                            pagecardnum(m, 8) = "013"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 20  '==�C2�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\010.png"
                            pagecardnum(m, 8) = "010"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 21  '==�C4�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b4b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\003.png"
                            pagecardnum(m, 8) = "003"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 22  '==�C5�S2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\002.png"
                            pagecardnum(m, 8) = "002"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 23  '==�j4�j4��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b4b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b4b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\015.png"
                            pagecardnum(m, 8) = "015"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 24  '==�j2�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\020.png"
                            pagecardnum(m, 8) = "020"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 25  '==�j3�S2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\018.png"
                            pagecardnum(m, 8) = "018"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 26  '==�j4�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b4b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\016.png"
                            pagecardnum(m, 8) = "016"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 27  '==�j5�S2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\014.png"
                            pagecardnum(m, 8) = "014"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 28  '==��5��5��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a2a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b5b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\022.png"
                            pagecardnum(m, 8) = "022"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 29  '==��3�S5��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a2a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b5b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).cardImage = app_path & "card\029.png"
                            pagecardnum(m, 8) = "029"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
             End Select
     Loop Until yne = True
     '==================================�H����P
     Randomize
     n = Int(Rnd() * 2) + 1
     If n = 2 Then
        uspce = pagecardnum(m, 1)
        uspme = pagecardnum(m, 2)
        pagecardnum(m, 1) = pagecardnum(m, 3)
        pagecardnum(m, 2) = pagecardnum(m, 4)
        pagecardnum(m, 3) = uspce
        pagecardnum(m, 4) = uspme
        If pageonin(m) = 1 Then
           pageonin(m) = 2
'           FormMainMode.card(m).Picture = LoadPicture(app_path & "card\" & pagecardnum(m, 8) & "-" & pageonin(m) & ".bmp")
        Else
           pageonin(m) = 1
'           FormMainMode.card(m).Picture = LoadPicture(app_path & "card\" & pagecardnum(m, 8) & "-" & pageonin(m) & ".bmp")
        End If
     End If
     FormMainMode.card(m).CardRotationType = pageonin(m)
     '==============================================
     Select Case k
            Case 1 '�ϥΪ�
                pagecardnum(m, 11) = 0
                BattleCardNum = BattleCardNum - 1
                �԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
                FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
                �԰��t����.�y�Эp��_�ϥΪ̤�P
                �P���ʼȮ��ܼ�(3) = m
                pagecardnum(m, 9) = 240 '���w�ثeLeft(�y��)
                pagecardnum(m, 10) = 960 '���w�ثeTop(�y��)
                FormMainMode.card(m).Left = 240
                FormMainMode.card(m).Top = 960
                �԰��t����.�p��P���ʶZ�����
                �԰��t����.���εP�^�_���� (�P���ʼȮ��ܼ�(3))
                FormMainMode.card(m).CardEventType = False
                FormMainMode.card(m).Visible = True
                FormMainMode.card(m).ZOrder
                �԰��t����.�P���ǼW�[_��P_�ϥΪ� m
                FormMainMode.�P����.Enabled = True
                �@��t����.���ļ��� 1
            Case 2 '�q��
                pagecardnum(m, 11) = 0
                BattleCardNum = BattleCardNum - 1
                �԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
                FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
                �԰��t����.�y�Эp��_�q����P
                �P���ʼȮ��ܼ�(3) = m
                pagecardnum(m, 9) = 240 '���w�ثeLeft(�y��)
                pagecardnum(m, 10) = 960 '���w�ثeTop(�y��)
                FormMainMode.card(m).Left = 240
                FormMainMode.card(m).Top = 960
                �԰��t����.�p��P���ʶZ�����
                �԰��t����.���εP�ܭI��
                FormMainMode.card(m).CardEventType = False
                FormMainMode.card(m).Visible = True
                FormMainMode.card(m).ZOrder
                �԰��t����.�P���ǼW�[_��P_�q�� m
                FormMainMode.�P����.Enabled = True
                �@��t����.���ļ��� 1
        End Select
End If
End Sub
Sub ���εP�a�ϵP�����t�m(ByVal name As String)
Select Case name
     Case "�ܤB���|�櫰��"
           ���εP�U�P����������(0, 2) = 57
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "���b�˪L"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 0
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 0
           ���εP�U�P����������(20, 2) = 0
           ���εP�U�P����������(21, 2) = 1
           ���εP�U�P����������(22, 2) = 1
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "�U������"
           ���εP�U�P����������(0, 2) = 55
           ���εP�U�P����������(1, 2) = 2
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "�B�ʴ�`(�s)"
           ���εP�U�P����������(0, 2) = 53
           ���εP�U�P����������(1, 2) = 4
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 2
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�H��Ӧa"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 4
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 0
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "���Y����"
           ���εP�U�P����������(0, 2) = 54
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 0
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 0
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "���ɯ"
           ���εP�U�P����������(0, 2) = 52
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 2
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 0
           ���εP�U�P����������(20, 2) = 0
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 0
           ���εP�U�P����������(29, 2) = 1
    Case "ÿ�e�઺���"
           ���εP�U�P����������(0, 2) = 49
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 1
           ���εP�U�P����������(3, 2) = 1
           ���εP�U�P����������(4, 2) = 3
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 1
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 1
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�]��ù�e�����J"
           ���εP�U�P����������(0, 2) = 42
           ���εP�U�P����������(1, 2) = 0
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 2
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 0
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 0
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�ƨg�s��"
           ���εP�U�P����������(0, 2) = 47
           ���εP�U�P����������(1, 2) = 2
           ���εP�U�P����������(2, 2) = 0
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 2
           ���εP�U�P����������(5, 2) = 0
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�]�k�s��"
           ���εP�U�P����������(0, 2) = 52
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 3
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 3
           ���εP�U�P����������(11, 2) = 1
           ���εP�U�P����������(12, 2) = 1
           ���εP�U�P����������(13, 2) = 0
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�Q�i�����´�"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 1
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 2
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 0
           ���εP�U�P����������(21, 2) = 1
           ���εP�U�P����������(22, 2) = 1
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 0
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 1
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "���]�������۰}"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 0
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 1
           ���εP�U�P����������(22, 2) = 1
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 1
           ���εP�U�P����������(28, 2) = 0
           ���εP�U�P����������(29, 2) = 0
    Case Else
           ���εP�U�P����������(0, 2) = 57
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
End Select
End Sub
Sub ���εP���ϥ��ˬd()
Dim i As Integer
For i = Val(���εP�U�P����������(0, 2)) + 1 To 70
     pagecardnum(i, 6) = 5
Next
End Sub
Sub �ˮ`����_�ߧY���`_�ϥΪ�(ByVal num As Integer)
'===============================
Vss_EventBloodActionOffNum = 0
VBEStageNum(0) = 46
VBEStageNum(1) = -1 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = num '����ˮ`�H���s��
VBEStageNum(3) = 3 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = liveus(����ݾ��H��������(1, num))  '����ˮ`���ƭ�(�{��HP)
Vss_EventBloodActionChangeNum(0) = 0
Vss_EventBloodActionChangeNum(1) = 1 '����ˮ`��(1.�ϥΪ�/2.�q��)
Vss_EventBloodActionChangeNum(2) = num '����ˮ`�H���s��
Vss_EventBloodActionChangeNum(3) = 3 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
Vss_EventBloodActionChangeNum(4) = liveus(����ݾ��H��������(1, num))   '����ˮ`���ƭ�
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 46, 1
'============================
If Vss_EventBloodActionOffNum = 0 And Vss_EventBloodActionChangeNum(0) = 0 Then
    Select Case num
       Case 1
            �԰��t����.�s���T�� "�z����F" & liveus(����H����ԤH��(1, 2)) & "�I�ˮ`�C"
            FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = 0
            FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).CurrentHP = 0
            liveus(����H����ԤH��(1, 2)) = 0
            FormMainMode.bloodnumus1.Caption = 0
            FormMainMode.bloodlineout1.Width = 0
            �P�`���q��(1) = �P�`���q��(1) + 1
            �԰��t����.����ˮ`����
       Case Is > 1
            liveus(����ݾ��H��������(1, num)) = 0
            FormMainMode.cardus(����ݾ��H��������(1, num)).CardMain_����HP = 0
            FormMainMode.PEAFpersoncardus(����ݾ��H��������(1, num)).CurrentHP = 0
            �P�`���q��(1) = �P�`���q��(1) + 1
    End Select
End If
End Sub
Sub �ˮ`����_�ߧY���`_�q��(ByVal num As Integer)
'===============================
Vss_EventBloodActionOffNum = 0
VBEStageNum(0) = 46
VBEStageNum(1) = -2 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = num '����ˮ`�H���s��
VBEStageNum(3) = 3 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = livecom(����ݾ��H��������(2, num)) '����ˮ`���ƭ�(�{��HP)
Vss_EventBloodActionChangeNum(0) = 0
Vss_EventBloodActionChangeNum(1) = 2 '����ˮ`��(1.�ϥΪ�/2.�q��)
Vss_EventBloodActionChangeNum(2) = num '����ˮ`�H���s��
Vss_EventBloodActionChangeNum(3) = 3 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
Vss_EventBloodActionChangeNum(4) = livecom(����ݾ��H��������(2, num))  '����ˮ`���ƭ�
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 46, 1
'============================
If Vss_EventBloodActionOffNum = 0 And Vss_EventBloodActionChangeNum(0) = 0 Then
    Select Case num
        Case 1
            �԰��t����.�s���T�� "������F" & livecom(����H����ԤH��(2, 2)) & "�I�ˮ`�C"
            FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).CurrentHP = 0
            FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = 0
            FormMainMode.bloodnumcom1.Caption = 0
            livecom(����H����ԤH��(2, 2)) = 0
            FormMainMode.bloodlineout2.Left = 11580
            �P�`���q��(2) = �P�`���q��(2) + 1
            �԰��t����.����ˮ`����
        Case Is > 1
            FormMainMode.cardcom(����ݾ��H��������(2, num)).CardMain_����HP = 0
            livecom(����ݾ��H��������(2, num)) = 0
            FormMainMode.PEAFpersoncardcom(����ݾ��H��������(2, num)).CurrentHP = 0
            �P�`���q��(2) = �P�`���q��(2) + 1
    End Select
End If
End Sub
Sub ����_��_�ϥΪ�(ByVal num As Integer)
If liveus(����ݾ��H��������(1, num)) > 0 Then Exit Sub
'===============================
Vss_EventPersonResurrectActionOffNum = 0
'===========================���涥�q���J�I(49)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 49, 1
'============================
If Vss_EventPersonResurrectActionOffNum = 0 Then
    Select Case num
       Case 1
            FormMainMode.cardus(����H����ԤH��(1, 2)).CardMain_����HP = 1
            FormMainMode.PEAFpersoncardus(����H����ԤH��(1, 2)).CurrentHP = 1
            liveus(����H����ԤH��(1, 2)) = 1
            FormMainMode.bloodlineout1.Width = �Z�����(1, 1, 1)
            FormMainMode.bloodnumus1.Caption = liveus(����H����ԤH��(1, 2))
       Case Is > 1
            liveus(����ݾ��H��������(1, num)) = 1
            FormMainMode.PEAFpersoncardus(����ݾ��H��������(1, num)).CurrentHP = 1
            FormMainMode.cardus(����ݾ��H��������(1, num)).CardMain_����HP = 1
    End Select
End If
End Sub
Sub ����_��_�q��(ByVal num As Integer)
'===============================
If livecom(����ݾ��H��������(2, num)) > 0 Then Exit Sub
'===============================
Vss_EventPersonResurrectActionOffNum = 0
'===========================���涥�q���J�I(49)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 49, 1
'============================
If Vss_EventPersonResurrectActionOffNum = 0 Then
    Select Case num
        Case 1
            FormMainMode.PEAFpersoncardcom(����H����ԤH��(2, 2)).CurrentHP = 1
            FormMainMode.cardcom(����H����ԤH��(2, 2)).CardMain_����HP = 1
            FormMainMode.bloodnumcom1.Caption = 1
            livecom(����H����ԤH��(2, 2)) = 1
            FormMainMode.bloodlineout2.Left = 11580 - �Z�����(1, 2, 1)
        Case Is > 1
            FormMainMode.cardcom(����ݾ��H��������(2, num)).CardMain_����HP = 1
            livecom(����ݾ��H��������(2, num)) = 1
            FormMainMode.PEAFpersoncardcom(����ݾ��H��������(2, num)).CurrentHP = 1
    End Select
End If
End Sub
Sub �ѪR��q�ܤ�(ByVal str As String, ByVal uscom As Integer)
Dim cmdstr() As String
Dim i As Integer

cmdstr = Split(str, "=")
If ��ܦC����ƭ���w������(uscom) = False Then
    For i = 0 To UBound(cmdstr) - 1
        Select Case Mid(cmdstr(i), 1, 1)
            Case "+"
                �������m��l�`��(uscom) = �������m��l�`��(uscom) + Mid(cmdstr(i), 2, Len(cmdstr(i)))
            Case "-"
                �������m��l�`��(uscom) = �������m��l�`��(uscom) - Mid(cmdstr(i), 2, Len(cmdstr(i)))
            Case "*"
                �������m��l�`��(uscom) = �������m��l�`��(uscom) * Mid(cmdstr(i), 2, Len(cmdstr(i)))
            Case "/"
                �������m��l�`��(uscom) = Int(�������m��l�`��(uscom) / Mid(cmdstr(i), 2, Len(cmdstr(i))) + 0.9)
            Case "\"
                �������m��l�`��(uscom) = �������m��l�`��(uscom) \ Mid(cmdstr(i), 2, Len(cmdstr(i)))
            Case "@"
                �������m��l�`��(uscom) = Mid(cmdstr(i), 2, Len(cmdstr(i)))
                ��ܦC����ƭ���w������(uscom) = True
                Exit Sub '==���w�ƭȮɨ�L�ܤƶq�L��
        End Select
    Next
End If
End Sub
Sub �C����Ե����������()
Dim i As Integer
'==========
For i = 1 To FormMainMode.PEAFvssc.UBound
   Unload FormMainMode.PEAFvssc(i)
Next
'==========
'==========
For i = 1 To FormMainMode.card.UBound
    Unload FormMainMode.card(i)
Next
'==========
For i = 1 To FormMainMode.cardus.UBound
    Unload FormMainMode.cardus(i)
Next
For i = 1 To FormMainMode.cardcom.UBound
    Unload FormMainMode.cardcom(i)
Next
'==========
End Sub
Sub �C������P����ŧi�{��()
Dim i As Integer

���εP����d�����j������(1) = ���εP�U�P����������(0, 2) + 18 + 18
���εP����d�����j������(2) = ���εP�U�P����������(0, 2)
���εP����d�����j������(3) = ���εP�U�P����������(0, 2) + 18
���εP����d�����j������(4) = ���εP�U�P����������(0, 2) + 18 + 18
���εP����d�����j������(5) = -1
For i = 1 To ���εP����d�����j������(1)
    Load FormMainMode.card(i)
    Set FormMainMode.card(i).Container = FormMainMode.PEAttackingForm
    FormMainMode.card(i).Left = 240
    FormMainMode.card(i).Top = 960
    FormMainMode.card(i).Visible = False
    FormMainMode.card(i).CardEventType = False
    FormMainMode.card(i).LocationType = 0
Next
End Sub
Sub �s���T��(ByVal messagestr As String)
FormMainMode.PEAFInterface.Message messagestr
End Sub
Sub �C������d������Х�()
Dim i As Integer

For i = 1 To 3
    Load FormMainMode.cardus(i)
    Load FormMainMode.cardcom(i)
Next
End Sub
Sub ����ʧ@_�t���`�d�P�i�Ƨ�s()
FormMainMode.PEAFInterface.Cardnum = BattleCardNum
FormMainMode.pageul.Caption = BattleCardNum
End Sub
Sub ����ʧ@_�q����U���q�X�P��������(ByVal turnnum As Integer)
Dim ckl As Integer

Select Case turnnum
    Case 1
        FormMainMode.�������q_���q2.Enabled = True
    Case 2
        FormMainMode.PEAFInterface.BnOKStartListen
        '==============
        �p�H���Y�����ʤ�V��(1) = 1
        �p�H���Y�����ʤ�V��(2) = 2
        FormMainMode.�p�H���Y������_�ϥΪ�.Enabled = True
        FormMainMode.�p�H���Y������_�q��.Enabled = True
        '==============
        ���q���A�� = 1
        �@��t����.���ļ��� 6
        �԰��t����.�ɶ��b_���]
        FormMainMode.trtimeline.Enabled = True
    Case 3
        turnpageonin = 1
        ���q���A�� = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        If Vss_EventPlayerAllActionOffNum(1) = 1 Then
            For ckl = 1 To ���εP����d�����j������(1)
                FormMainMode.card(ckl).CardEnabledType = False
            Next
            FormMainMode.PEAFInterface.BnOKEnabled False
            ���ݮɶ���C(2).Add 47
            FormMainMode.���ݮɶ�_2.Enabled = True
        ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
            For ckl = 1 To ���εP����d�����j������(1)
                FormMainMode.card(ckl).CardEnabledType = False
            Next
            FormMainMode.PEAFInterface.BnOKEnabled False
            ���ݮɶ���C(2).Add 45
            FormMainMode.���ݮɶ�_2.Enabled = True
        End If
End Select
End Sub
Sub ���ʶ��q���ʫe���涥�q�I�s(ByVal ns As Integer)
Dim moveusTempnum As Integer, movecomTempnum As Integer, moveusSelectnum As Integer, movecomSelectnum As Integer
If Vss_PersonMoveControlNum(1, 2) = 0 Then
    moveusTempnum = moveus + Vss_PersonMoveControlNum(1, 1)
Else
    moveusTempnum = Vss_PersonMoveControlNum(1, 1)
End If
If Vss_PersonMoveControlNum(2, 2) = 0 Then
    movecomTempnum = movecom + Vss_PersonMoveControlNum(2, 1)
Else
    movecomTempnum = Vss_PersonMoveControlNum(2, 1)
End If
'==================================
If moveusTempnum < 0 Then moveusTempnum = 0
If movecomTempnum < 0 Then movecomTempnum = 0
'==================================
If Vss_PersonMoveActionChangeNum(1, 1) = 1 Then
    moveusSelectnum = Vss_PersonMoveActionChangeNum(1, 2)
Else
    moveusSelectnum = FormMainMode.��ܦC1.���ʶ��q��ܭ�
End If
If Vss_PersonMoveActionChangeNum(2, 1) = 1 Then
    movecomSelectnum = Vss_PersonMoveActionChangeNum(2, 2)
Else
    movecomSelectnum = �q���貾�ʶ��q��ܼ�
    If movecomTempnum <= 0 Then
       movecomSelectnum = 2
    End If
End If
'===============
If Vss_EventPlayerAllActionOffNum(1) = 1 Then moveusSelectnum = 0
If Vss_EventPlayerAllActionOffNum(2) = 1 Then movecomSelectnum = 0
ReDim VBEStageNum(0 To 4) As Integer
VBEStageNum(0) = ns
VBEStageNum(1) = moveusTempnum '�ϥΪ̤��`���ʼ�
VBEStageNum(2) = movecomTempnum '�q�����`���ʼ�
VBEStageNum(3) = moveusSelectnum '�ϥΪ̤�ثe���ʶ��q��ʿ��
VBEStageNum(4) = movecomSelectnum '�q����ثe���ʶ��q��ʿ��
'===========================���涥�q���J�I(ns)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, ns, 1
'============================
End Sub
Sub ��������p�d��T�g�J(ByVal uscom As Integer, ByVal num As Integer)
'Dim tmpobj As New clsPersonCard

Select Case uscom
 Case 1
    With FormMainMode.PEAFpersoncardus(num)
        .Level = uslevel(num)
        .ATK = atkus(num)
        .DEF = defus(num)
        .CurrentHP = liveus(num)
        .AllHP = liveusmax(num)
        .PersonName = nameus(num)
    End With
 Case 2
    With FormMainMode.PEAFpersoncardcom(num)
        .Level = comlevel(num)
        .ATK = atkcom(num)
        .DEF = defcom(num)
        .CurrentHP = livecom(num)
        .AllHP = livecommax(num)
        .PersonName = namecom(num)
    End With
End Select
End Sub
