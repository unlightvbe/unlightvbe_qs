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
Public �P�`���q��(1 To 3) As Integer '�P�֦��`���q��(1.�ϥΪ�/2.�q��/3.�`�p)
Public �P���ʼȮ��ܼ�(1 To 3) As Long '�P���ʭp�ƾ��Ȯ��ܼ�(1.Left���/2.Top���/3.�P�i�s��)
Public �ثe��(1 To 33) As Integer '�`�Ȯ��ܼ�
'2.(1)�ܦ��ϥΪ̵o�P���q-(2)�ܦ��q���o�P���q-(3)�ܦ��o�P�ˬd���q
'3.�ϥΪ̥X�P����Z���έp,
'4.�ϥΪ̤�P�a������Z���έp,
'5.�ϥΪ̵P����z���,
'6.�q���X�P/�G�P�p�ƼȮɼ�,
'7.�q���X�P����Z���έp,
'8.�q����P����Z���έp,
'9.�q���P����z���,
'10.���P���q��,
'11.���P�έp�Ȯɭ�,
'12.���P�P��ȼȮɼ�,
'13.�ϥΪ̤�P���2�C��1�i�P�s�����ʼȮɼ�,
'14.���ݮɶ��p�ƾ�(1,2)�Ȯɼ�,
'15�P���ʭp�ƾ����涥�q��-(1)�o�P���q-(2)���P���q-�q����P���,
'16.�P½�P/�^�P/��P�P�s���Ȯɼ�,
'17.�q����P������q��-(1)-�q���X�P���q-(2)���P�������q
'20.�ϥΪ�½�P�P�s���Ȯɼ�
'21.�ϥΪ̤�P������q��
'22.���ݮɶ��p�ƾ����涥�q��
'25.�q�����ʶ��q�X�P�p�ƼȮ��ܼ�
'26.��l�����ޯ�Ұʶ��qHP�ˬd�O�_������
'29.�ޯ�Ұʰʵe�p�ƼȮɼ�
'30.�P��~�P�ɬJ���d�P�ƶq�Ȯɬ�����
'31.���ʶ��q��lTimer-�O�_�Ĥ@���ҰʼȮɬ�����
'32.�ϥΪ̥X�P-AI�X�P����P�ثe��
'33.�ϥΪ̥X�P-AI�X�P���ʶ��q��ܦ�ʼ�
Public �Z�����_���P�Ȯɼ�() As Integer  '���P�ӧO�Z�����Ȯ��x�s�ܼ�(��x����,1.Left���/2.Top���/3.�P�i�s��)
Public ���q���A�� As Integer '�C���q�}�l�������A�ˬd��(1.�}�l���q(�ϥΪ�)/2.�������q(�ϥΪ�)/3.�}�l���q(�q��)/4.�������q(�q��)/5.�洫����)
Public �p�H���Y�����ʤ�V��(1 To 2) As Integer '�p�H���Y�����ʤ�V���A��(1.�ϥΪ�/2.�q��[1.�V��,2.�V�~])
Public ��q�p�ƾ��ʵe�Ȯ��ܼ�(1 To 2, 1 To 2) As Integer '�}�l��l���q-��q�ʵe�p�ƾ��Ȯ��ܼ�(1.�ϥΪ̦��/2.�q�����,1.�C�����ʶq/2.�O�_�w����)
Public �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1 To 4, 1 To 3) As Integer '�ɶ��b�i���C���ܤƶ��q�����Ȯ��ܼ�(1~3(1)����ܤƶq(1(1).�ɶ��b(���~))/2.�ثe�֭p�q/3.�ثe�C��(R,G,B),4.(1)�ɶ��b(�~)���q��-(1)���ܬ�-(2)���ܶ�/2.�ثe�֭p�q/3.�ثe�C��(R))
Public �}�l�d�����ʰʵe������(1 To 2, 1 To 4) As Integer   '�}�l�ɨC�i�d�����ʰʵe����������(1.�ϥΪ�/2.�q��,1~3.�d��/4.�ثe�ĴX�i)
Public �洫��������Ȯ��ܼ�(1 To 4) As Integer '�洫������������Ȯɼ�(1.�ϥΪ�/2.�q��/3.�O�_��U����/4.�洫���⧹���涥�q��)
Public �԰��Ҧ��ӱѬ����� As Integer '�԰��t�η�e�ӱѬ����Ȯ��ܼ�(1.�ϥΪ̤�ӧQ/2.�ϥΪ̤�ѥ_/3.����)
Public �q���貾�ʶ��q��ܼ� As Integer '���ʶ��q�q�����ܤ���ʼȮ��ܼ�
Public �q����ƥ�d�O�_�X����ܼ� As Boolean '�q������X�ƥ�d�O�_�X���Ȯɬ���
Public �H���d���I���s��������(1 To 7) As Integer '�H���d���I���ޯ໡���H���s���Ȯ��ܼ�(1.(1).�ϥΪ�/(2).�q��,2.��n��,3.�ثe�ϥΪ̤�ϥΤH���s��/4.�ثe��ܤ��ޯ�s��(�ϥΪ̤�ϥΤH��)/5.�ثe��ܤ��ޯ�s��(��L)/6~7.�ثe��ܤ��ޯ�s��(�洫����)
Public �Y���淾�q�Ȯ��ܼ�(1 To 10) As Integer '�Y�뤶�����q�Ȯ��ܼ�(1.�@�^�X������P�_(1.�e/2.��),2.�Y��ᦳ�Ķˮ`��,3.�Y���ˮ`��H(1.�ϥΪ�/2.�q��),4.(1.�ϥΪ̥���/2.�q������)/5.��e���(�ϥΪ�)/6.��e���(�q��)/7.�t�Τ��λ��(�ϥΪ�)/8.�t�Τ��λ��(�q��)/9.�Y��e���-�`��(�ϥΪ�)/10.�Y��e���-�`��(�q��))
Public �H�������ˬd�Ȯ��ܼ�(1 To 3) As Integer '�H�������ˬd�p�ƾ������Ȯ��ܼ�(1.�ثe�p��/2.�ϥΪ̼аO/3.�q���аO)
Public ���εP�U�P����������(0 To 29, 1 To 2) As Integer '�U�������εP�P���������Ȯ��ܼ�(0.(1)�ثe�w�o�P�`�ƶq/(2)�ثe�����P�`�ƶq,1~29.(1)�ثe�w�ϥΤ��P��/(2)�ӵP����ϥΤ��`�ƶq)
Public �d���H����T�ɮ�Ū�����Ѭ����� As String '�d���H����T�ɮ�Ū�����Ѯ��ɮצW�����Ȯ��ܼ�
Public ��ܦC����ƭ���w������(1 To 2) As Boolean '�԰��t����ܦC����ƭ���w��ܬ����ܼ�(1.�ϥΪ̤�/2.�q����)
Public �O�_�t�Τ��� As Boolean '�O�_���t�Τ��������
Public �԰��Y�뤶���H����ø�ϸ��|������(1 To 2) As String '�԰��t���Y�뤶������H����ø�ϸ��|������(1.�ϥΪ̤�/2.�q����)
Public �H����ڪ��A��Ʈw(1 To 2, 1 To 3, 1 To 9) As String '�H����ڪ��A���
Public �t����ܬɭ������� As Integer '�԰��t����ܤ����]�w������(1.�ª�/2.�s��)
Public ���ݮɶ���C(1 To 2) As New Collection '�԰��t�ε��ݮɶ��p�ƾ��u�@��C
Public �H�����`���A�C��(1 To 2, 1 To 3) As Collection '���`���A�C��(1.�ϥΪ�/2.�q��,��n��)
Public ActiveSkillObj(1 To 2, 1 To 4) As clsPersonActiveSkill '�԰��t�ΥD�ʧޯ໡������(1.�ϥΪ̤�/2.�q����,��n��)
Public PersonCardShowOnMode(1 To 2, 1 To 3) As Boolean '�԰��t�ΤH���d����T�O�_�i��(1.�ϥΪ̤�/2.�q����,��n��)
Public CardDeckCollection(0 To 9) As Collection '�԰��t�Υd�P�P�ﶰ�X(0.�d�P����/1.�P��/2.�Ӧa�P/3.�ƥ�d�P��(�ϥΪ̤�)/4.�ƥ�d�P��(�q����)/5.��P(�ϥΪ̤�)/6.�X�P(�ϥΪ̤�)/7.��P(�q����)/8.�X�P(�q����)/9.��P)
Public ActionCardTotNum As Integer '�԰��t�Υd�P�`�o�������
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
    Dim stageInfoListObj As clsVSStageObj
    Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
    '===============================
    VBEStageNum(0) = 46
    VBEStageNum(1) = -1 '����ˮ`��(1.�ϥΪ�/2.�q��)
    VBEStageNum(2) = num '����ˮ`�H���s��
    VBEStageNum(3) = 2 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    VBEStageNum(4) = tot '����ˮ`���ƭ�
    stageInfoListObj.Argument = tot  '����ˮ`���ƭ�
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`��(1.�ϥΪ�/2.�q��)
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num) '����ˮ`�H���s��
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2" '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    '===========================���涥�q���J�I(46)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 46, 1
    '============================
    If stageInfoListObj.CommandStr = "PersonBloodControl" Then
        If stageInfoListObj.Value = "BLOODOFF" Then
            Exit Sub
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                tot = Val(tmpstr(1))
            End If
        End If
    End If
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
Sub �^�_����_�ϥΪ�(ByVal tot As Integer, ByVal num As Integer, ByVal statusfrom As Integer, ByVal isEvent As Boolean, ByVal isSysCall As Boolean)
If isEvent = True Then
    Dim stageInfoListObj As clsVSStageObj
    Dim tmpflagoff As Boolean
    If isSysCall = True Then
        Set stageInfoListObj = New clsVSStageObj
        stageInfoListObj.StageNum = 0
        stageInfoListObj.CommandStr = "@System"
        stageInfoListObj.Value = "0"
        ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
    Else
        Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
    End If
    '===============================
    If statusfrom = 0 Then
        ReDim VBEStageNum(0 To 5) As Integer
        VBEStageNum(4) = 0 'Ĳ�o�ƥ��
        VBEStageNum(5) = 0 'Ĳ�o�ƥ���t
    End If
    VBEStageNum(0) = 48
    VBEStageNum(1) = -1 '�^�_��(1.�ϥΪ�/2.�q��)
    VBEStageNum(2) = num '�^�_�H���s��
    VBEStageNum(3) = tot '�^�_���ƭ�
    stageInfoListObj.Argument = tot '�^�_���ƭ�
    '===========================���涥�q���J�I(48)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 48, 1
    '============================
    tmpflagoff = False
    If stageInfoListObj.CommandStr = "PersonBloodControl" Or (stageInfoListObj.CommandStr = "@System" And isSysCall = True) Then
        If stageInfoListObj.Value = "HPLOFF" Then
            tmpflagoff = True
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "HPLCHANGE" Then
                tot = Val(tmpstr(1))
            End If
        End If
    End If
    If isSysCall = True Then
        ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
    End If
    '===============================
    If tmpflagoff = True Then Exit Sub
    '===============================
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
Sub �^�_����_�q��(ByVal tot As Integer, ByVal num As Integer, ByVal statusfrom As Integer, ByVal isEvent As Boolean, ByVal isSysCall As Boolean)
If isEvent = True Then
    Dim stageInfoListObj As clsVSStageObj
    Dim tmpflagoff As Boolean
    If isSysCall = True Then
        Set stageInfoListObj = New clsVSStageObj
        stageInfoListObj.StageNum = 0
        stageInfoListObj.CommandStr = "@System"
        stageInfoListObj.Value = "0"
        ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
    Else
        Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
    End If
    '===============================
    If statusfrom = 0 Then
        ReDim VBEStageNum(0 To 5) As Integer
        VBEStageNum(4) = 0 'Ĳ�o�ƥ��
        VBEStageNum(5) = 0 'Ĳ�o�ƥ���t
    End If
    VBEStageNum(0) = 48
    VBEStageNum(1) = -2 '�^�_��(�t�ΥN��)
    VBEStageNum(2) = num '�^�_�H���s��
    VBEStageNum(3) = tot '�^�_���ƭ�
    stageInfoListObj.Argument = tot '�^�_���ƭ�
    '===========================���涥�q���J�I(48)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 48, 1
    '============================
    tmpflagoff = False
    If stageInfoListObj.CommandStr = "PersonBloodControl" Or (stageInfoListObj.CommandStr = "@System" And isSysCall = True) Then
        If stageInfoListObj.Value = "HPLOFF" Then
            tmpflagoff = True
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "HPLCHANGE" Then
                tot = Val(tmpstr(1))
            End If
        End If
    End If
    If isSysCall = True Then
        ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
    End If
    '===============================
    If tmpflagoff = True Then Exit Sub
    '===============================
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
Dim stageInfoListObj As New clsVSStageObj
Dim tmpflagoff As Boolean
stageInfoListObj.StageNum = 0
stageInfoListObj.CommandStr = "@System"
stageInfoListObj.Value = "0"
���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
'===============================
ReDim VBEStageNum(0 To 6) As Integer
VBEStageNum(0) = 46
VBEStageNum(1) = -1 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = 1 '����ˮ`�H���s��
VBEStageNum(3) = 1 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = tot '����ˮ`���ƭ�
VBEStageNum(5) = 0 '�Ӧۨt�Ϊ��ˮ`
VBEStageNum(6) = 0 '�Ӧۨt�Ϊ��ˮ`
stageInfoListObj.Argument = tot  '����ˮ`���ƭ�
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`��(1.�ϥΪ�/2.�q��)
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`�H���s��
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 46, 1
'============================
tmpflagoff = False
If stageInfoListObj.CommandStr = "@System" Then
    If stageInfoListObj.Value = "BLOODOFF" Then
        tmpflagoff = True
    Else
        Dim tmpstr() As String
        tmpstr = Split(stageInfoListObj.Value, "%")
        If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
            tot = Val(tmpstr(1))
        End If
    End If
End If
���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
'===============================
If tmpflagoff = True Then Exit Sub
'===============================
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
End Sub
Sub �ˮ`����_�ޯઽ��_�q��(ByVal tot As Integer, ByVal num As Integer, ByVal isEvent As Boolean)
If tot <= 0 Then Exit Sub
If isEvent = True Then
    Dim stageInfoListObj As clsVSStageObj
    Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
    '===============================
    VBEStageNum(0) = 46
    VBEStageNum(1) = -2 '����ˮ`��(1.�ϥΪ�/2.�q��)
    VBEStageNum(2) = num '����ˮ`�H���s��
    VBEStageNum(3) = 2 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    VBEStageNum(4) = tot '����ˮ`���ƭ�
    stageInfoListObj.Argument = tot  '����ˮ`���ƭ�
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2" '����ˮ`��(1.�ϥΪ�/2.�q��)
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num) '����ˮ`�H���s��
    stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2" '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
    '===========================���涥�q���J�I(46)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 46, 1
    '============================
    If stageInfoListObj.CommandStr = "PersonBloodControl" Then
        If stageInfoListObj.Value = "BLOODOFF" Then
            Exit Sub
        Else
            Dim tmpstr() As String
            tmpstr = Split(stageInfoListObj.Value, "%")
            If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
                tot = Val(tmpstr(1))
            End If
        End If
    End If
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
Dim stageInfoListObj As New clsVSStageObj
Dim tmpflagoff As Boolean
stageInfoListObj.StageNum = 0
stageInfoListObj.CommandStr = "@System"
stageInfoListObj.Value = "0"
���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
'===============================
ReDim VBEStageNum(0 To 6) As Integer
VBEStageNum(0) = 46
VBEStageNum(1) = -2 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = 1 '����ˮ`�H���s��
VBEStageNum(3) = 1 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = tot '����ˮ`���ƭ�
VBEStageNum(5) = 0 '�Ӧۨt�Ϊ��ˮ`
VBEStageNum(6) = 0 '�Ӧۨt�Ϊ��ˮ`
stageInfoListObj.Argument = tot  '����ˮ`���ƭ�
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`��(1.�ϥΪ�/2.�q��)
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`�H���s��
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 46, 1
'============================
tmpflagoff = False
If stageInfoListObj.CommandStr = "@System" Then
    If stageInfoListObj.Value = "BLOODOFF" Then
        tmpflagoff = True
    Else
        Dim tmpstr() As String
        tmpstr = Split(stageInfoListObj.Value, "%")
        If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
            tot = Val(tmpstr(1))
        End If
    End If
End If
���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
'===============================
If tmpflagoff = True Then Exit Sub
'============================
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
End Sub
Sub ����ʧ@_�ϥΪ�_��P(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)))(CStr(n))
    
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) - 1
    �ثe��(5) = Utils.IndexOf(�԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Location = 3
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �ثe��(15) = 4
    Select Case tmpcard.CardType
        Case 1 '���εP
            �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 2
        Case 2 '�ƥ�d
            �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 9
    End Select
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�P��_�^�P_�ϥΪ�(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)))(CStr(n))
    
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    tmpcard.Owner = 1
    tmpcard.Location = 1
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.���εP�^�_���� n
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�P���ǼW�[_��P_�ϥΪ�
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 5
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�q���P_���P_�ϥΪ�(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)))(CStr(n))
    
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    �ثe��(9) = Utils.IndexOf(�԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Owner = 1
    tmpcard.Location = 1
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�P���ǼW�[_��P_�ϥΪ�
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 5
    �ثe��(15) = 2
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�ϥΪ̵P_���P_�q��(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)))(CStr(n))
    
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pageusglead = Val(FormMainMode.pageusglead) - 1
    �ثe��(5) = Utils.IndexOf(�԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Owner = 2
    tmpcard.Location = 1
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�P���ǼW�[_��P_�q��
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 7
    �ثe��(15) = 20
    �԰��t����.���εP�ܭI��
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�P��_�^�P_�q��(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)))(CStr(n))
    
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    tmpcard.Owner = 2
    tmpcard.Location = 1
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�P���ǼW�[_��P_�q��
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 7
    �԰��t����.���εP�ܭI��
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_½�P(ByVal n As Integer)
    FormMainMode.card(n).Width = 810
    FormMainMode.card(n).Height = 1260
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
Sub �P���ǼW�[_�X�P_�q��()
pagecomleadmax(1) = pagecomleadmax(1) + 1
End Sub
Sub �P���ǼW�[_��P_�q��()
pagecomleadmax(0) = pagecomleadmax(0) + 1
End Sub
Sub �P���ǼW�[_��P_�ϥΪ�()
pageusleadmax(0) = pageusleadmax(0) + 1
End Sub
Sub �P���ǼW�[_�X�P_�ϥΪ�()
pageusleadmax(1) = pageusleadmax(1) + 1
End Sub
Sub ����ʧ@_�q��_��P(ByVal n As Integer)
    Dim tmpcard As clsActionCard
    Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)))(CStr(n))
    
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) - 1
    �ثe��(9) = Utils.IndexOf(�԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n))), tmpcard)
    tmpcard.Location = 3
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = n
    tmpcard.XYLeft = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    Select Case tmpcard.CardType
        Case 1 '���εP
            �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 2
        Case 2 '�ƥ�d
            �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(n)), 9
    End Select
    �ثe��(15) = 5
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
End Sub
Sub ����ʧ@_�~�P()
�԰��t����.����ʧ@_�~�P_�Ӧa�P�^�P
�԰��t����.����ʧ@_�~�P_�P��~�P
End Sub
Sub ����ʧ@_�~�P_�Ӧa�P�^�P()
Dim tmpcard As clsActionCard

�ثe��(30) = �԰��t����.CardDeckCollection(1).Count

Do While �԰��t����.CardDeckCollection(2).Count > 0
    Set tmpcard = �԰��t����.CardDeckCollection(2)(1)
    
    tmpcard.Owner = 0
    tmpcard.Location = 4
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 2, 1
Loop

BattleCardNum = �԰��t����.CardDeckCollection(1).Count
�԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
End Sub
Sub ����ʧ@_�~�P_�P��~�P()
Dim tmpnewCollection As New Collection
Dim tmpcard As clsActionCard
Dim i As Integer

For i = 1 To �ثe��(30) '�N�J���d�P�O�d��P��̤W�h
    Set tmpcard = �԰��t����.CardDeckCollection(1)(1)
    
    tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
    �԰��t����.CardDeckCollection(1).Remove 1
Next

Do While �԰��t����.CardDeckCollection(1).Count > 0
    Randomize
    i = Int(Rnd() * �԰��t����.CardDeckCollection(1).Count) + 1
    Set tmpcard = �԰��t����.CardDeckCollection(1)(i)
    
    tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
    �԰��t����.CardDeckCollection(1).Remove i
Loop

Set �԰��t����.CardDeckCollection(1) = Nothing
Set �԰��t����.CardDeckCollection(1) = tmpnewCollection

End Sub
Sub ����ʧ@_�~�P_�ƥ�d�P��~�P()
Dim tmpnewCollection As Collection
Dim tmpcard As clsActionCard
Dim i As Integer, m As Integer

For m = 3 To 4
    Set tmpnewCollection = New Collection
    Do While �԰��t����.CardDeckCollection(m).Count > 0
        Randomize
        i = Int(Rnd() * �԰��t����.CardDeckCollection(m).Count) + 1
        Set tmpcard = �԰��t����.CardDeckCollection(m)(i)
        
        tmpnewCollection.Add tmpcard, CStr(tmpcard.CardNum)
        �԰��t����.CardDeckCollection(m).Remove i
    Loop
    
    Set �԰��t����.CardDeckCollection(m) = Nothing
    Set �԰��t����.CardDeckCollection(m) = tmpnewCollection
Next

End Sub
Sub ����ʧ@_��P_���εP(ByVal uscom As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(1)(1)
Dim n As Integer
'==================================�H����P
Randomize
n = Int(Rnd() * 2) + 1
If n = 2 Then
   Call tmpcard.Reverse
End If
FormMainMode.card(tmpcard.CardNum).CardRotationType = tmpcard.CardOnIn
'==============================================
Select Case uscom
    Case 1 '�ϥΪ�
        tmpcard.ComMark = 0
        tmpcard.Owner = 1
        tmpcard.Location = 1
        �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 1, 5
        '=================
        BattleCardNum = BattleCardNum - 1
        �԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
        FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
        �԰��t����.�y�Эp��_�ϥΪ̤�P
        �P���ʼȮ��ܼ�(3) = tmpcard.CardNum
        tmpcard.XYLeft = 240 '���w�ثeLeft(�y��)
        tmpcard.XYTop = 960 '���w�ثeTop(�y��)
        FormMainMode.card(tmpcard.CardNum).Left = 240
        FormMainMode.card(tmpcard.CardNum).Top = 960
        �԰��t����.�p��P���ʶZ�����
        �԰��t����.���εP�^�_���� (�P���ʼȮ��ܼ�(3))
        FormMainMode.card(tmpcard.CardNum).CardEventType = False
        FormMainMode.card(tmpcard.CardNum).Visible = True
        FormMainMode.card(tmpcard.CardNum).ZOrder
        �԰��t����.�P���ǼW�[_��P_�ϥΪ�
        FormMainMode.�P����.Enabled = True
        �@��t����.���ļ��� 1
    Case 2 '�q��
        tmpcard.ComMark = 0
        tmpcard.Owner = 2
        tmpcard.Location = 1
        �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 1, 7
        '=================
        BattleCardNum = BattleCardNum - 1
        �԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
        FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
        �԰��t����.�y�Эp��_�q����P
        �P���ʼȮ��ܼ�(3) = tmpcard.CardNum
        tmpcard.XYLeft = 240 '���w�ثeLeft(�y��)
        tmpcard.XYTop = 960 '���w�ثeTop(�y��)
        FormMainMode.card(tmpcard.CardNum).Left = 240
        FormMainMode.card(tmpcard.CardNum).Top = 960
        �԰��t����.�p��P���ʶZ�����
        �԰��t����.���εP�ܭI��
        FormMainMode.card(tmpcard.CardNum).CardEventType = False
        FormMainMode.card(tmpcard.CardNum).Visible = True
        FormMainMode.card(tmpcard.CardNum).ZOrder
        �԰��t����.�P���ǼW�[_��P_�q��
        FormMainMode.�P����.Enabled = True
        �@��t����.���ļ��� 1
End Select
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
Sub ����ʧ@_�Z���ܧ�(ByVal m As Integer, ByVal isEvent As Boolean, ByVal isSysCall As Boolean)
'===========================���涥�q���J�I(47)
If isEvent = True Then
    Dim stageInfoListObj As clsVSStageObj
    Dim tmpflagoff As Boolean
    Dim tmpuscom As Integer
    If isSysCall = True Then
        ReDim VBEStageNum(0 To 3) As Integer
        VBEStageNum(3) = 0  'Ĳ�o�ƥ��
        '======================
        Set stageInfoListObj = New clsVSStageObj
        stageInfoListObj.StageNum = 0
        stageInfoListObj.CommandStr = "@System"
        stageInfoListObj.Value = "0"
        ���涥�q�t����.VBEVSStageInfoList.Add stageInfoListObj
    Else
        Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
    End If
    VBEStageNum(0) = 47
    VBEStageNum(1) = movecp '�ܧ�e�Z��
    VBEStageNum(2) = m  '�ܧ��Z��
    If isSysCall = True Then
        tmpuscom = 1
    Else
        tmpuscom = Abs(VBEStageNum(3))
    End If
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l tmpuscom, 47, 1
    '=====================
    tmpflagoff = False
    If stageInfoListObj.CommandStr = "BattleMoveControl" Or (stageInfoListObj.CommandStr = "@System" And isSysCall = True) Then
        If stageInfoListObj.Value = "BMCOFF" Then
            tmpflagoff = True
        End If
    End If
    If isSysCall = True Then
        ���涥�q�t����.VBEVSStageInfoList.Remove ���涥�q�t����.VBEVSStageInfoList.Count
    End If
    '===============================
    If tmpflagoff = True Then Exit Sub
    '===============================
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
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(�P���ʼȮ��ܼ�(3))))(CStr(�P���ʼȮ��ܼ�(3)))

If �P���ʼȮ��ܼ�(1) >= tmpcard.XYLeft Then
   �Z�����(2, 1, 1) = (�P���ʼȮ��ܼ�(1) - tmpcard.XYLeft) \ 8
Else
   �Z�����(2, 1, 1) = -((tmpcard.XYLeft - �P���ʼȮ��ܼ�(1)) \ 8)
End If

If �P���ʼȮ��ܼ�(2) >= tmpcard.XYTop Then
   �Z�����(2, 1, 2) = (�P���ʼȮ��ܼ�(2) - tmpcard.XYTop) \ 8
Else
   �Z�����(2, 1, 2) = -((tmpcard.XYTop - �P���ʼȮ��ܼ�(2)) \ 8)
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
        For ckl = 1 To �԰��t����.ActionCardTotNum
            FormMainMode.card(ckl).CardEnabledType = False
        Next
        FormMainMode.PEAFInterface.BnOKEnabled False
        ���ݮɶ���C(2).Add 47
        FormMainMode.���ݮɶ�_2.Enabled = True
    ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
        For ckl = 1 To �԰��t����.ActionCardTotNum
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
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(num)))(CStr(num))

FormMainMode.card(num).Width = 810
FormMainMode.card(num).Height = 1260
FormMainMode.card(num).LocationType = 1
FormMainMode.card(num).CardEventType = False
FormMainMode.card(num).CardRotationType = tmpcard.CardOnIn
End Sub
Sub ���P�p��Z�����_�ϥΪ�()
Dim i As Integer
Dim tmpcard As clsActionCard

If �԰��t����.CardDeckCollection(6).Count > 0 Then
    ReDim �Z�����_���P�Ȯɼ�(1 To �԰��t����.CardDeckCollection(6).Count, 1 To 3) As Integer
Else
    Erase �Z�����_���P�Ȯɼ�
End If

For i = 1 To �԰��t����.CardDeckCollection(6).Count
    Set tmpcard = �԰��t����.CardDeckCollection(6)(i)
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = tmpcard.CardNum
    tmpcard.XYLeft = FormMainMode.card(tmpcard.CardNum).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(tmpcard.CardNum).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �Z�����_���P�Ȯɼ�(i, 1) = �Z�����(2, 1, 1)
    �Z�����_���P�Ȯɼ�(i, 2) = �Z�����(2, 1, 2)
    �Z�����_���P�Ȯɼ�(i, 3) = tmpcard.CardNum
Next
End Sub
Sub ���P�p��Z�����_�q��()
Dim i As Integer
Dim tmpcard As clsActionCard

If �԰��t����.CardDeckCollection(8).Count > 0 Then
    ReDim �Z�����_���P�Ȯɼ�(1 To �԰��t����.CardDeckCollection(8).Count, 1 To 3) As Integer
Else
    Erase �Z�����_���P�Ȯɼ�
End If

For i = 1 To �԰��t����.CardDeckCollection(8).Count
    Set tmpcard = �԰��t����.CardDeckCollection(8)(i)
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = tmpcard.CardNum
    tmpcard.XYLeft = FormMainMode.card(tmpcard.CardNum).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(tmpcard.CardNum).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �Z�����_���P�Ȯɼ�(i, 1) = �Z�����(2, 1, 1)
    �Z�����_���P�Ȯɼ�(i, 2) = �Z�����(2, 1, 2)
    �Z�����_���P�Ȯɼ�(i, 3) = tmpcard.CardNum
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
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(7)(CStr(Index))

If tmpcard.Location = 1 And tmpcard.Owner = 2 Then
   tmpcard.Location = 2
   If tmpcard.UpperType = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + Val(tmpcard.UpperNum)
      If turnatk = 2 And movecp = 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(tmpcard.UpperNum)
          �������m��l�`��(4) = �������m��l�`��(4) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + Val(tmpcard.UpperNum)
      If turnatk = 2 And movecp > 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(tmpcard.UpperNum)
          �������m��l�`��(4) = �������m��l�`��(4) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + Val(tmpcard.UpperNum)
      If turnatk = 1 And �������m��l�`��(4) = 0 Then
          �������m��l�`��(4) = �������m��l�`��(4) + defcom(����H����ԤH��(2, 2))
      End If
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) + Val(tmpcard.UpperNum)
         �������m��l�`��(4) = �������m��l�`��(4) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + Val(tmpcard.UpperNum)
   End If
   If tmpcard.UpperType = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + Val(tmpcard.UpperNum)
   End If
   '===================
    �ثe��(9) = Utils.IndexOf(�԰��t����.CardDeckCollection(7), tmpcard)
    pagecomleadmax(1) = Val(pagecomleadmax(1)) + 1
    pageqlead(2) = Val(pageqlead(2)) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) + 1
    tmpcard.ComMark = 2
   '===================�H�U�O�X�P���
    �ثe��(7) = 0
    FormMainMode.�q���X�P_�X�P���_�a��.Enabled = True
   '=============�H�U�O�P����(�X�P)(�q��)
    �԰��t����.�y�Эp��_�q���X�P
    �P���ʼȮ��ܼ�(3) = Index
    tmpcard.XYLeft = FormMainMode.card(Index).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 7, 8
    �ثe��(15) = 0
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
   '================�H�U�O��P���
   �ثe��(8) = 0
   �ثe��(17) = 1
   FormMainMode.�q���X�P_��P���.Enabled = True
   '===================�H�U�O�ƥ�d�ˬd�αҰ�
   If tmpcard.UpperType = a6a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.���|_�q�� Index, tmpcard.UpperNum
   End If
   If turnatk = 1 Or turnatk = 2 Then
        If tmpcard.UpperType = a7a Then
            �ƥ�d�O���Ȯɼ�(2, 3) = 1
            �ƥ�d.�A�G�N_�q�� Index, tmpcard.UpperNum
        End If
   End If
   If tmpcard.UpperType = a8a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.HP�^�__�q�� Index, tmpcard.UpperNum
   End If
   If tmpcard.UpperType = a9a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.�t��_�q�� Index, tmpcard.UpperNum
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
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(8)(CStr(Index))

If tmpcard.Location = 2 And tmpcard.Owner = 2 Then
   tmpcard.Location = 1
   If tmpcard.UpperType = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - Val(tmpcard.UpperNum)
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(tmpcard.UpperNum)
          �������m��l�`��(4) = �������m��l�`��(4) - Val(tmpcard.UpperNum)
      End If
      If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
          �������m��l�`��(4) = 0
      End If
   End If
   If tmpcard.UpperType = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - Val(tmpcard.UpperNum)
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(tmpcard.UpperNum)
          �������m��l�`��(4) = �������m��l�`��(4) - Val(tmpcard.UpperNum)
      End If
      If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
          �������m��l�`��(4) = 0
      End If
   End If
   If tmpcard.UpperType = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - Val(tmpcard.UpperNum)
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) - Val(tmpcard.UpperNum)
         �������m��l�`��(4) = �������m��l�`��(4) - Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - Val(tmpcard.UpperNum)
   End If
   If tmpcard.UpperType = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - Val(tmpcard.UpperNum)
   End If
   '================
    �ثe��(9) = Utils.IndexOf(�԰��t����.CardDeckCollection(8), tmpcard)
    pagecomleadmax(0) = Val(pagecomleadmax(0)) + 1
    pageqlead(2) = Val(pageqlead(2)) - 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
    tmpcard.ComMark = 0
   '=============�H�U�O�P����(�^�P)(�q��)
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = Index
    tmpcard.XYLeft = FormMainMode.card(Index).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = FormMainMode.card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.���εP�ܭI��
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 8, 7
    �ثe��(15) = 0
    FormMainMode.�P����.Enabled = True
    �@��t����.���ļ��� 1
   '================�H�U�O�X�P���
   �ثe��(7) = 0
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
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(8)(CStr(Index))

Call tmpcard.Reverse
�@��t����.���ļ��� 3

FormMainMode.card(Index).CardRotationType = tmpcard.CardOnIn

If tmpcard.UpperType = a1a Then
   atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + tmpcard.UpperNum
   If turnatk = 2 And movecp = 1 And �������m��l�`��(4) = 0 Then
       �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
   End If
   If turnatk = 2 And movecp = 1 Then
       �������m��l�`��(2) = �������m��l�`��(2) + Val(tmpcard.UpperNum)
       �������m��l�`��(4) = �������m��l�`��(4) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a5a Then
   atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + tmpcard.UpperNum
   If turnatk = 2 And movecp > 1 And �������m��l�`��(4) = 0 Then
       �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
   End If
   If turnatk = 2 And movecp > 1 Then
       �������m��l�`��(2) = �������m��l�`��(2) + Val(tmpcard.UpperNum)
       �������m��l�`��(4) = �������m��l�`��(4) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a2a Then
   atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + tmpcard.UpperNum
   If turnatk = 1 And �������m��l�`��(4) = 0 Then
       �������m��l�`��(4) = �������m��l�`��(4) + defcom(����H����ԤH��(2, 2))
   End If
   If turnatk = 1 Then
      �������m��l�`��(2) = �������m��l�`��(2) + Val(tmpcard.UpperNum)
      �������m��l�`��(4) = �������m��l�`��(4) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a3a Then
   atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + tmpcard.UpperNum
End If
If tmpcard.UpperType = a4a Then
   atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + tmpcard.UpperNum
End If
'======================================
If tmpcard.LowerType = a1a Then
   atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - tmpcard.LowerNum
   If turnatk = 2 And movecp = 1 Then
       �������m��l�`��(2) = �������m��l�`��(2) - Val(tmpcard.LowerNum)
       �������m��l�`��(4) = �������m��l�`��(4) - Val(tmpcard.LowerNum)
   End If
   If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
       �������m��l�`��(4) = 0
   End If
End If
If tmpcard.LowerType = a5a Then
   atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - tmpcard.LowerNum
   If turnatk = 2 And movecp > 1 Then
       �������m��l�`��(2) = �������m��l�`��(2) - Val(tmpcard.LowerNum)
       �������m��l�`��(4) = �������m��l�`��(4) - Val(tmpcard.LowerNum)
   End If
   If �������m��l�`��(4) = atkcom(����H����ԤH��(2, 2)) Then
       �������m��l�`��(4) = 0
   End If
End If
If tmpcard.LowerType = a2a Then
   atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - tmpcard.LowerNum
   If turnatk = 1 Then
       �������m��l�`��(2) = �������m��l�`��(2) - Val(tmpcard.LowerNum)
       �������m��l�`��(4) = �������m��l�`��(4) - Val(tmpcard.LowerNum)
   End If
End If
If tmpcard.LowerType = a3a Then
   atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - tmpcard.LowerNum
End If
If tmpcard.LowerType = a4a Then
   atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - tmpcard.LowerNum
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
Dim tmpcard As clsActionCard

For a = 1 To �԰��t����.CardDeckCollection(7).Count
    Set tmpcard = �԰��t����.CardDeckCollection(7)(a)
    If tmpcard.ComMark <> 1 Then
        If tmpcard.UpperType = a1a Then
            tmpcard.ComMark = 1
        ElseIf tmpcard.LowerType = a1a Then
            Call tmpcard.Reverse
            tmpcard.ComMark = 1
        End If
    End If
Next
End Sub
Sub comatk2()
Dim j As Integer
Dim tmpcard As clsActionCard

For j = 1 To �԰��t����.CardDeckCollection(7).Count
    Set tmpcard = �԰��t����.CardDeckCollection(7)(j)
    If tmpcard.ComMark <> 1 Then
        If tmpcard.UpperType = a5a Then
            tmpcard.ComMark = 1
        ElseIf tmpcard.LowerType = a5a Then
            Call tmpcard.Reverse
            tmpcard.ComMark = 1
        End If
    End If
Next
End Sub
Sub comatk_���z��AI�޾ɵ{��_�W�X�P�i��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal choose As Integer)
Dim werstr As String, werbo As Boolean
Dim a As Integer, k As Integer
Dim tmpcard As clsActionCard

If movecpre = 1 And turn = 1 Then
   werstr = a1a
ElseIf movecpre > 1 And turn = 1 Then
   werstr = a5a
ElseIf turn = 2 Then
   werstr = a2a
End If
'=================================
For a = 1 To �԰��t����.CardDeckCollection(7).Count
    werbo = False
    Set tmpcard = �԰��t����.CardDeckCollection(7)(a)
    For k = 1 To UBound(cardAInumOvertenrecord)
        If tmpcard.CardNum = cardAInumOvertenrecord(k) Then
            werbo = True
        End If
    Next
    If tmpcard.ComMark <> 1 And werbo = False Then
        If tmpcard.UpperType = werstr Then
            tmpcard.ComMark = 1
        ElseIf tmpcard.LowerType = werstr Then
            Call tmpcard.Reverse
            tmpcard.ComMark = 1
        End If
        If choose = 1 And tmpcard.ComMark = 0 Then
            tmpcard.ComMark = 1
        End If
    End If
Next
End Sub
Sub moveatkin()
Dim j As Integer
Dim tmpcard As clsActionCard

Do
    For j = 1 To �԰��t����.CardDeckCollection(7).Count
        Set tmpcard = �԰��t����.CardDeckCollection(7)(j)
        If tmpcard.CardType = 2 And tmpcard.ComMark <> 1 Then
            If tmpcard.UpperType = a3a And tmpcard.LowerType = a3a Then '���ʳ歱�ƥ�d�u��
                 tmpcard.ComMark = 1
                 �ثe��(25) = �ثe��(25) + tmpcard.UpperNum
            End If
            If �ثe��(25) >= 2 Then Exit Do
        End If
    Next
    For j = 1 To �԰��t����.CardDeckCollection(7).Count
        Set tmpcard = �԰��t����.CardDeckCollection(7)(j)
        If tmpcard.ComMark <> 1 Then
            If tmpcard.UpperType = a3a Then
                tmpcard.ComMark = 1
                �ثe��(25) = �ثe��(25) + 1
            ElseIf tmpcard.LowerType = a3a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                �ثe��(25) = �ثe��(25) + tmpcard.UpperNum
            End If
            If �ثe��(25) >= 2 Then Exit Do
        End If
    Next
    Exit Do
Loop
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
����ʧ@_�Z���ܧ� movecp, False, True
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
����ʧ@_�Z���ܧ� movecp, False, True
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
Sub �ƥ�d�B�z_��l_�ϥΪ̤�()
Dim ck As Boolean
Dim m As Integer, i As Integer, j As Integer, tmpfailed As Integer, tmpcardstr As String

If Formsetting.comboeventcarrdus.Text = "�L" Then '=====(�L)
    For i = 1 To 18
        Randomize
        m = Int(Rnd() * 3) + 1
        Select Case m
            Case 1
                tmpcardstr = "�C1"
            Case 2
                tmpcardstr = "�j1"
            Case 3
                tmpcardstr = "��1"
        End Select
        �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
    Next
ElseIf Formsetting.comboeventcarrdus.Text = "�ۭq" Then '=====�ۭq
   If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgreus.Value = 0 Then
        For i = 1 To 18
            If �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
            Else
                tmpcardstr = Formsetting.personus(i).Text
            End If
            �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
         Next
    ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
         For i = 1 To 18
            If i >= 7 Or �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
            Else
                tmpcardstr = Formsetting.personus(i).Text
            End If
            �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
         Next
    End If
ElseIf Formsetting.comboeventcarrdus.Text = "�̤j��" Then '===============��̤ܳj��
    If Formsetting.persontgreus.Value = 1 Then  '===��u�W�h
        For i = 1 To 18
            If i = 7 And �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then Exit For
            
            Select Case Formsetting.persontgus(i).Caption
                Case 0
                    Randomize
                    m = Int(Rnd() * 8) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�C3/�j1"
                        Case 2
                            tmpcardstr = "�j3/�C1"
                        Case 3
                            tmpcardstr = "��3/��1"
                        Case 4
                            tmpcardstr = "�C3/��1"
                        Case 5
                            tmpcardstr = "�j3/��1"
                        Case 6
                            tmpcardstr = "�C3/��1"
                        Case 7
                            tmpcardstr = "�j3/��1"
                        Case 8
                            tmpcardstr = "�S2"
                    End Select
                Case 1
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�C5/�j3"
                        Case 2
                            tmpcardstr = "�C5/��1"
                        Case 3
                            tmpcardstr = "�C8"
                    End Select
                Case 2
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�j5/�C3"
                        Case 2
                            tmpcardstr = "�j5/��1"
                        Case 3
                            tmpcardstr = "�j8"
                    End Select
                 Case 3
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "��5/��1"
                        Case 2
                            tmpcardstr = "��7"
                        Case 3
                            tmpcardstr = "HP�^�_3"
                      End Select
                Case 4
                    Randomize
                    m = Int(Rnd() * 2) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "��3/�S3"
                        Case 2
                            tmpcardstr = "��5"
                    End Select
                Case 5
                    tmpcardstr = "���|5"
                Case 6
                    tmpcardstr = "�A�G�N5"
                Case 7
                    Randomize
                    m = Int(Rnd() * 2) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�S3/��3"
                        Case 1
                            tmpcardstr = "�S5"
                    End Select
            End Select
            �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
        If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
                �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
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
                        Exit Do
                    Case "�j8"
                        Exit Do
                    Case "��7"
                        Exit Do
                    Case "��5"
                        Exit Do
                    Case "HP�^�_3"
                        Exit Do
                    Case "���|5"
                        Exit Do
                    Case "�A�G�N5"
                        Exit Do
                    Case "�S5"
                        Exit Do
                    Case "�C5/�j3"
                        Exit Do
                    Case "�j5/�C3"
                        Exit Do
                    Case "��5/��1"
                        Exit Do
                    Case "�j5/��1"
                        Exit Do
                    Case "�C5/��1"
                        Exit Do
                    Case "��3/�S3"
                        Exit Do
                    Case "�S3/��3"
                        Exit Do
                End Select
            Loop
            tmpcardstr = Formsetting.personus(i).List(m)
            �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
    End If
ElseIf Formsetting.comboeventcarrdus.Text = "�H��" Or Formsetting.comboeventcarrdus.Text = "�H��(���t�t��)" Then '=====�H��
    If Formsetting.persontgreus.Value = 1 Then '===��u�W�h
        For i = 1 To 18
            If i = 7 And �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then Exit For
            
            tmpfailed = 0
            Do
                Randomize
                m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).List(m), 1) = Formsetting.persontgus(i).Caption Or _
                   (tmpfailed > 10 And �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).List(m), 1) = 0) Then
                    If Formsetting.comboeventcarrdus.Text = "�H��(���t�t��)" And Formsetting.personus(i).List(m) = "�t��" Then
                    Else
                        tmpcardstr = Formsetting.personus(i).List(m)
                        Exit Do
                    End If
                End If
                tmpfailed = tmpfailed + 1
            Loop
            �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
        If �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
                �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
            Next
        End If
    Else '=============================����u�W�h
        For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
            If Formsetting.comboeventcarrdus.Text = "�H��(���t�t��)" And Formsetting.personus(i).List(m) = "�t��" Then
                i = i - 1
            Else
                tmpcardstr = Formsetting.personus(i).List(m)
            End If
            �԰��t����.�o��d�P_�ƥ�d 1, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
    End If
End If
End Sub
Sub �ƥ�d�B�z_��l_�q����()
Dim m As Integer, i As Integer, j As Integer, tmpfailed As Integer, tmpcardstr As String
Dim ay() As String

If Formsetting.comboeventcarrdcom.Text = "�L" Then '=====(�L)
    For i = 1 To 18
        Randomize
        m = Int(Rnd() * 3) + 1
        Select Case m
            Case 1
                tmpcardstr = "�C1"
            Case 2
                tmpcardstr = "�j1"
            Case 3
                tmpcardstr = "��1"
        End Select
        �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
    Next
ElseIf Formsetting.comboeventcarrdcom.Text = "�ۭq" Then '=====�ۭq
    If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgrecom.Value = 0 Then
        For i = 1 To 18
            If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
            Else
                tmpcardstr = Formsetting.personcom(i).Text
            End If
            �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
         Next
    ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
         For i = 1 To 18
            If i >= 7 Or �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
            Else
                tmpcardstr = Formsetting.personcom(i).Text
            End If
            �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
         Next
    End If
ElseIf Formsetting.comboeventcarrdcom.Text = "�̤j��" Then '=====��̤ܳj��
    If Formsetting.persontgrecom.Value = 1 Then  '===��u�W�h
        For i = 1 To 18
            If i = 7 And �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then Exit For
            
            Select Case Formsetting.persontgcom(i).Caption
                Case 0
                    Randomize
                    m = Int(Rnd() * 8) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�C3/�j1"
                        Case 2
                            tmpcardstr = "�j3/�C1"
                        Case 3
                            tmpcardstr = "��3/��1"
                        Case 4
                            tmpcardstr = "�C3/��1"
                        Case 5
                            tmpcardstr = "�j3/��1"
                        Case 6
                            tmpcardstr = "�C3/��1"
                        Case 7
                            tmpcardstr = "�j3/��1"
                        Case 8
                            tmpcardstr = "�S2"
                    End Select
                Case 1
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�C5/�j3"
                        Case 2
                            tmpcardstr = "�C5/��1"
                        Case 3
                            tmpcardstr = "�C8"
                    End Select
                Case 2
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�j5/�C3"
                        Case 2
                            tmpcardstr = "�j5/��1"
                        Case 3
                            tmpcardstr = "�j8"
                    End Select
                 Case 3
                    Randomize
                    m = Int(Rnd() * 3) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "��5/��1"
                        Case 2
                            tmpcardstr = "��7"
                        Case 3
                            tmpcardstr = "HP�^�_3"
                      End Select
                Case 4
                    Randomize
                    m = Int(Rnd() * 2) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "��3/�S3"
                        Case 2
                            tmpcardstr = "��5"
                    End Select
                Case 5
                    tmpcardstr = "���|5"
                Case 6
                    tmpcardstr = "�A�G�N5"
                Case 7
                    Randomize
                    m = Int(Rnd() * 2) + 1
                    Select Case m
                        Case 1
                            tmpcardstr = "�S3/��3"
                        Case 1
                            tmpcardstr = "�S5"
                    End Select
            End Select
            �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
        If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
                �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
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
                        Exit Do
                    Case "�j8"
                        Exit Do
                    Case "��7"
                        Exit Do
                    Case "��5"
                        Exit Do
                    Case "HP�^�_3"
                        Exit Do
                    Case "���|5"
                        Exit Do
                    Case "�A�G�N5"
                        Exit Do
                    Case "�S5"
                        Exit Do
                    Case "�C5/�j3"
                        Exit Do
                    Case "�j5/�C3"
                        Exit Do
                    Case "��5/��1"
                        Exit Do
                    Case "�j5/��1"
                        Exit Do
                    Case "�C5/��1"
                        Exit Do
                    Case "��3/�S3"
                        Exit Do
                    Case "�S3/��3"
                        Exit Do
                End Select
            Loop
            tmpcardstr = Formsetting.personcom(i).List(m)
            �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
    End If
ElseIf Formsetting.comboeventcarrdcom.Text = "�H��" Or Formsetting.comboeventcarrdcom.Text = "�H��(���t�t��)" Then '=====�H��
    If Formsetting.persontgrecom.Value = 1 Then '===��u�W�h
        For i = 1 To 18
            If i = 7 And �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then Exit For
            
            tmpfailed = 0
            Do
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                   (tmpfailed > 10 And �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = 0) Then
                    If Formsetting.comboeventcarrdcom.Text = "�H��(���t�t��)" And Formsetting.personcom(i).List(m) = "�t��" Then
                    Else
                        tmpcardstr = Formsetting.personcom(i).List(m)
                        Exit Do
                    End If
                End If
                tmpfailed = tmpfailed + 1
            Loop
            �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
        If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                        tmpcardstr = "�C1"
                    Case 2
                        tmpcardstr = "�j1"
                    Case 3
                        tmpcardstr = "��1"
                End Select
                �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
            Next
        End If
    Else '=============================����u�W�h
        For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
            If Formsetting.comboeventcarrdcom.Text = "�H��(���t�t��)" And Formsetting.personcom(i).List(m) = "�t��" Then
                i = i - 1
            Else
                tmpcardstr = Formsetting.personcom(i).List(m)
            End If
            �԰��t����.�o��d�P_�ƥ�d 2, tmpcardstr, �@��t����.�ƥ�d��Ʈw(tmpcardstr, 2)
        Next
    End If
End If
End Sub
Sub �ƥ�d�B�z_����_�ϥΪ̤�()
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(3)(1)

FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
�԰��t����.�y�Эp��_�ϥΪ̤�P
�P���ʼȮ��ܼ�(3) = tmpcard.CardNum
�԰��t����.�P���ǼW�[_��P_�ϥΪ�
tmpcard.Location = 1
tmpcard.Owner = 1
tmpcard.XYLeft = �P���ʼȮ��ܼ�(1) '���w�ثeLeft(�y��)
tmpcard.XYTop = �P���ʼȮ��ܼ�(2) '���w�ثeTop(�y��)
FormMainMode.card(tmpcard.CardNum).Left = �P���ʼȮ��ܼ�(1)
FormMainMode.card(tmpcard.CardNum).Top = �P���ʼȮ��ܼ�(2)
FormMainMode.card(tmpcard.CardNum).ZOrder
FormMainMode.card(tmpcard.CardNum).Visible = True

�԰��t����.�d�P�P�ﶰ�X�� tmpcard, 3, 5
End Sub
Sub �ƥ�d�B�z_����_�q����()
Dim i As Integer
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(4)(1)

FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
�԰��t����.�y�Эp��_�q����P
�P���ʼȮ��ܼ�(3) = tmpcard.CardNum
�԰��t����.���εP�ܭI��
�԰��t����.�P���ǼW�[_��P_�q��
tmpcard.Location = 1
tmpcard.Owner = 2
tmpcard.XYLeft = �P���ʼȮ��ܼ�(1) '���w�ثeLeft(�y��)
tmpcard.XYTop = �P���ʼȮ��ܼ�(2) '���w�ثeTop(�y��)
FormMainMode.card(tmpcard.CardNum).Left = �P���ʼȮ��ܼ�(1)
FormMainMode.card(tmpcard.CardNum).Top = �P���ʼȮ��ܼ�(2)
FormMainMode.card(tmpcard.CardNum).ZOrder
FormMainMode.card(tmpcard.CardNum).Visible = True
�԰��t����.�d�P�P�ﶰ�X�� tmpcard, 4, 7

For i = 1 To 3
    FormMainMode.PEAFpersoncardcom(i).ZOrder
Next
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
Sub �o��d�P_���εP()
Dim i As Integer, j As Integer, tmpnewActionCardNum As Integer
Dim tmpcard As clsActionCard, tmpindexobj As clsCollectionIndex

For i = 1 To ���εP�U�P����������(0, 2)
    tmpnewActionCardNum = �԰��t����.�C������P����Ыصo��
    
    Set tmpcard = New clsActionCard
    tmpcard.CardNum = tmpnewActionCardNum
    tmpcard.Location = 4
    tmpcard.CardType = 1
    For j = 1 To UBound(���εP�U�P����������, 1)
        If Val(���εP�U�P����������(j, 1)) < Val(���εP�U�P����������(j, 2)) Then
            Select Case j
                Case 1  '==��1�j1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\021.png"
                    tmpcard.ImageStr = "021"
                    tmpcard.CardOnIn = 1
                Case 2  '==��1�j2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\019.png"
                    tmpcard.ImageStr = "019"
                    tmpcard.CardOnIn = 1
                Case 3  '==��1�j3��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b3b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\017.png"
                    tmpcard.ImageStr = "017"
                    tmpcard.CardOnIn = 1
                Case 4  '==��1��1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\025.png"
                    tmpcard.ImageStr = "025"
                    tmpcard.CardOnIn = 1
                Case 5  '==��1��2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\024.png"
                    tmpcard.ImageStr = "024"
                    tmpcard.CardOnIn = 1
                Case 6  '==��1��3��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b3b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\023.png"
                    tmpcard.ImageStr = "023"
                    tmpcard.CardOnIn = 1
                Case 7  '==��2�S3��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b3b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\026.png"
                    tmpcard.ImageStr = "026"
                    tmpcard.CardOnIn = 1
                Case 8  '==��3��3��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a3a
                    tmpcard.UpperNum = b3b
                    tmpcard.LowerType = a3a
                    tmpcard.LowerNum = b3b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\027.png"
                    tmpcard.ImageStr = "027"
                    tmpcard.CardOnIn = 1
                Case 9  '==�C6�C6��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b6b
                    tmpcard.LowerType = a1a
                    tmpcard.LowerNum = b6b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\001.png"
                    tmpcard.ImageStr = "001"
                    tmpcard.CardOnIn = 1
                Case 10  '==�C1�j1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\011.png"
                    tmpcard.ImageStr = "011"
                    tmpcard.CardOnIn = 1
                Case 11  '==�C2�j1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\007.png"
                    tmpcard.ImageStr = "007"
                    tmpcard.CardOnIn = 1
                Case 12  '==�C2�j2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\006.png"
                    tmpcard.ImageStr = "006"
                    tmpcard.CardOnIn = 1
                Case 13  '==�C3�j3��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b3b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b3b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\004.png"
                    tmpcard.ImageStr = "004"
                    tmpcard.CardOnIn = 1
                Case 14  '==�C5�j5��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b5b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b5b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\028.png"
                    tmpcard.ImageStr = "028"
                    tmpcard.CardOnIn = 1
                Case 15  '==�C1��1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\012.png"
                    tmpcard.ImageStr = "012"
                    tmpcard.CardOnIn = 1
                Case 16  '==�C2��1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\009.png"
                    tmpcard.ImageStr = "009"
                    tmpcard.CardOnIn = 1
                Case 17  '==�C2��2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\008.png"
                    tmpcard.ImageStr = "008"
                    tmpcard.CardOnIn = 1
                Case 18  '==�C3��3��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b3b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b3b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\005.png"
                    tmpcard.ImageStr = "005"
                    tmpcard.CardOnIn = 1
                Case 19  '==�C1�S1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b1b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\013.png"
                    tmpcard.ImageStr = "013"
                    tmpcard.CardOnIn = 1
                Case 20  '==�C2�S1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\010.png"
                    tmpcard.ImageStr = "010"
                    tmpcard.CardOnIn = 1
                Case 21  '==�C4�S1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b4b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\003.png"
                    tmpcard.ImageStr = "003"
                    tmpcard.CardOnIn = 1
                Case 22  '==�C5�S2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a1a
                    tmpcard.UpperNum = b5b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\002.png"
                    tmpcard.ImageStr = "002"
                    tmpcard.CardOnIn = 1
                Case 23  '==�j4�j4��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a5a
                    tmpcard.UpperNum = b4b
                    tmpcard.LowerType = a5a
                    tmpcard.LowerNum = b4b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\015.png"
                    tmpcard.ImageStr = "015"
                    tmpcard.CardOnIn = 1
                Case 24  '==�j2�S1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a5a
                    tmpcard.UpperNum = b2b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\020.png"
                    tmpcard.ImageStr = "020"
                    tmpcard.CardOnIn = 1
                Case 25  '==�j3�S2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a5a
                    tmpcard.UpperNum = b3b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\018.png"
                    tmpcard.ImageStr = "018"
                    tmpcard.CardOnIn = 1
                Case 26  '==�j4�S1��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a5a
                    tmpcard.UpperNum = b4b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b1b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\016.png"
                    tmpcard.ImageStr = "016"
                    tmpcard.CardOnIn = 1
                Case 27  '==�j5�S2��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a5a
                    tmpcard.UpperNum = b5b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b2b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\014.png"
                    tmpcard.ImageStr = "014"
                    tmpcard.CardOnIn = 1
                Case 28  '==��5��5��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a2a
                    tmpcard.UpperNum = b5b
                    tmpcard.LowerType = a2a
                    tmpcard.LowerNum = b5b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\022.png"
                    tmpcard.ImageStr = "022"
                    tmpcard.CardOnIn = 1
                Case 29  '==��3�S5��
                    ���εP�U�P����������(j, 1) = Val(���εP�U�P����������(j, 1)) + 1
                    ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                    tmpcard.UpperType = a2a
                    tmpcard.UpperNum = b3b
                    tmpcard.LowerType = a4a
                    tmpcard.LowerNum = b5b
                    tmpcard.Owner = 0
                    FormMainMode.card(i).CardImage = app_path & "card\029.png"
                    tmpcard.ImageStr = "029"
                    tmpcard.CardOnIn = 1
            End Select
            �԰��t����.CardDeckCollection(1).Add tmpcard, CStr(tmpnewActionCardNum)
            
            Set tmpindexobj = New clsCollectionIndex
            tmpindexobj.CollectionIndex = 1
            tmpindexobj.CardNum = tmpnewActionCardNum
            tmpindexobj.Index = �԰��t����.CardDeckCollection(0).Count + 1
            �԰��t����.CardDeckCollection(0).Add tmpindexobj, CStr(tmpnewActionCardNum)
            Exit For
        End If
    Next
Next
End Sub
Sub �o��d�P_�ƥ�d(ByVal uscom As Integer, ByVal cardname As String, ByVal filename As String, Optional ByVal beforeindex As Integer)
Dim ay() As String
Dim tn As Integer, tmpnewActionCardNum As Integer
Dim tmpcard As clsActionCard
Dim tmpindexobj As clsCollectionIndex

If cardname = "" Or filename = "" Then Exit Sub

tmpnewActionCardNum = �԰��t����.�C������P����Ыصo��

Set tmpcard = New clsActionCard
tmpcard.CardNum = tmpnewActionCardNum
'============
Erase ay
ay = Split(�@��t����.�ƥ�d��Ʈw(cardname, 3), "=")
'============
tmpcard.UpperType = ay(0)
tmpcard.UpperNum = ay(1)
tmpcard.LowerType = ay(2)
tmpcard.LowerNum = ay(3)
tmpcard.Owner = 0
tmpcard.Location = 0
tmpcard.ImageStr = filename
tmpcard.ComMark = 0
FormMainMode.card(tmpnewActionCardNum).CardImage = app_path & "card\" & filename & ".png"
FormMainMode.card(tmpnewActionCardNum).CardRotationType = 1
tmpcard.CardOnIn = 1
tmpcard.CardType = 2

Set tmpindexobj = New clsCollectionIndex
tmpindexobj.CardNum = tmpnewActionCardNum
tmpindexobj.Index = �԰��t����.CardDeckCollection(0).Count + 1

Select Case uscom
    Case 1
        tmpindexobj.CollectionIndex = 3
        �԰��t����.CardDeckCollection(0).Add tmpindexobj, CStr(tmpnewActionCardNum)
        
        If beforeindex > 0 And beforeindex <= �԰��t����.CardDeckCollection(3).Count Then
            �԰��t����.CardDeckCollection(3).Add tmpcard, CStr(tmpnewActionCardNum), beforeindex
        Else
            �԰��t����.CardDeckCollection(3).Add tmpcard, CStr(tmpnewActionCardNum)
        End If
    Case 2
        tmpindexobj.CollectionIndex = 4
        �԰��t����.CardDeckCollection(0).Add tmpindexobj, CStr(tmpnewActionCardNum)
        
        If beforeindex > 0 And beforeindex <= �԰��t����.CardDeckCollection(4).Count Then
            �԰��t����.CardDeckCollection(4).Add tmpcard, CStr(tmpnewActionCardNum), beforeindex
        Else
            �԰��t����.CardDeckCollection(4).Add tmpcard, CStr(tmpnewActionCardNum)
        End If
End Select
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
Sub �ˮ`����_�ߧY���`_�ϥΪ�(ByVal num As Integer)
Dim stageInfoListObj As clsVSStageObj
Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
'===============================
VBEStageNum(0) = 46
VBEStageNum(1) = -1 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = num '����ˮ`�H���s��
VBEStageNum(3) = 3 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = liveus(����ݾ��H��������(1, num))  '����ˮ`���ƭ�(�{��HP)
stageInfoListObj.Argument = liveus(����ݾ��H��������(1, num))   '����ˮ`���ƭ�
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "1" '����ˮ`��(1.�ϥΪ�/2.�q��)
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num) '����ˮ`�H���s��
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "3" '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 46, 1
'============================
If stageInfoListObj.CommandStr = "PersonBloodControl" Then
    If stageInfoListObj.Value = "BLOODOFF" Then
        Exit Sub
    Else
        Dim tmpstr() As String
        tmpstr = Split(stageInfoListObj.Value, "%")
        If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
            Exit Sub
        End If
    End If
End If
'============================
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
End Sub
Sub �ˮ`����_�ߧY���`_�q��(ByVal num As Integer)
Dim stageInfoListObj As clsVSStageObj
Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
'===============================
VBEStageNum(0) = 46
VBEStageNum(1) = -2 '����ˮ`��(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = num '����ˮ`�H���s��
VBEStageNum(3) = 3 '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
VBEStageNum(4) = livecom(����ݾ��H��������(2, num)) '����ˮ`���ƭ�(�{��HP)
stageInfoListObj.Argument = livecom(����ݾ��H��������(2, num))  '����ˮ`���ƭ�
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "2" '����ˮ`��(1.�ϥΪ�/2.�q��)
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + str(num) '����ˮ`�H���s��
stageInfoListObj.Argument = stageInfoListObj.Argument + "%" + "3" '����ˮ`���Φ�(1.���/2.����/3.�ߧY���`)
'===========================���涥�q���J�I(46)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 46, 1
'============================
If stageInfoListObj.CommandStr = "PersonBloodControl" Then
    If stageInfoListObj.Value = "BLOODOFF" Then
        Exit Sub
    Else
        Dim tmpstr() As String
        tmpstr = Split(stageInfoListObj.Value, "%")
        If UBound(tmpstr) = 1 And tmpstr(0) = "BLOODCHANGE" Then
            Exit Sub
        End If
    End If
End If
'============================
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
End Sub
Sub ����_��_�ϥΪ�(ByVal num As Integer)
If liveus(����ݾ��H��������(1, num)) > 0 Then Exit Sub
'===============================
Dim stageInfoListObj As New clsVSStageObj
Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
VBEStageNum(0) = 49
VBEStageNum(1) = -1 '����_����(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = num '����_���H���s��
'===========================���涥�q���J�I(49)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 49, 1
'============================
Dim tmpflag As Boolean
tmpflag = False
If stageInfoListObj.CommandStr = "PersonResurrect" Then
    If stageInfoListObj.Value = "OFF" Then
        tmpflag = True
    End If
End If

If tmpflag = False Then
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
Dim stageInfoListObj As New clsVSStageObj
Set stageInfoListObj = ���涥�q�t����.VBEVSStageInfoList(���涥�q�t����.VBEVSStageInfoList.Count)
VBEStageNum(0) = 49
VBEStageNum(1) = -2 '����_����(1.�ϥΪ�/2.�q��)
VBEStageNum(2) = num '����_���H���s��
'===========================���涥�q���J�I(49)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 49, 1
'============================
Dim tmpflag As Boolean
tmpflag = False
If stageInfoListObj.CommandStr = "PersonResurrect" Then
    If stageInfoListObj.Value = "OFF" Then
        tmpflag = True
    End If
End If

If tmpflag = False Then
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
Function �C������P����Ыصo��() As Integer
Dim i As Integer

�԰��t����.ActionCardTotNum = �԰��t����.ActionCardTotNum + 1
i = �԰��t����.ActionCardTotNum

Load FormMainMode.card(i)
Set FormMainMode.card(i).Container = FormMainMode.PEAttackingForm
FormMainMode.card(i).Left = 240
FormMainMode.card(i).Top = 960
FormMainMode.card(i).Visible = False
FormMainMode.card(i).CardEventType = False
FormMainMode.card(i).LocationType = 0

�C������P����Ыصo�� = �԰��t����.ActionCardTotNum
End Function
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
FormMainMode.PEAFInterface.CardNum = BattleCardNum
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
            For ckl = 1 To �԰��t����.ActionCardTotNum
                FormMainMode.card(ckl).CardEnabledType = False
            Next
            FormMainMode.PEAFInterface.BnOKEnabled False
            ���ݮɶ���C(2).Add 47
            FormMainMode.���ݮɶ�_2.Enabled = True
        ElseIf Formsetting.chkusenewaipersonauto.Value = 1 Then
            For ckl = 1 To �԰��t����.ActionCardTotNum
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
Sub �d�P�P�ﶰ�X��(ByRef tmpcard As clsActionCard, ByVal torigc As Integer, ByVal tnewc As Integer)
Dim tmpindexobj As clsCollectionIndex
Set tmpindexobj = �԰��t����.CardDeckCollection(0)(CStr(tmpcard.CardNum))

�԰��t����.CardDeckCollection(torigc).Remove CStr(tmpcard.CardNum)
�԰��t����.CardDeckCollection(tnewc).Add tmpcard, CStr(tmpcard.CardNum)
'=========���ާ�s
tmpindexobj.CollectionIndex = tnewc
End Sub
Function �d�P�P�ﶰ�X����_CollectionIndex(ByVal tmpindex As Variant) As Integer
Dim tmpindexobj As clsCollectionIndex
Set tmpindexobj = �԰��t����.CardDeckCollection(0)(tmpindex)
�d�P�P�ﶰ�X����_CollectionIndex = tmpindexobj.CollectionIndex
End Function
Function �d�P�P�ﶰ�X����_CardNum(ByVal tmpindex As Variant) As Integer
Dim tmpindexobj As clsCollectionIndex
Set tmpindexobj = �԰��t����.CardDeckCollection(0)(tmpindex)
�d�P�P�ﶰ�X����_CardNum = tmpindexobj.CardNum
End Function
Function �d�P�P�ﶰ�X����_Index(ByVal tmpindex As Variant) As Integer
Dim tmpindexobj As clsCollectionIndex
Set tmpindexobj = �԰��t����.CardDeckCollection(0)(tmpindex)
�d�P�P�ﶰ�X����_Index = tmpindexobj.Index
End Function
