Attribute VB_Name = "���z��AI�t����"
Public cardcountAInum() As String  '���εP�p��Ȯɰ򥻸��(��x�i,1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.�P�s��)
Public cardcountAInumMOV() As String  '���εP�p��Ȯɰ򥻸��-���ʶ��q��-�쥻(��x�i,1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.�P�s��)
Dim cardAIn() As Integer '�ƦC�զX�p��Ȯ��ܼ�
Dim cardAInumans As String '�ƦC�զX�p��Ȯ��ܼ�
Public cardAInumnm() As String '�ƦC�զX�p��̲׼ƭ�
Public cardAInumFinal() As Integer '�ƦC�զX�p��̲״����
Public cardAInumFinal2() As Integer '�ƦC�զX�p��̲״����-�ƦC��
Public cardAInumcase(1 To 5, 1 To 2) As Integer '���εP�p��έp���(1.ATK-�C/2.DEF/3.MOV/4.SPE/5.ATK-�j,1.�զX�U�̧C�ƭ�/2.�զX�U�̰��ƭ�)
Public cardAInumcaseperson() As Integer '���εP�p��έp�Ȯɸ��-�ӧO�զX
Public cardAInumuscom As Integer '��P�֦��̵P�ưO���Ȯ��ܼ�
Public cardAITotalNUM As Integer '�ƦC�զX�p���`�@�զX��
Public cardAInumcasepersonTER() As Integer '���εP�p��έp�Ȯɸ��-�ӧO�զX-�ӧO�d���ƭȭp�Ʋέp
Public cardAInumselect1 As Integer  '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰������
Public cardAInumselect4 As Integer  '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰��ӧO�[�`�����
Public cardAInumselect2 As String '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰�����ȤU�s����-��l
Public cardAInumselect3() As String '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰�����ȤU�s����-�}�C
Public cardAInumchoose As Integer '���εP�p��̲׿�ܲզX�s��
Public cardAInumMOVmain(1 To 2, 1 To 15) As String 'AI-���ʶ��q��-�զX�Ȯɬ���
Public cardAInumMOVnm() As String 'AI-���ʶ��q��-���V��-�p��ƦC�զX��Ȯɬ���
Public cardAInumMOVnmtot() As String 'AI-���ʶ��q��-���V��-�`�@�ƦC�զX�������ƼȮɬ���
Public cardAInumMOVFinal(1 To 3) As String 'AI-���ʶ��q��-���V��-�̲׵��G������(1.�̲ױƦC�զX��/2.�̲ױƦC�զX�s��/3.�̲׿�w�ؼжZ��[1.��/2.��])
Public �O�_���ʶ��q����p�P�_�{�� As Boolean 'AI-���ʶ��q��-�O�_�����p�P�_�{�ǼаO��
Public cardAInumOvertenrecord() As Integer 'AI�޾ɵ{��-�W�X�P�i��-�P�����Ȯ��ܼ�(1~10.�P�s��)
Public personatkingtfr(1 To 5) As Integer '�p��ӧO�ޯ�-�O�_��Ex��(1~4.(1)��/(2)�L,5.�O�_���ʦL)
Sub ���z��AI�t�έp��_�@���q_��l(ByVal pagenumber As Integer)
Erase cardcountAInum
Erase cardAInumnm
Erase cardAInumcase
Erase cardAInumselect3
cardAInumans = ""
cardAInumselect1 = 0
cardAInumselect4 = 0
cardAInumselect2 = ""
cardAInumchoose = 0
cardAInumuscom = pagenumber
cardAITotalNUM = 2 ^ cardAInumuscom
ReDim cardcountAInum(1 To cardAInumuscom, 1 To 5) As String
ReDim cardAInumcaseperson(1 To cardAITotalNUM, 1 To 2, 1 To 15) As Integer
ReDim cardAInumcasepersonTER(1 To cardAITotalNUM, 1 To 5, 1 To 10) As Integer
ReDim cardAInumFinal(1 To cardAITotalNUM, 1 To 4) As Integer
ReDim cardAInumFinal2(1 To cardAITotalNUM, 1 To 4) As Integer
'=========�p�⥿�ϭ��ƦC�զX�ƭ�
���z��AI�t����.�ƦC�զX�p�� pagenumber
End Sub
Sub ���z��AI�t�έp��_�@���q_���o�P�����(ByVal �O�_�@�� As Boolean, ByVal uscom As Integer)
If �O�_�@�� = True Then
        '=========�^���ثe�P�����
        Select Case uscom
            Case 1
                �԰��t����.�X�P���ǭp��_�ϥΪ�_��P
            Case 2
                �԰��t����.�X�P���ǭp��_�q��_��P
        End Select
        Dim w As Integer '�Ȯ��ܼ�
        w = 2 * uscom '(2-�ϥΪ̤�P/4-�q����P)
        For i = 1 To pageglead(uscom)
            cardcountAInum(i, 5) = �X�P���ǲέp�Ȯ��ܼ�(w, i, 2)
            cardcountAInum(i, 1) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 1)
            cardcountAInum(i, 2) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 2)
            cardcountAInum(i, 3) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 3)
            cardcountAInum(i, 4) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 4)
        Next
End If
'======================
'���z��AI�t����.�ƦC�զX�έp�ƭȭp��_��P�`�p
���z��AI�t����.�ƦC�զX�έp�ƭȭp��_�����ۦ��k�h�����ƲզX
���z��AI�t����.�ƦC�զX�έp�ƭȭp��_�ӧO�զX
End Sub
Sub ���z��AI�t�έp��_�G���q_�p������_��l(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
Dim wnum As Integer, whnum As Integer '�Ȯ��ܼ�
Select Case turn
    Case 1 '===�������q
         If uscom = 1 Then whnum = atkus(����H����ԤH��(1, 2)) Else whnum = atkcom(����H����ԤH��(2, 2))
         '==========================
         For i = 0 To (cardAITotalNUM) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If (cardcountAInum(j, 1) = a1a And movecpre = 1) Or (cardcountAInum(j, 1) = a5a And movecpre > 1) Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                      Case 1
                          If (cardcountAInum(j, 3) = a1a And movecpre = 1) Or (cardcountAInum(j, 3) = a5a And movecpre > 1) Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
             If Val(wnum) > 0 Then
                 cardAInumFinal(i + 1, 1) = Val(cardAInumFinal(i + 1, 1)) + Val(whnum)
             End If
         Next
    Case 2  '===���m���q
         For i = 0 To (cardAITotalNUM) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If cardcountAInum(j, 1) = a2a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                      Case 1
                          If cardcountAInum(j, 3) = a2a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = 1
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
         Next
    Case 3  '===���ʶ��q
         For i = 0 To (cardAITotalNUM) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If cardcountAInum(j, 1) = a3a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                          End If
                      Case 1
                          If cardcountAInum(j, 3) = a3a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
'                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
         Next
End Select
End Sub
Sub ���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ�(ByVal name As String, ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
'���z��AI�t����.�ˬd�H���ޯ�O�_��EX�� uscom, name
'If personatkingtfr(5) = 1 Then
'   Exit Sub '���ʦL���A�ɵL�k�o�ʧޯ�
'End If
'Select Case name
'     Case "��B�����S"
'           ���z��AI�H����.��B�����S turn, movecpre, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "����"
'           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "���"
'           ���z��AI�H����.��� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�j�|�˺��h"
'           ���z��AI�H����.�j�|�˺��h turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "���["
'           ���z��AI�H����.���[ turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�v��L"
'           ���z��AI�H����.�v��L turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "C.C."
'           ���z��AI�H����.CC turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "��ܵY"
'           ���z��AI�H����.��ܵY turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "����"
'           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "����"
'           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "������"
'           ���z��AI�H����.������ turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "��̬d�w"
'           ���z��AI�H����.��̬d�w turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "������"
'           ���z��AI�H����.������ turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�Q��"
'           ���z��AI�H����.�Q�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�L���S"
'           ���z��AI�H����.�L���S turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "���纸"
'           ���z��AI�H����.���纸 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "������S"
'           ���z��AI�H����.������S turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�w�ǥ���"
'           ���z��AI�H����.�w�ǥ��� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "����P��"
'           ���z��AI�H����.����P�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�h�g�H"
'           ���z��AI�H����.�h�g�H turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�Ǧh"
'           ���z��AI�H����.�Ǧh turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "���_�i���h"
'           ���z��AI�H����.���_�i���h turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�S�{��"
'           ���z��AI�H����.�S�{�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "����"
'           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "���Y�F"
'           ���z��AI�H����.���Y�F turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "��"
'           ���z��AI�H����.�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "ù��Y"
'           ���z��AI�H����.ù��Y turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�����g"
'           ���z��AI�H����.�����g turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�J�y"
'           ���z��AI�H����.�J�y turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�����i"
'           ���z��AI�H����.�����i turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'     Case "�ײ��d"
'           ���z��AI�H����.�ײ��d turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
'End Select
'===========================���涥�q���J�I(99)
���z��AI�t����.���z��AI�t��_���涥�q99_�D�ʧޯ���� uscom, turn, movecpre
'============================
End Sub
Sub �ƦC�զX�p��(ByVal qnum As Integer)
'===========
ReDim cardAIn(1 To Val(qnum))
Erase cardAInumnm
cardAInumans = ""
Dim s As Integer
For i = 1 To qnum   '���]�϶��ƭ�
    cardAIn(i) = 0
Next
s = 1
'================
Do
    For i = qnum To 1 Step -1
        cardAInumans = cardAInumans & cardAIn(i)
    Next
    '================
    cardAIn(1) = cardAIn(1) + 1
    ���z��AI�t����.�ƦC�զX�p��_�϶��i�� qnum '�@[qnum]���
    '================
    s = s + 1
    cardAInumans = cardAInumans & "="
Loop Until s > (2 ^ qnum)
cardAInumnm = Split(cardAInumans, "=")

End Sub
Sub �ƦC�զX�p��_�϶��i��(ByVal num As Integer)
For i = 1 To num - 1
    If cardAIn(i) = 2 Then
        cardAIn(i + 1) = cardAIn(i + 1) + 1
        cardAIn(i) = 0
    End If
Next

End Sub
Sub �ƦC�զX�έp�ƭȭp��_��P�`�p()
Dim we As Integer  '�Ȯ��ܼ�
For i = 1 To cardAInumuscom
    For j = 1 To 2
        we = 2 * j
        Select Case cardcountAInum(i, j)
             Case a1a
                  If cardcountAInum(i, we) < cardAInumcase(1, 1) Or cardAInumcase(1, 1) = 0 Then
                      cardAInumcase(1, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(1, 2) Or cardAInumcase(1, 2) = 0 Then
                      cardAInumcase(1, 2) = cardcountAInum(i, we)
                  End If
             Case a2a
                  If cardcountAInum(i, we) < cardAInumcase(2, 1) Or cardAInumcase(2, 1) = 0 Then
                      cardAInumcase(2, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(2, 2) Or cardAInumcase(2, 2) = 0 Then
                      cardAInumcase(2, 2) = cardcountAInum(i, we)
                  End If
             Case a3a
                  If cardcountAInum(i, we) < cardAInumcase(3, 1) Or cardAInumcase(3, 1) = 0 Then
                      cardAInumcase(3, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(3, 2) Or cardAInumcase(3, 2) = 0 Then
                      cardAInumcase(3, 2) = cardcountAInum(i, we)
                  End If
             Case a4a
                  If cardcountAInum(i, we) < cardAInumcase(4, 1) Or cardAInumcase(4, 1) = 0 Then
                      cardAInumcase(4, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(4, 2) Or cardAInumcase(4, 2) = 0 Then
                      cardAInumcase(4, 2) = cardcountAInum(i, we)
                  End If
             Case a5a
                  If cardcountAInum(i, we) < cardAInumcase(5, 1) Or cardAInumcase(5, 1) = 0 Then
                      cardAInumcase(5, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(5, 2) Or cardAInumcase(5, 2) = 0 Then
                      cardAInumcase(5, 2) = cardcountAInum(i, we)
                  End If
        End Select
    Next
Next
End Sub
Sub �ƦC�զX�έp�ƭȭp��_�ӧO�զX()
Dim we As Integer '�Ȯ��ܼ�
For i = 1 To cardAITotalNUM
    For j = 1 To cardAInumuscom
        Select Case Mid(cardAInumnm(i - 1), j, 1)
            Case 0
                 we = 2
                  Select Case cardcountAInum(j, 1)
                     Case a1a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 1) Or cardAInumcaseperson(i, 1, 1) = 0 Then
                              cardAInumcaseperson(i, 1, 1) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 2) Or cardAInumcaseperson(i, 1, 2) = 0 Then
                              cardAInumcaseperson(i, 1, 2) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 11) = cardAInumcaseperson(i, 1, 11) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 1, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 1, cardcountAInum(j, we))) + 1
                     Case a2a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 3) Or cardAInumcaseperson(i, 1, 3) = 0 Then
                              cardAInumcaseperson(i, 1, 3) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 4) Or cardAInumcaseperson(i, 1, 4) = 0 Then
                              cardAInumcaseperson(i, 1, 4) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 12) = cardAInumcaseperson(i, 1, 12) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 2, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 2, cardcountAInum(j, we))) + 1
                     Case a3a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 5) Or cardAInumcaseperson(i, 1, 5) = 0 Then
                              cardAInumcaseperson(i, 1, 5) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 6) Or cardAInumcaseperson(i, 1, 6) = 0 Then
                              cardAInumcaseperson(i, 1, 6) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 13) = cardAInumcaseperson(i, 1, 13) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 3, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 3, cardcountAInum(j, we))) + 1
                     Case a4a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 7) Or cardAInumcaseperson(i, 1, 7) = 0 Then
                              cardAInumcaseperson(i, 1, 7) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 8) Or cardAInumcaseperson(i, 1, 8) = 0 Then
                              cardAInumcaseperson(i, 1, 8) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 14) = cardAInumcaseperson(i, 1, 14) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 4, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 4, cardcountAInum(j, we))) + 1
                     Case a5a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 9) Or cardAInumcaseperson(i, 1, 9) = 0 Then
                              cardAInumcaseperson(i, 1, 9) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 10) Or cardAInumcaseperson(i, 1, 10) = 0 Then
                              cardAInumcaseperson(i, 1, 10) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 15) = cardAInumcaseperson(i, 1, 15) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 5, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 5, cardcountAInum(j, we))) + 1
                End Select
            Case 1
                 we = 4
                  Select Case cardcountAInum(j, 3)
                     Case a1a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 1) Or cardAInumcaseperson(i, 1, 1) = 0 Then
                              cardAInumcaseperson(i, 1, 1) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 2) Or cardAInumcaseperson(i, 1, 2) = 0 Then
                              cardAInumcaseperson(i, 1, 2) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 11) = cardAInumcaseperson(i, 1, 11) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 1, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 1, cardcountAInum(j, we))) + 1
                     Case a2a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 3) Or cardAInumcaseperson(i, 1, 3) = 0 Then
                              cardAInumcaseperson(i, 1, 3) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 4) Or cardAInumcaseperson(i, 1, 4) = 0 Then
                              cardAInumcaseperson(i, 1, 4) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 12) = cardAInumcaseperson(i, 1, 12) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 2, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 2, cardcountAInum(j, we))) + 1
                     Case a3a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 5) Or cardAInumcaseperson(i, 1, 5) = 0 Then
                              cardAInumcaseperson(i, 1, 5) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 6) Or cardAInumcaseperson(i, 1, 6) = 0 Then
                              cardAInumcaseperson(i, 1, 6) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 13) = cardAInumcaseperson(i, 1, 13) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 3, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 3, cardcountAInum(j, we))) + 1
                     Case a4a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 7) Or cardAInumcaseperson(i, 1, 7) = 0 Then
                              cardAInumcaseperson(i, 1, 7) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 8) Or cardAInumcaseperson(i, 1, 8) = 0 Then
                              cardAInumcaseperson(i, 1, 8) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 14) = cardAInumcaseperson(i, 1, 14) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 4, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 4, cardcountAInum(j, we))) + 1
                     Case a5a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 9) Or cardAInumcaseperson(i, 1, 9) = 0 Then
                              cardAInumcaseperson(i, 1, 9) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 10) Or cardAInumcaseperson(i, 1, 10) = 0 Then
                              cardAInumcaseperson(i, 1, 10) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 15) = cardAInumcaseperson(i, 1, 15) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 5, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 5, cardcountAInum(j, we))) + 1
                End Select
        End Select
    Next
Next
End Sub
Sub �ƦC�զX�έp�ƭȭp��_�����ۦ��k�h�����ƲզX()
Dim �d���w���ƼаO() As Boolean
Dim �d���ۦ��аO() As Boolean
Dim �d���ƦC�զX�ƻs() As String
Dim CardChooseNum As Integer, �d���ۦ��аO�� As Integer
Dim c1 As String, c2 As String
ReDim �d���w���ƼаO(1 To cardAITotalNUM) As Boolean
ReDim �d���ƦC�զX�ƻs(cardAITotalNUM - 1) As String
'ReDim �d���ۦ��аO(1 To cardAInumuscom) As Boolean
'=========================================
For i = 1 To cardAITotalNUM
    For j = i - 1 To 1 Step -1
        �d���ۦ��аO�� = 0
        ReDim �d���ۦ��аO(1 To cardAInumuscom) As Boolean
'        If �d���w���ƼаO(j) = False Then
            For k = 1 To cardAInumuscom
                c1 = Mid(cardAInumnm(i - 1), k, 1)
                For p = 1 To cardAInumuscom
                     c2 = Mid(cardAInumnm(j - 1), p, 1)
                     If c1 = "0" And c2 = "0" Then
                         If cardcountAInum(k, 1) = cardcountAInum(p, 1) And cardcountAInum(k, 2) = cardcountAInum(p, 2) And �d���ۦ��аO(p) = False Then
                             �d���ۦ��аO(p) = True
                             �d���ۦ��аO�� = �d���ۦ��аO�� + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     ElseIf c1 = "0" And c2 = "1" Then
                         If cardcountAInum(k, 1) = cardcountAInum(p, 3) And cardcountAInum(k, 2) = cardcountAInum(p, 4) And �d���ۦ��аO(p) = False Then
                             �d���ۦ��аO(p) = True
                             �d���ۦ��аO�� = �d���ۦ��аO�� + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     ElseIf c1 = "1" And c2 = "0" Then
                         If cardcountAInum(k, 3) = cardcountAInum(p, 1) And cardcountAInum(k, 4) = cardcountAInum(p, 2) And �d���ۦ��аO(p) = False Then
                             �d���ۦ��аO(p) = True
                             �d���ۦ��аO�� = �d���ۦ��аO�� + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     ElseIf c1 = "1" And c2 = "1" Then
                         If cardcountAInum(k, 3) = cardcountAInum(p, 3) And cardcountAInum(k, 4) = cardcountAInum(p, 4) And �d���ۦ��аO(p) = False Then
                             �d���ۦ��аO(p) = True
                             �d���ۦ��аO�� = �d���ۦ��аO�� + 1
                             p = cardAInumuscom 'Exit For
                         End If
                     End If
                Next
            Next
'        End If
        If �d���ۦ��аO�� = cardAInumuscom Then
            �d���w���ƼаO(i) = True
            CardChooseNum = CardChooseNum + 1
            j = 1 'Exit For
        End If
    Next
Next
'MsgBox "�����ۦ����ƲզX��:" & CardChooseNum
'============================
If CardChooseNum > 0 Then
    '====================
    For i = 0 To cardAITotalNUM - 1
        �d���ƦC�զX�ƻs(i) = cardAInumnm(i)
    Next
    ReDim cardAInumnm((cardAITotalNUM - 1) - CardChooseNum) As String
    '====================
    k = 0
    For i = 0 To cardAITotalNUM - 1
        If �d���w���ƼаO(i + 1) = False Then
            cardAInumnm(k) = �d���ƦC�զX�ƻs(i)
            k = k + 1
        End If
    Next
    cardAITotalNUM = cardAITotalNUM - CardChooseNum
    '=====================
    ReDim cardAInumcaseperson(1 To cardAITotalNUM, 1 To 2, 1 To 15) As Integer
    ReDim cardAInumcasepersonTER(1 To cardAITotalNUM, 1 To 5, 1 To 10) As Integer
    ReDim cardAInumFinal(1 To cardAITotalNUM, 1 To 4) As Integer
    ReDim cardAInumFinal2(1 To cardAITotalNUM, 1 To 4) As Integer
    '=====================
End If
End Sub
Sub ���z��AI�t�έp��_�T���q_�έp�ƦC()
'=================�ƻs���e
For k = 1 To cardAITotalNUM
    cardAInumFinal2(k, 1) = cardAInumFinal(k, 1)
    cardAInumFinal2(k, 2) = cardAInumFinal(k, 2)
Next
'=================
Dim wer As Integer, wes As Integer
For i = cardAITotalNUM To 1 Step -1
    For j = 1 To i - 1
        If Val(cardAInumFinal2(j, 1)) < Val(cardAInumFinal2(j + 1, 1)) Then
            wer = cardAInumFinal2(j + 1, 1)
            wes = cardAInumFinal2(j + 1, 2)
            cardAInumFinal2(j + 1, 1) = cardAInumFinal2(j, 1)
            cardAInumFinal2(j + 1, 2) = cardAInumFinal2(j, 2)
            cardAInumFinal2(j, 1) = wer
            cardAInumFinal2(j, 2) = wes
        End If
    Next
Next
End Sub
Sub ���z��AI�t�έp��_�|���q_���_1_��l()
For i = 1 To cardAITotalNUM
    If Val(cardAInumFinal2(i, 1)) > Val(cardAInumselect1) Then
        cardAInumselect1 = cardAInumFinal2(i, 1)
    End If
Next
'====================
If cardAInumselect1 < 0 Then cardAInumselect1 = 0 '�h���`����Ȭ��t�Ƥ��զX
'====================
For i = 1 To cardAITotalNUM
    If cardAInumFinal2(i, 1) = cardAInumselect1 Then
        cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
    End If
Next
'====================
If cardAInumselect2 = "" Then  '�S������զX�ŦX����
    cardAInumselect2 = "-10=-10"
End If
End Sub
Sub ���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1()
'cardAInumselect3 = Split(cardAInumselect2, "=")
'If UBound(cardAInumselect3) > 1 Then
'    For i = 1 To cardAITotalNUM
'        For j = 1 To cardAInumuscom
'             If cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, j) < 0 Then
'                 cardAInumFinal2(i, 3) = 1
'             End If
'             cardAInumFinal2(i, 4) = Val(cardAInumFinal2(i, 4)) + Val(cardAInumcaseperson(cardAInumFinal2(i, 2), 2, j))
'        Next
'    Next
'    '===============
'    Erase cardAInumselect3
'    cardAInumselect2 = ""
'    '======
'    For i = 1 To cardAITotalNUM
'        If cardAInumFinal2(i, 1) = cardAInumselect1 And cardAInumFinal2(i, 3) = 0 Then
'            cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
'        End If
'    Next
'    cardAInumselect3 = Split(cardAInumselect2, "=")
'End If
End Sub
Sub ���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2()
cardAInumselect3 = Split(cardAInumselect2, "=")
If UBound(cardAInumselect3) > 1 Then
    Dim wer As Integer
    wer = cardAInumFinal2(1, 4)  '�ؼп���̰��P�i�ơA�`����ȳ̰����զX
    '===============
    For i = 1 To cardAITotalNUM
         If Val(cardAInumFinal2(i, 4)) > Val(wer) And cardAInumFinal2(i, 1) = cardAInumselect1 Then
             wer = cardAInumFinal2(i, 4)
         End If
    Next
    '===============
    Erase cardAInumselect3
    cardAInumselect2 = ""
    cardAInumselect4 = wer
    '======
    For i = 1 To cardAITotalNUM
        If cardAInumFinal2(i, 4) = wer And cardAInumFinal2(i, 1) = cardAInumselect1 Then
            cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
        End If
    Next
    cardAInumselect3 = Split(cardAInumselect2, "=")
End If
End Sub
Sub ���z��AI�t�έp��_�|���q_���_3_��ܲզX()
If UBound(cardAInumselect3) > 1 Then
    Dim wtr As Integer '�Ȯ��ܼ�
    wtr = Int(Rnd() * UBound(cardAInumselect3)) + 1
    cardAInumchoose = cardAInumselect3(wtr)
Else
    cardAInumchoose = cardAInumselect3(1)
End If
End Sub
Sub ���z��AI�t�έp��_�̫ᶥ�q_����P(ByVal choose As Integer, ByVal uscom As Integer)
Dim wer As Integer '�Ȯ��ܼ�
If choose = 1 Then
    wer = 0
Else
    wer = 1
End If
'=================
Dim pu As Integer '�Ȯ��ܼ�
'=====
If cardAInumchoose = -10 Then  '==�S������զX�ŦX�X�P����
    Exit Sub
End If
'=======================�p�զX�ŦX�X�P���󪺸�
Select Case uscom
     Case 1 '==�ϥΪ̤�
            For i = 1 To UBound(cardcountAInum, 1)
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 4
                    ElseIf cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 3
                    End If
            Next
     Case 2 '==�q����
            For i = 1 To UBound(cardcountAInum, 1)
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        cspce = pagecardnum(pu, 1)
                        cspme = pagecardnum(pu, 2)
                        pagecardnum(pu, 1) = pagecardnum(pu, 3)
                        pagecardnum(pu, 2) = pagecardnum(pu, 4)
                        pagecardnum(pu, 3) = cspce
                        pagecardnum(pu, 4) = cspme
                        If pageonin(pu) = 2 Then
                           pageonin(pu) = 1
                        Else
                           pageonin(pu) = 2
                        End If
                    End If
                    If cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 1
                    End If
            Next
End Select
End Sub
Sub ���z��AI�t�έp��_�ȮɶץX(ByVal uscom As Integer)
'If Formsetting.checktest.Value = 1 Then
''    Open App.Path & "\test\out1.txt" For Output As #1
'    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & BattleTurn & "turn_" & �԰��t����.turnatk & "_" & uscom & "_1.txt" For Output As #1
'    For i = 1 To cardAITotalNUM
'        Print #1, cardAInumnm(Val(cardAInumFinal2(i, 2)) - 1) & "=" & cardAInumFinal2(i, 1) & "/" & cardAInumFinal2(i, 4) & "#" & cardAInumFinal2(i, 2) & "@";
'        For k = 1 To cardAInumuscom
'            Print #1, cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, k) & "=";
'        Next
'        Print #1,
'    Next
'    Close
'    'MsgBox "�w�ץX����1"
'End If
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_����1(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer)
���z��AI�t����.���z��AI�t�έp��_�@���q_��l uscom
���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turn, movecpre, uscom
���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turn, movecpre, uscom
���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
���z��AI�t����.���z��AI�t�έp��_�ȮɶץX uscom
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_���(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer)
Dim CardMaxNum As Integer
If Formsetting.chksetcomaipagenum.Value = 1 Then
    CardMaxNum = Val(Formsetting.�ۭqAI��P�i��.Text)
Else
    CardMaxNum = 7
End If
'========================
If Val(pageglead(uscom)) > CardMaxNum Then
    ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_�W�X�P�i�� uscom, turn, name, movecpre, choose, CardMaxNum
ElseIf Val(pageglead(uscom)) > 0 And Val(pageglead(uscom)) <= CardMaxNum Then
    ���z��AI�t����.���z��AI�t�έp��_�@���q_��l pageglead(uscom)
    ���z��AI�t����.���z��AI�t�έp��_�@���q_���o�P����� True, uscom
    ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turn, movecpre, uscom
    ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turn, movecpre, uscom
    ���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_1_��l
'    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2
    ���z��AI�t����.���z��AI�t�έp��_�ȮɶץX uscom
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_3_��ܲզX
    If turn = 3 And cardAInumchoose > 0 Then
        ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_���ʶ��q�� uscom, turn, name, movecpre, choose, pageglead(uscom)
    Else
        ���z��AI�t����.���z��AI�t�έp��_�̫ᶥ�q_����P choose, uscom
    End If
End If
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_���ʶ��q��(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer, ByVal pagenumber As Integer)
If Val(pagenumber) > 0 Then
    Select Case ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P���(uscom)
        Case True
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�@���q_�ǳƶi����
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�G���q_�i����p�ƦC�զX��p�� pagenumber, uscom
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�T���q_�i����p����ȭp�� uscom, name, choose, movecpre, pagenumber
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�|���q_�έp���p����ȤΧP�_ uscom
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�����q_����P choose, uscom, pagenumber
        Case False
'            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_�_�w��_�@���q_���]�����_�ӧO
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_�_�w��_�G���q_��ܦ�� uscom
            ���z��AI�t����.���z��AI�t�έp��_�̫ᶥ�q_����P choose, uscom
    End Select
End If
End Sub
Function ���z��AI�t��_�ثe�i���椧�H���P�_(ByVal name As String) As Boolean
If Formsetting.chkusenewai.Value = 0 Then
    ���z��AI�t��_�ثe�i���椧�H���P�_ = False
    Exit Function
End If
Select Case name
    Case "��B�����S"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�j�|�˺��h"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���["
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�v��L"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "C.C."
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "��ܵY"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "������"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "��̬d�w"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "������"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�Q��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�L���S"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���纸"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "������S"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�w�ǥ���"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����P��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�h�g�H"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�Ǧh"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���_�i���h"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�S�{��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���Y�F"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "ù��Y"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�����g"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�J�y"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�����i"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�ײ��d"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case Else
            ���z��AI�t��_�ثe�i���椧�H���P�_ = False
End Select
End Function
Function ���h��(ByVal num As Integer) As Single
Dim w As Double
w = 1
If num <> 0 Then
    For i = 1 To Val(num)
        w = Val(w) * Val(i)
    Next
Else
    w = 1
End If
���h�� = w
End Function
Function ���h��_��C(ByVal c1 As Integer, ByVal c2 As Integer) As Single

���h��_��C = (���z��AI�t����.���h��(c1) / ���z��AI�t����.���h��(Val(c1) - Val(c2))) / ���z��AI�t����.���h��(c2)

End Function
Sub ���z��AI�t�έp��_���ʶ��q��_���o�p�⤧�ƦC�զX(ByVal n1 As Integer, ByVal n2 As Integer)
Dim wtstr As String, wtall As Integer, wtpnum() As String, wtn As Integer
'===================
���z��AI�t����.�ƦC�զX�p�� n1
wtall = ���z��AI�t����.���h��_��C(n1, n2)
ReDim cardAInumMOVnm(1 To wtall) As String
'====================
For i = 1 To 2 ^ n1
    wtn = 0
    For j = 1 To n1
        If Val(Mid(cardAInumnm(i - 1), j, 1)) = 1 Then
            wtn = wtn + 1
        End If
    Next
    If wtn = n2 Then '==��n2�i�X�P���զX
        wtstr = wtstr & "=" & i
    End If
Next
wtpnum = Split(wtstr, "=")
'If UBound(wtpnum) = wtall Then
'    MsgBox wtstr
'    For i = 1 To UBound(wtpnum)
'        Debug.Print wtpnum(i) & "=" & cardAInumnm(wtpnum(i) - 1)
'    Next
'Else
'    MsgBox "����"
'End If
For i = 1 To UBound(wtpnum)
    cardAInumMOVnm(i) = cardAInumnm(wtpnum(i) - 1)
Next
End Sub
Function ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P���(ByVal uscom As Integer) As Boolean
Erase cardAInumMOVmain
Erase cardAInumMOVnm
Erase cardAInumMOVnmtot
Dim wtmovnum As Integer '�Ȯ��ܼ�
Dim buffobj As clsStatus
If cardAInumchoose = -10 Then
    ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P��� = False
    Exit Function
End If
'============�����ثe�զX
cardAInumMOVmain(1, 1) = cardAInumselect1
cardAInumMOVmain(1, 2) = cardAInumselect4
cardAInumMOVmain(1, 3) = cardAInumnm(cardAInumchoose - 1)
cardAInumMOVmain(1, 4) = cardAInumcaseperson(cardAInumchoose, 1, 13)
cardAInumMOVmain(1, 5) = cardAInumchoose
For i = 1 To cardAInumuscom
    cardAInumMOVmain(2, i) = cardAInumcaseperson(cardAInumchoose, 2, i)
Next
'==============�p�⦳�Ĳ��ʼ�
wtmovnum = cardAInumMOVmain(1, 4)
For Each n In �H�����`���A�C��(uscom, ����H����ԤH��(uscom, 2))
    Set buffobj = n
    If buffobj.Identifier = "BUFFN00302" Then
        wtmovnum = Val(wtmovnum) - buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00301" Then
        wtmovnum = Val(wtmovnum) + buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00801" Then
        wtmovnum = -100
    ElseIf buffobj.Identifier = "BUFFN00501" And _
        ((uscom = 1 And liveus(����H����ԤH��(uscom, 2)) = 1) Or (uscom = 2 And livecom(����H����ԤH��(uscom, 2)) = 1)) Then
        wtmovnum = -100
    End If
Next
'=====================
If wtmovnum >= 2 Then
    ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P��� = True
Else
    ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P��� = False
End If
End Function
Sub ���z��AI�t�έp��_���ʶ��q��_�_�w��_�@���q_���]�����_�ӧO()
'For i = 1 To cardAInumuscom
'    If cardAInumcaseperson(cardAInumchoose, 2, i) < 10 Then
'        cardAInumcaseperson(cardAInumchoose, 2, i) = 0
'        cardAInumMOVmain(2, i) = 0
'    End If
'Next
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�@���q_�ǳƶi����()
Dim wercnum As Integer, werct As String, werpnum As Integer
ReDim cardcountAInumMOV(1 To cardAInumuscom, 1 To 5) As String
�O�_���ʶ��q����p�P�_�{�� = True
For k = 1 To cardAInumuscom
    Select Case Mid(cardAInumMOVmain(1, 3), k, 1)
         Case 0
              If cardcountAInum(k, 1) = a3a And cardAInumMOVmain(2, k) = 0 Then
                  wercnum = Val(wercnum) + 1
                  werct = werct & "=" & k
'              ElseIf cardAInumMOVmain(2, k) >= 10 Then
'                 werpnum = Val(werpnum) + 1
              End If
         Case 1
              If cardcountAInum(k, 3) = a3a And cardAInumMOVmain(2, k) = 0 Then
                  wercnum = Val(wercnum) + 1
                  werct = werct & "=" & k
'              ElseIf cardAInumMOVmain(2, k) >= 10 Then
'                 werpnum = Val(werpnum) + 1
              End If
    End Select
    For q = 1 To 5
         cardcountAInumMOV(k, q) = cardcountAInum(k, q)
    Next
Next
'===============
'If Val(werpnum) >= 1 Then werpnum = 1
'===============
ReDim cardAInumMOVnmtot(0 To (2 ^ wercnum), 1 To 8) As String
cardAInumMOVnmtot(0, 1) = werct
cardAInumMOVnmtot(0, 2) = 1
cardAInumMOVnmtot(0, 3) = wercnum
cardAInumMOVnmtot(0, 4) = 1
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�G���q_�i����p�ƦC�զX��p��(ByVal pagenumber As Integer, ByVal uscom As Integer)
Dim weru As Integer, wernum As Integer, werqr As String
Dim werstru As String
Dim werpstr() As String
Dim wermovnm As Integer, wermovynm As Integer
'============�i����p�����ʵP�ƦC�զX�p��
For i = 1 To Val(cardAInumMOVnmtot(0, 3))
       ���z��AI�t�έp��_���ʶ��q��_���o�p�⤧�ƦC�զX Val(cardAInumMOVnmtot(0, 3)), i
       weru = 1
       wernum = ���h��_��C(Val(cardAInumMOVnmtot(0, 3)), i)
        For k = Val(cardAInumMOVnmtot(0, 2)) To (Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)) - 1
             cardAInumMOVnmtot(k, 1) = cardAInumMOVnm(weru)
             weru = Val(weru) + 1
        Next
        cardAInumMOVnmtot(0, 2) = Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)
Next
'=====================�i��Ѿl���ʵP���ƦC�զX���X
'werpstr = Split(cardAInumMOVnmtot(1, 1), "=")
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    weru = 0
    werstru = ""
    wermovnm = 0
    wermovynm = 0
    For k = 1 To pagenumber
        Select Case Mid(cardAInumMOVmain(1, 3), k, 1)
              Case 0
                    If cardcountAInum(k, 1) = a3a And i <= (2 ^ Val(cardAInumMOVnmtot(0, 3)) - 1) And cardAInumMOVmain(2, k) = 0 Then
                        weru = Val(weru) + 1
                        If Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 1 Then
                            werstru = werstru & "1"
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
                            wermovynm = Val(wermovynm) + 1
                        ElseIf Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 0 Then
'                            werstru = werstru & "1"
'                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
'                            wermovynm = Val(wermovynm) + 1
'                        Else
                            werstru = werstru & "n"
                        End If
                    ElseIf cardAInumMOVmain(2, k) = 1 Then
                        werstru = werstru & "1"
                        If cardcountAInum(k, 1) = a3a Then
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
                            wermovynm = Val(wermovynm) + 1
                        End If
                    Else
                        werstru = werstru & "n"
                    End If
              Case 1
                    If cardcountAInum(k, 3) = a3a And i <= (2 ^ Val(cardAInumMOVnmtot(0, 3)) - 1) And cardAInumMOVmain(2, k) = 0 Then
                        weru = Val(weru) + 1
                        If Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 1 Then
                            werstru = werstru & "1"
                            wermovynm = Val(wermovynm) + 1
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
                        ElseIf Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 0 Then
'                            werstru = werstru & "1"
'                            wermovynm = Val(wermovynm) + 1
'                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
'                        Else
                            werstru = werstru & "n"
                        End If
                    ElseIf cardAInumMOVmain(2, k) = 1 Then
                        werstru = werstru & "1"
                        If cardcountAInum(k, 3) = a3a Then
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
                            wermovynm = Val(wermovynm) + 1
                        End If
                    Else
                        werstru = werstru & "n"
                    End If
        End Select
    Next
    cardAInumMOVnmtot(i, 2) = werstru
    cardAInumMOVnmtot(i, 6) = wermovnm
    cardAInumMOVnmtot(i, 7) = wermovynm
Next
'=========================���եζץX
'If Formsetting.checktest.Value = 1 Then
''    Open App.Path & "\test\out2.txt" For Output As #1
'    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & BattleTurn & "turn_" & �԰��t����.turnatk & "_" & uscom & "_2.txt" For Output As #1
'    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
'        Print #1, cardAInumMOVnmtot(i, 2)
'    Next
'    Print #1, cardAInumMOVmain(1, 5) & "=" & cardAInumMOVmain(1, 3)
'    Close
'    'MsgBox "�w�ץX����2"
'End If
'==============================
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�T���q_�i����p����ȭp��(ByVal uscom As Integer, ByVal name As String, ByVal choose As Integer, ByVal movecpre As Integer, ByVal pagenumber As Integer)
Dim weru As Integer, wertp As Integer, movecpren As Integer, turnm As Integer, werucount As Boolean
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    For k = 1 To 2
         '===========�N����ಾ�ܫݹB����
         weru = 0
         For wp = 1 To pagenumber
              If Mid(cardAInumMOVnmtot(i, 2), wp, 1) = "n" Then
                  weru = Val(weru) + 1
              End If
         Next
         If Val(weru) > 0 Then
                 ���z��AI�t����.���z��AI�t�έp��_�@���q_��l weru
                 wertp = 0
                 '=======
                 For q = 1 To pagenumber
                     If Mid(cardAInumMOVnmtot(i, 2), q, 1) = "n" Then
                           wertp = Val(wertp) + 1
                           For wds = 1 To 5
                                 cardcountAInum(wertp, wds) = cardcountAInumMOV(q, wds)
                           Next
                    End If
                Next
                '========================
                If k = 1 Then movecpren = 1 Else movecpren = 3
                If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) And werucount = True Then
                    turnm = 2
                    movecpren = movecpre
                Else
                    turnm = 1
                End If
                '========================
                ���z��AI�t����.���z��AI�t�έp��_�@���q_���o�P����� False, uscom
                ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turnm, movecpren, uscom
                ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turnm, movecpren, uscom
                ���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_1_��l
'                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_3_��ܲզX
        Else
                cardAInumselect1 = 0
        End If
        '=======================�N���s���p�����x�s
        If k = 1 And werucount = False Then
           movecpren = 3
        ElseIf k = 2 And werucount = False Then
           movecpren = 4
        ElseIf werucount = True Then
           movecpren = 8
        End If
        '=========
        cardAInumMOVnmtot(i, movecpren) = cardAInumselect1
        '=========
        If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) And k = 2 And werucount = False Then
           werucount = True
           k = 0
        ElseIf werucount = True Then
           k = 2
        End If
        '==========================
    Next
Next
'=========================���եζץX
'If Formsetting.checktest.Value = 1 Then
''    Open App.Path & "\test\out3.txt" For Output As #1
'    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & BattleTurn & "turn_" & �԰��t����.turnatk & "_" & uscom & "_3.txt" For Output As #1
'    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
'        Print #1, i & "=" & cardAInumMOVnmtot(i, 2) & "=";
'        For k = 3 To 4
'              Print #1, cardAInumMOVnmtot(i, k) & "#";
'        Next
'        If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) Then
'            Print #1, cardAInumMOVnmtot(i, 8);
'        End If
'        Print #1,
'    Next
'
'    Close
'    'MsgBox "�w�ץX����3"
'End If
'==============================
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�|���q_�έp���p����ȤΧP�_(ByVal uscom As Integer)
Dim atk1max As Integer, atk2max As Integer, defmax As Integer, chemax As Integer, chestr As String
Dim wtmovnum As Integer
Dim buffobj As clsStatus
'==================�z��O�_�ŦX���ʶq
For Each n In �H�����`���A�C��(uscom, ����H����ԤH��(uscom, 2))
    Set buffobj = n
    If buffobj.Identifier = "BUFFN00302" Then
        wtmovnum = Val(wtmovnum) - buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00301" Then
        wtmovnum = Val(wtmovnum) + buffobj.Value
    ElseIf buffobj.Identifier = "BUFFN00801" Then
        wtmovnum = -100
    End If
Next
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, 6)) + Val(wtmovnum) < 2 Then
         cardAInumMOVnmtot(i, 5) = "x"
     Else
         cardAInumMOVnmtot(i, 5) = "y"
     End If
Next
'===================
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, 3)) > Val(atk1max) And cardAInumMOVnmtot(i, 5) = "y" Then
         atk1max = cardAInumMOVnmtot(i, 3)
     End If
     If Val(cardAInumMOVnmtot(i, 4)) > Val(atk2max) And cardAInumMOVnmtot(i, 5) = "y" Then
         atk2max = cardAInumMOVnmtot(i, 4)
     End If
Next
defmax = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 8)
'==================
If Val(atk1max) >= Val(atk2max) And Val(atk1max) >= Val(defmax) Then
    chemax = 1
ElseIf Val(atk1max) <= Val(atk2max) And Val(atk2max) >= Val(defmax) Then
    chemax = 2
ElseIf Val(defmax) >= Val(atk1max) And Val(defmax) >= Val(atk2max) Then
    chemax = 3
Else
    chemax = 3
End If
'==================
Select Case chemax
     Case 1
           cardAInumMOVFinal(3) = 1
           cardAInumMOVFinal(2) = atk1max
           ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�T�{���_��̲ܳײզX 1, atk1max
     Case 2
           cardAInumMOVFinal(3) = 2
           cardAInumMOVFinal(2) = atk2max
           ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�T�{���_��̲ܳײզX 2, atk2max
     Case 3
           cardAInumMOVFinal(1) = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 2)
           cardAInumMOVFinal(3) = 3
           cardAInumMOVFinal(2) = defmax
End Select
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�T�{���_��̲ܳײզX(ByVal movche As Integer, ByVal atkmax As Integer)
Dim werstr As String, werg() As String, werg2() As String, werg3() As String
Dim werpagenum As Integer, werpgnumstr As String
Dim wermovmaxnum As Integer, wermvaxstr As String
Dim werrndnum As Integer, werche As Integer
'==========================
If movche = 1 Then werche = 3 Else werche = 4
'==========================
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, werche)) = Val(atkmax) Then
         werstr = werstr & "=" & i
     End If
Next
werg = Split(werstr, "=")
If UBound(werg) > 1 Then
        '====================================
        werpagenum = 0 '==�ت����̤j���X�P��
        For k = 1 To UBound(werg)
            If cardAInumMOVnmtot(werg(k), 7) > werpagenum Then
                werpagenum = cardAInumMOVnmtot(werg(k), 7)
            End If
        Next
        For k = 1 To UBound(werg)
            If cardAInumMOVnmtot(werg(k), 7) = werpagenum Then
                werpgnumstr = werpgnumstr & "=" & werg(k)
            End If
        Next
        werg2 = Split(werpgnumstr, "=")
        If UBound(werg2) > 1 Then
                '====================================
                wermovmaxnum = 0 '==�ت����̤j�����ʼ�
                For k = 1 To UBound(werg2)
                    If Val(cardAInumMOVnmtot(werg(k), 6)) > Val(wermovmaxnum) Then
                        wermovmaxnum = cardAInumMOVnmtot(werg(k), 6)
                    End If
                Next
                For k = 1 To UBound(werg2)
                    If Val(cardAInumMOVnmtot(werg(k), 6)) = wermovmaxnum Then
                        wermvaxstr = wermvaxstr & "=" & werg2(k)
                    End If
                Next
                werg3 = Split(wermvaxstr, "=")
                If UBound(werg3) > 1 Then
                     Randomize
                     werrndnum = Int(Rnd() * UBound(werg3)) + 1
                     cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg3(werrndnum), 2)
                Else
                     cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg3(1), 2)
                End If
                '==========================================
        Else
                cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg2(1), 2)
        End If
        '====================================
Else
        cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg(1), 2)
End If
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�����q_����P(ByVal choose As Integer, ByVal uscom As Integer, ByVal pagenumber As Integer)
'Dim wer As Integer '�Ȯ��ܼ�
'If choose = 1 Then
'    wer = 0
'Else
'    wer = 1
'End If
'=================
Dim pu As Integer '�Ȯ��ܼ�
'=======================�p�զX�ŦX�X�P���󪺸�
Select Case uscom
     Case 1 '==�ϥΪ̤�
            For i = 1 To pagenumber
                    pu = cardcountAInumMOV(i, 5)
                    If Mid(cardAInumMOVFinal(1), i, 1) = 1 Then
'                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 And Val(cardAInumMOVmain(2, i)) >= wer Then
                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 Then
                                pagecardnum(pu, 11) = 4
'                            ElseIf Val(cardAInumMOVmain(2, i)) >= wer Then
                            Else
                                pagecardnum(pu, 11) = 3
                            End If
                    End If
            Next
            '===================��ܦ��
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      �ثe��(33) = 3
                 Case 2
                      �ثe��(33) = 1
                 Case 3
                      �ثe��(33) = 2
            End Select
     Case 2 '==�q����
            For i = 1 To pagenumber
                    pu = cardcountAInumMOV(i, 5)
                    If Mid(cardAInumMOVFinal(1), i, 1) = 1 Then
'                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 And Val(cardAInumMOVmain(2, i)) >= wer Then
                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 Then
                                cspce = pagecardnum(pu, 1)
                                cspme = pagecardnum(pu, 2)
                                pagecardnum(pu, 1) = pagecardnum(pu, 3)
                                pagecardnum(pu, 2) = pagecardnum(pu, 4)
                                pagecardnum(pu, 3) = cspce
                                pagecardnum(pu, 4) = cspme
                                If pageonin(pu) = 2 Then
                                   pageonin(pu) = 1
                                Else
                                   pageonin(pu) = 2
                                End If
                            End If
                            '==================
                            pagecardnum(pu, 11) = 1
                    End If
            Next
            '===================��ܦ��
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      �q���貾�ʶ��q��ܼ� = 3
                 Case 2
                      �q���貾�ʶ��q��ܼ� = 1
                 Case 3
                      �q���貾�ʶ��q��ܼ� = 2
            End Select
End Select

�O�_���ʶ��q����p�P�_�{�� = False
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_�W�X�P�i��(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer, ByVal CardNumMax As Integer)
If Val(pageglead(uscom)) > CardNumMax Then
    Dim CardOverCountNUM As Integer, CardNowNUM1 As Integer, CardNowNUM2 As Integer, CardNowCountNUM As Integer
    Dim w As Integer, k As Integer '�Ȯ��ܼ�
    CardOverCountNUM = Int(Val(pageglead(uscom)) / Val(CardNumMax) + Val(0.9))
    CardNowNUM1 = 1: CardNowNUM2 = CardNumMax
    CardNowCountNUM = 0
    '==========================
    Do
        ReDim cardAInumOvertenrecord(1 To (CardNowNUM2 - CardNowNUM1 + 1)) As Integer
        ���z��AI�t����.���z��AI�t�έp��_�@���q_��l (CardNowNUM2 - CardNowNUM1 + 1)
        '=========�^���ثe�P�����(�e[CardNumMax]�i)
            Select Case uscom
                Case 1
                    �԰��t����.�X�P���ǭp��_�ϥΪ�_��P
                Case 2
                    �԰��t����.�X�P���ǭp��_�q��_��P
            End Select
            w = 2 * uscom '(2-�ϥΪ̤�P/4-�q����P)
            k = 1
            For i = CardNowNUM1 To CardNowNUM2
                cardcountAInum(k, 5) = �X�P���ǲέp�Ȯ��ܼ�(w, i, 2)
                cardAInumOvertenrecord(k) = �X�P���ǲέp�Ȯ��ܼ�(w, i, 2)
                cardcountAInum(k, 1) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 1)
                cardcountAInum(k, 2) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 2)
                cardcountAInum(k, 3) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 3)
                cardcountAInum(k, 4) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 4)
                k = k + 1
            Next
         '========================
        ���z��AI�t����.���z��AI�t�έp��_�@���q_���o�P����� False, uscom
        ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turn, movecpre, uscom
        ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turn, movecpre, uscom
        ���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
        ���z��AI�t����.���z��AI�t�έp��_�|���q_���_1_��l
    '    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1
        ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2
        ���z��AI�t����.���z��AI�t�έp��_�ȮɶץX uscom
        ���z��AI�t����.���z��AI�t�έp��_�|���q_���_3_��ܲզX
        If turn = 3 And cardAInumchoose > 0 Then
            ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_���ʶ��q�� uscom, turn, name, movecpre, choose, (CardNowNUM2 - CardNowNUM1 + 1)
            Exit Do
        Else
            ���z��AI�t����.���z��AI�t�έp��_�̫ᶥ�q_����P choose, uscom
        End If
        '==============================
        CardNowNUM1 = CardNowNUM1 + CardNumMax
        CardNowNUM2 = CardNowNUM2 + CardNumMax
        If CardNowNUM2 > Val(pageglead(uscom)) Then CardNowNUM2 = Val(pageglead(uscom))
        CardNowCountNUM = CardNowCountNUM + 1
    Loop Until CardNowCountNUM >= CardOverCountNUM
    '==========================
'    If turn <> 3 Then
'        �԰��t����.comatk_���z��AI�޾ɵ{��_�W�X�P�i�� turn, movecpre, choose
'    End If
End If
End Sub
Sub �ˬd�H���ޯ�O�_��EX��(ByVal uscom As Integer, ByVal name As String)
'Erase personatkingtfr
'For i = 1 To 3
'     If VBEPerson(uscom, i, 1, 1, 1) = name Then
'         For k = 1 To 4
'               If Mid(VBEPerson(uscom, i, 3, k, 1), 1, 2) = "Ex" Then
'                   personatkingtfr(k) = 1
'               Else
'                   personatkingtfr(k) = 0
'               End If
'          Next
'          For k = 1 To 14
''                If (�H�����`���A��Ʈw(uscom, i, k, 3) = 22 And uscom = 1) Or _
''                    (�H�����`���A��Ʈw(uscom, i, k, 3) = 23 And uscom = 2) Then
'                 If �H�����`���A��Ʈw(uscom, i, k, 3) = "BUFFN00701" Then
'                    personatkingtfr(5) = 1
'                End If
'          Next
'     End If
'Next
End Sub
Sub ���z��AI�t��_�ϥΪ̥X�P���q�P�_����()
For i = 1 To ���εP����d�����j������(1)
    If Val(pagecardnum(i, 11)) = 4 And Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
        If pageonin(i) = 1 Then
           pageonin(i) = 2
        Else
           pageonin(i) = 1
        End If
        FormMainMode.card(i).CardRotationType = pageonin(i)
        FormMainMode.card_CardButtonClickin (i)
        pagecardnum(i, 11) = 3
    End If
Next
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_�_�w��_�G���q_��ܦ��(ByVal uscom As Integer)
Select Case uscom
    Case 1
        �ثe��(33) = 2
    Case 2
         �q���貾�ʶ��q��ܼ� = 2
End Select
End Sub
Sub ���z��AI�t��_���涥�q99_�D�ʧޯ����(ByVal uscom As Integer, ByVal turn As Integer, ByVal movecpre As Integer)
'=======================
���涥�q�t��_�ŧi�}�l�ε��� 1
'=======================
Dim VBEStageNumMain(1 To 1) As Integer
ReDim Vss_EventActiveAIScoreNum(1 To 1) As Integer
'=======================
For i = 1 To cardAITotalNUM
    For atkingnum = 1 To 4
        If Vss_PersonAtkingOffNum(uscom, ����H����ԤH��(uscom, 2), atkingnum) = 0 And Val(VBEPerson(uscom, ����H����ԤH��(uscom, 2), 3, atkingnum, 8)) = turn Then
            If ���涥�q�t����.���涥�q�t��_����(atkingnum, 99, VBEPerson(uscom, ����H����ԤH��(uscom, 2), 3, atkingnum, 11), uscom, ����H����ԤH��(uscom, 2)) = True Then
                   ���z��AI�t����.���z��AI�t��_���涥�q�ǳ��ܼƲΦX��� uscom, VBEStageNumMain, turn, movecpre, i
                   ���z��AI�t����.���z��AI�t��_���涥�q99_�p��ӧO������˭Ȳέp uscom, atkingnum, i, turn, ����H����ԤH��(uscom, 2)
            End If
        End If
    Next
Next
'=======================
���涥�q�t��_�ŧi�}�l�ε��� 2
'=======================
End Sub
Sub ���z��AI�t��_���涥�q�ǳ��ܼƲΦX���(ByVal uscom As Integer, ByRef VBEStageNumMain() As Integer, ByVal turnai As Integer, ByVal movecpre As Integer, ByVal cardAICaseNum As Integer)
    '===========================
    Erase VBEPersonVS 'VBE�H���Τ@�ܼ�-VS��
    Erase atkingpagetotVS '�C���q�X�P�����μƭȲέp���-VS��
    Erase VBEPersonBuffVSF  '���`���A���-VS-F��
    Erase VBEPersonBuffVSS  '���`���A���-VS-S��
    Erase AtkingckVSS '�ޯ��T�@��-S��(�ޯ�ҰʽX)
    Erase AtkingckVSF '�ޯ��T�@��-F��(�ޯ�Ƶ��r��)
    Erase VBEAtkingVSF 'VBE>VS�����ܼƲΤ@���-F��
    Erase VBEAtkingVSS 'VBE>VS�����ܼƲΤ@���-S��
'    Erase VBEPageCardNumVS '���εP���-VS��
    ReDim VBEPageCardNumVS(1 To cardAInumuscom, 1 To 6) As Variant '���εP���-VS��
'    Erase VBEVSStageNum '���涥�q�t��-���涥�q�h�γ~�����ܼ�-VS��
    ReDim VBEVSStageNum(1 To UBound(VBEStageNumMain)) As Variant '���涥�q�t��-���涥�q�h�γ~�����ܼ�-VS��
    Erase VBEActualStatusVS '�H����ڪ��A���-VS��
    '===========================
    Dim q As Integer, w As Integer, rr As Integer, cs1 As Variant, cs2 As Variant, tempc As Integer, buffobj As clsStatus
    tempc = 1
    For i = 1 To 2
        For j = 1 To 3
            If �H�����`���A�C��(i, j).Count > tempc Then
                tempc = �H�����`���A�C��(i, j).Count
            End If
        Next
    Next
    ReDim VBEPersonBuffVSF(1 To 2, 1 To 3, 1 To tempc, 1 To 2) As Variant '���`���A���-VS-F��
    ReDim VBEPersonBuffVSS(1 To 2, 1 To 3, 1 To tempc) As Variant '���`���A���-VS-S��
    '===========================
    Select Case uscom
         Case 1
             '(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11)
             For i = 1 To 2
                 For j = 1 To 3
                     For k = 1 To 4
                         For m = 1 To 30
                             For p = 1 To 11
                                 VBEPersonVS(i, j, k, m, p) = VBEPerson(i, ����ݾ��H��������(i, j), k, m, p)
                             Next
                         Next
                      Next
                 Next
            Next
            '======================
            For i = 1 To cardAInumuscom
                For j = 1 To 6
                    If j = 1 Or j = 3 Then
                       Select Case cardcountAInum(i, j)
                           Case "ATK-�C"
                               VBEPageCardNumVS(i, j) = 1
                           Case "DEF"
                               VBEPageCardNumVS(i, j) = 2
                           Case "MOV"
                               VBEPageCardNumVS(i, j) = 3
                           Case "SPE"
                               VBEPageCardNumVS(i, j) = 4
                           Case "ATK-�j"
                               VBEPageCardNumVS(i, j) = 5
                           Case "DRAW"
                               VBEPageCardNumVS(i, j) = 6
                           Case "BRK"
                               VBEPageCardNumVS(i, j) = 7
                           Case "HPL"
                               VBEPageCardNumVS(i, j) = 8
                           Case Else
                               VBEPageCardNumVS(i, j) = 0
                       End Select
                    ElseIf j >= 5 Then
                       VBEPageCardNumVS(i, j) = 1
                    Else
                        VBEPageCardNumVS(i, j) = Val(cardcountAInum(i, j))
                    End If
                Next
                '==================
                If Mid(cardAInumnm(cardAICaseNum - 1), i, 1) = 1 Then
                    cs1 = VBEPageCardNumVS(i, 1)
                    cs2 = VBEPageCardNumVS(i, 2)
                    VBEPageCardNumVS(i, 1) = VBEPageCardNumVS(i, 3)
                    VBEPageCardNumVS(i, 2) = VBEPageCardNumVS(i, 4)
                    VBEPageCardNumVS(i, 3) = cs1
                    VBEPageCardNumVS(i, 4) = cs2
                End If
                '==================
            Next
            '======================
            '(1 To 2, 1 To 5)
            For j = 1 To 5
                atkingpagetotVS(1, j) = cardAInumcaseperson(cardAICaseNum, 1, 10 + j)
            Next
            For j = 1 To 5
                atkingpagetotVS(2, j) = atkingpagetot(2, j)
            Next
            '======================
            '(1 To 2, 1 To 3, 1 To 14, 1 To 3)
            For i = 1 To 2
                For rr = 1 To 3
                    For j = 1 To �H�����`���A�C��(i, ����ݾ��H��������(i, rr)).Count
                        Set buffobj = �H�����`���A�C��(i, ����ݾ��H��������(i, rr))(j)
                        VBEPersonBuffVSF(i, rr, j, 1) = buffobj.Value
                        VBEPersonBuffVSF(i, rr, j, 2) = buffobj.Total
                        VBEPersonBuffVSS(i, rr, j) = buffobj.Identifier
                    Next
                Next
            Next
            '======================
            '(1 to 2,1 to 3,1 to 2)
            For i = 1 To 2
                For rr = 1 To 3
                    VBEActualStatusVS(i, rr, 1) = �H����ڪ��A��Ʈw(i, ����ݾ��H��������(i, rr), 1)
                    VBEActualStatusVS(i, rr, 2) = �H����ڪ��A��Ʈw(i, ����ݾ��H��������(i, rr), 9)
                Next
            Next
            '======================
            '(1 to 8,1 to 3)
            For i = 1 To 8
                For j = 1 To 3
                    AtkingckVSS(i, j) = atkingck(uscom, ����H����ԤH��(uscom, 2), i, j)
                Next
                AtkingckVSF(i, 1) = Vss_AtkingInformationRecordStr(uscom, ����H����ԤH��(uscom, 2), i)
            Next
            '======================
            For i = 1 To 3
                VBEAtkingVSF(1, i, 1) = liveus(����ݾ��H��������(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(1, i, 2) = liveusmax(����ݾ��H��������(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(2, i, 1) = livecom(����ݾ��H��������(2, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(2, i, 2) = livecommax(����ݾ��H��������(2, i))
            Next
            '========================
            VBEAtkingVSS(0) = 1
            VBEAtkingVSS(1) = cardAInumuscom
            VBEAtkingVSS(2) = 0
            VBEAtkingVSS(3) = pageqlead(2)
            VBEAtkingVSS(4) = pageglead(2)
            VBEAtkingVSS(6) = movecpre
            If �O�_���ʶ��q����p�P�_�{�� = False Then
                VBEAtkingVSS(5) = �Y���淾�q�Ȯ��ܼ�(2)
                VBEAtkingVSS(7) = Val(�������m��l�`��(1))
                VBEAtkingVSS(8) = Val(�������m��l�`��(2))
                VBEAtkingVSS(14) = �Y���淾�q�Ȯ��ܼ�(5)
                VBEAtkingVSS(15) = �Y���淾�q�Ȯ��ܼ�(6)
                VBEAtkingVSS(16) = moveturn
            Else
                VBEAtkingVSS(5) = 0
                VBEAtkingVSS(7) = 0
                VBEAtkingVSS(8) = 0
                VBEAtkingVSS(14) = 0
                VBEAtkingVSS(15) = 0
                VBEAtkingVSS(16) = 1
            End If
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            Select Case turnai
                Case 1
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 2
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
             End Select
             VBEAtkingVSS(17) = ����H����ԤH��(1, 1)
             VBEAtkingVSS(18) = ����H����ԤH��(2, 1)
             VBEAtkingVSS(19) = �P�`���q��(1)
             VBEAtkingVSS(20) = �P�`���q��(2)
             '=========================
             For i = 1 To UBound(VBEStageNumMain)
                 If VBEStageNumMain(i) = -1 Or VBEStageNumMain(i) = -2 Then
                     VBEVSStageNum(i) = Abs(VBEStageNumMain(i))
                 Else
                     VBEVSStageNum(i) = VBEStageNumMain(i)
                 End If
             Next
         Case 2 '===============================================================
             '(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11)
             For i = 1 To 2
                 If i = 1 Then q = 2 Else q = 1
                 For j = 1 To 3
                     For k = 1 To 4
                         For m = 1 To 30
                             For p = 1 To 11
                                 VBEPersonVS(i, j, k, m, p) = VBEPerson(q, ����ݾ��H��������(q, j), k, m, p)
                             Next
                         Next
                      Next
                 Next
            Next
            '======================
            For i = 1 To cardAInumuscom
                For j = 1 To 6
                     If j = 1 Or j = 3 Then
                       Select Case cardcountAInum(i, j)
                           Case "ATK-�C"
                               VBEPageCardNumVS(i, j) = 1
                           Case "DEF"
                               VBEPageCardNumVS(i, j) = 2
                           Case "MOV"
                               VBEPageCardNumVS(i, j) = 3
                           Case "SPE"
                               VBEPageCardNumVS(i, j) = 4
                           Case "ATK-�j"
                               VBEPageCardNumVS(i, j) = 5
                           Case "DRAW"
                               VBEPageCardNumVS(i, j) = 6
                           Case "BRK"
                               VBEPageCardNumVS(i, j) = 7
                           Case "HPL"
                               VBEPageCardNumVS(i, j) = 8
                           Case Else
                               VBEPageCardNumVS(i, j) = 0
                       End Select
                    ElseIf j >= 5 Then
                        VBEPageCardNumVS(i, j) = 1
                    Else
                       VBEPageCardNumVS(i, j) = Val(cardcountAInum(i, j))
                    End If
                Next
                '==================
                If Mid(cardAInumnm(cardAICaseNum - 1), i, 1) = 1 Then
                    cs1 = VBEPageCardNumVS(i, 1)
                    cs2 = VBEPageCardNumVS(i, 2)
                    VBEPageCardNumVS(i, 1) = VBEPageCardNumVS(i, 3)
                    VBEPageCardNumVS(i, 2) = VBEPageCardNumVS(i, 4)
                    VBEPageCardNumVS(i, 3) = cs1
                    VBEPageCardNumVS(i, 4) = cs2
                End If
                '==================
            Next
            '======================
            '(1 To 2, 1 To 5)
            For j = 1 To 5
                atkingpagetotVS(1, j) = cardAInumcaseperson(cardAICaseNum, 1, 10 + j)
            Next
            For j = 1 To 5
                atkingpagetotVS(2, j) = atkingpagetot(1, j)
            Next
            '======================
            '(1 To 2, 1 To 3, 1 To 14, 1 To 3)
            For i = 1 To 2
                If i = 1 Then q = 2 Else q = 1
                For rr = 1 To 3
                    For j = 1 To �H�����`���A�C��(q, ����ݾ��H��������(q, rr)).Count
                        Set buffobj = �H�����`���A�C��(q, ����ݾ��H��������(q, rr))(j)
                        VBEPersonBuffVSF(i, rr, j, 1) = buffobj.Value
                        VBEPersonBuffVSF(i, rr, j, 2) = buffobj.Total
                        VBEPersonBuffVSS(i, rr, j) = buffobj.Identifier
                    Next
                Next
            Next
            '======================
            '(1 to 2,1 to 3,1 to 2)
            For i = 1 To 2
                If i = 1 Then w = 2 Else w = 1
                For rr = 1 To 3
                    VBEActualStatusVS(i, rr, 1) = �H����ڪ��A��Ʈw(w, ����ݾ��H��������(w, rr), 1)
                    VBEActualStatusVS(i, rr, 2) = �H����ڪ��A��Ʈw(w, ����ݾ��H��������(w, rr), 9)
                Next
            Next
            '======================
            '(1 to 8,1 to 3)
            For i = 1 To 8
                For j = 1 To 3
                    AtkingckVSS(i, j) = atkingck(uscom, ����H����ԤH��(uscom, 2), i, j)
                Next
                AtkingckVSF(i, 1) = Vss_AtkingInformationRecordStr(uscom, ����H����ԤH��(uscom, 2), i)
            Next
            '======================
            For i = 1 To 3
                VBEAtkingVSF(2, i, 1) = liveus(����ݾ��H��������(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(2, i, 2) = liveusmax(����ݾ��H��������(1, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(1, i, 1) = livecom(����ݾ��H��������(2, i))
            Next
            For i = 1 To 3
                VBEAtkingVSF(1, i, 2) = livecommax(����ݾ��H��������(2, i))
            Next
            '========================
            VBEAtkingVSS(0) = 1
            VBEAtkingVSS(1) = cardAInumuscom
            VBEAtkingVSS(2) = 0
            VBEAtkingVSS(3) = pageqlead(1)
            VBEAtkingVSS(4) = pageglead(1)
            VBEAtkingVSS(6) = movecpre
            If �O�_���ʶ��q����p�P�_�{�� = False Then
                VBEAtkingVSS(5) = �Y���淾�q�Ȯ��ܼ�(2)
                VBEAtkingVSS(7) = Val(�������m��l�`��(2))
                VBEAtkingVSS(8) = Val(�������m��l�`��(1))
                VBEAtkingVSS(14) = �Y���淾�q�Ȯ��ܼ�(6)
                VBEAtkingVSS(15) = �Y���淾�q�Ȯ��ܼ�(5)
                If moveturn = 2 Then VBEAtkingVSS(16) = 1 Else VBEAtkingVSS(16) = 2
            Else
                VBEAtkingVSS(5) = 0
                VBEAtkingVSS(7) = 0
                VBEAtkingVSS(8) = 0
                VBEAtkingVSS(14) = 0
                VBEAtkingVSS(15) = 0
                VBEAtkingVSS(16) = 1
            End If
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            Select Case turnai
                Case 1
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 2
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
             End Select
             VBEAtkingVSS(17) = ����H����ԤH��(2, 1)
             VBEAtkingVSS(18) = ����H����ԤH��(1, 1)
             VBEAtkingVSS(19) = �P�`���q��(2)
             VBEAtkingVSS(20) = �P�`���q��(1)
             '=========================
             For i = 1 To UBound(VBEStageNumMain)
                 If VBEStageNumMain(i) = -1 Then
                     VBEVSStageNum(i) = 2
                 ElseIf VBEStageNumMain(i) = -2 Then
                     VBEVSStageNum(i) = 1
                 Else
                     VBEVSStageNum(i) = VBEStageNumMain(i)
                 End If
             Next
   End Select
End Sub
Sub ���z��AI�t��_���涥�q99_�p��ӧO������˭Ȳέp(ByVal uscom As Integer, ByVal atkingnum As Integer, ByVal cardAICaseNum As Integer, ByVal turn As Integer, ByVal personnum As Integer)
Dim vsstr As String, vsstr2() As String, vsstr3() As String, vsstr4() As String, vstest As String, uscomt As Integer
'============�^�����涥�q99���������
vsstr = ���涥�q�t����.���涥�q�t��_����}��_�H���D�ʧޯ���(atkingnum, 99, uscom, personnum)
vsstr2 = Split(vsstr, "=")
For i = 0 To UBound(vsstr2)
    If vsstr2(i) <> "" Then
        vsstr3 = Split(vsstr2(i), "#")
        If vsstr3(0) = "EventActiveAIScore" Then
            vsstr4 = Split(vsstr3(1), ",")
            vstest = vsstr2(i)
            '===================================
            ReDim Vss_EventActiveAIScoreNum(1 To UBound(vsstr4) + 1) As Integer
            For k = 0 To UBound(vsstr4)
                Vss_EventActiveAIScoreNum(k + 1) = vsstr4(k)
            Next
            '===================================
            Exit For
        End If
    End If
Next
'==================================================================
If Vss_EventActiveAIScoreNum(1) = 1 Then
    If Vss_EventActiveAIScoreNum(2) = 1 Then
        cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + 10
    ElseIf Vss_EventActiveAIScoreNum(2) = 2 Then
        '============�^�����涥�q45���`����ܤƶq���
        vsstr = ���涥�q�t����.���涥�q�t��_����}��_�H���D�ʧޯ���(atkingnum, 45, uscom, personnum)
        vsstr2 = Split(vsstr, "=")
        For i = 0 To UBound(vsstr2)
            If vsstr2(i) <> "" Then
                vsstr3 = Split(vsstr2(i), "#")
                If vsstr3(0) = "EventTotalDiceChange" Then
                    vsstr4 = Split(vsstr3(1), ",")
                    '===================================
                    If Val(vsstr4(0)) = 1 Then  '���ۨ����ܤƤ��q
                        Select Case Val(vsstr4(1))
                            Case 1
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + Val(vsstr4(2))
                            Case 2
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) - Val(vsstr4(2))
                            Case 3
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) * Val(vsstr4(2))
                            Case Is <= 5
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) \ Val(vsstr4(2))
                            Case 6
                                If turn = 1 Then
                                    cardAInumFinal(cardAICaseNum, 1) = Val(vsstr4(2))
                                ElseIf turn = 2 Then
                                    cardAInumFinal(cardAICaseNum, 1) = Val(vsstr4(2)) - VBEPerson(uscom, ����H����ԤH��(uscom, 2), 1, 3, 3)
                                End If
                        End Select
                    ElseIf Val(vsstr4(0)) = 2 Then  '�������ܤƤ��q
                        Select Case Val(vsstr4(1))
                            Case 1
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) - Val(vsstr4(2))
                            Case 2
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + Val(vsstr4(2))
                            Case 3
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) \ Val(vsstr4(2))
                            Case Is <= 5
                                cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) * Val(vsstr4(2))
                            Case 6
                                If uscom = 1 Then uscomt = 2 Else uscomt = 1
                                If turn = 1 Then
                                    cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + (VBEPerson(uscomt, ����H����ԤH��(uscomt, 2), 1, 3, 3) - Val(vsstr4(2))) * 2 + 5
                                ElseIf turn = 2 Then
                                    cardAInumFinal(cardAICaseNum, 1) = cardAInumFinal(cardAICaseNum, 1) + (VBEPerson(uscomt, ����H����ԤH��(uscomt, 2), 1, 3, 2) - Val(vsstr4(2))) * 2 + 5
                                End If
                        End Select
                    End If
                    '===================================
'                    Exit For
                End If
            End If
        Next
    End If
    '=====================================
    If Vss_EventActiveAIScoreNum(2) = 1 Or Vss_EventActiveAIScoreNum(2) = 2 Then
        For i = 3 To UBound(Vss_EventActiveAIScoreNum)
            If Vss_EventActiveAIScoreNum(i) > 0 And Vss_EventActiveAIScoreNum(i) <= cardAInumuscom Then
                cardAInumcaseperson(cardAICaseNum, 2, Vss_EventActiveAIScoreNum(i)) = 1
'                 MsgBox vstest & Chr(10) & cardcountAInum(Vss_EventActiveAIScoreNum(i), 1) & "," & cardcountAInum(Vss_EventActiveAIScoreNum(i), 2) & ",  " & cardcountAInum(Vss_EventActiveAIScoreNum(i), 3) & "," & cardcountAInum(Vss_EventActiveAIScoreNum(i), 4) & Chr(10) & "uscom:" & uscom & "  ,atkingnum:" & atkingnum
            End If
        Next
    End If
    '=====================================
ElseIf Vss_EventActiveAIScoreNum(1) = 3 Then
    cardAInumFinal(cardAICaseNum, 1) = -100
End If
'=================
ReDim Vss_EventActiveAIScoreNum(1 To 1) As Integer
End Sub
