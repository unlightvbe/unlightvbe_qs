Attribute VB_Name = "�ƥ�d"
Option Explicit
Public �ƥ�d�O���Ȯɼ�(0 To 2, 1 To 6) As Integer '�ƥ�d�ϥά����Ȯ��ܼ�(0.(1)�`�@�����^�X��,1.�ϥΪ�/2.�q��,1.�`�@�ƭ�/2.�ثe�B�z�ƭ�/3.�ثe���q/4.�ƥ�d�P�s��/5.�ƥ����/6.�O�_�Ұ�)
Sub ���|_�ϥΪ�(ByVal num As Integer, ByVal tot As Integer)
Select Case �ƥ�d�O���Ȯɼ�(1, 3)
    Case 1
        �ثe��(15) = 7
        �ƥ�d�O���Ȯɼ�(1, 4) = num
        �ƥ�d�O���Ȯɼ�(1, 1) = tot
        �ƥ�d�O���Ȯɼ�(1, 5) = 1
        �ƥ�d�O���Ȯɼ�(1, 6) = 1
        FormMainMode.��������ˬd.Enabled = False
    Case 2
        �@��t����.���ļ��� 7
        '=============�H�U�O�P����(���P)(�ϥΪ�)
'         �԰��t����.�y�Эp��_�ϥΪ̤�P
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(1, 4)
         �ثe��(5) = pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 7)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 6
         FormMainMode.�P����.Enabled = True
        '================�H�U�O�X�P���
        �ثe��(3) = 0
        �԰��t����.�X�P���ǭp��_�ϥΪ�_�X�P
        FormMainMode.�ϥΪ̥X�P_�X�P���_�a�k.Enabled = True
        '=====================
        �ƥ�d�O���Ȯɼ�(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        If BattleCardNum < �ƥ�d�O���Ȯɼ�(1, 1) Then
           �԰��t����.����ʧ@_�~�P
        End If
    Case 3
         If �ƥ�d�O���Ȯɼ�(1, 2) > �ƥ�d�O���Ȯɼ�(1, 1) Or BattleCardNum <= 0 Then
             turnpageonin = 1
             FormMainMode.PEAFInterface.BnOKStartListen
             �ƥ�d�O���Ȯɼ�(1, 6) = 0
             Exit Sub
         End If
         Do Until �ƥ�d�O���Ȯɼ�(1, 2) > �ƥ�d�O���Ȯɼ�(1, 1)
             �ثe��(15) = 8
             FormMainMode.tr�P��_��P_�ϥΪ�.Enabled = True
             �ƥ�d�O���Ȯɼ�(1, 2) = �ƥ�d�O���Ȯɼ�(1, 2) + 1
             Exit Do
         Loop
End Select
End Sub
Sub ���|_�q��(ByVal num As Integer, ByVal tot As Integer)
Select Case �ƥ�d�O���Ȯɼ�(2, 3)
    Case 1
        �ثe��(15) = 9
        �ثe��(17) = 2
        �ƥ�d�O���Ȯɼ�(2, 4) = num
        �ƥ�d�O���Ȯɼ�(2, 1) = tot
        �ƥ�d�O���Ȯɼ�(2, 5) = 1
        �ƥ�d�O���Ȯɼ�(2, 6) = 1
    Case 2
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Width = 810
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Height = 1260
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).cardImage = app_path & "card\" & pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 8) & ".png"
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).CardRotationType = pageonin(�ƥ�d�O���Ȯɼ�(2, 4))
        �@��t����.���ļ��� 7
        ���ݮɶ���C(2).Add 9
        FormMainMode.���ݮɶ�_2.Enabled = True
    Case 3
        '=============�H�U�O�P����(���P)(�q��)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(2, 4)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 10
         FormMainMode.�P����.Enabled = True
        '=====================
        �ƥ�d�O���Ȯɼ�(2, 2) = 1
        If BattleCardNum < �ƥ�d�O���Ȯɼ�(2, 1) Then
           �԰��t����.����ʧ@_�~�P
        End If
    Case 4
         If �ƥ�d�O���Ȯɼ�(2, 2) > �ƥ�d�O���Ȯɼ�(2, 1) Or BattleCardNum <= 0 Then
             ���ݮɶ���C(2).Add 10
             FormMainMode.���ݮɶ�_2.Enabled = True
             �ƥ�d�O���Ȯɼ�(2, 6) = 0
             Exit Sub
         End If
         Do Until �ƥ�d�O���Ȯɼ�(2, 2) > �ƥ�d�O���Ȯɼ�(2, 1)
             �ثe��(15) = 11
             FormMainMode.tr�P��_��P_�q��.Enabled = True
             �ƥ�d�O���Ȯɼ�(2, 2) = �ƥ�d�O���Ȯɼ�(2, 2) + 1
             Exit Do
         Loop
End Select
End Sub
Sub �A�G�N_�ϥΪ�(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case �ƥ�d�O���Ȯɼ�(1, 3)
    Case 1
        �ثe��(15) = 12
        �ƥ�d�O���Ȯɼ�(1, 4) = num
        �ƥ�d�O���Ȯɼ�(1, 1) = tot
        �ƥ�d�O���Ȯɼ�(1, 5) = 2
        �ƥ�d�O���Ȯɼ�(1, 6) = 1
        FormMainMode.��������ˬd.Enabled = False
    Case 2
        �ƥ�d�O���Ȯɼ�(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        '=======================
        Do Until �ƥ�d�O���Ȯɼ�(1, 2) > �ƥ�d�O���Ȯɼ�(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * ���εP����d�����j������(1)) + 1
            If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                �ثe��(17) = 6
                �ثe��(16) = m
                �ƥ�d�O���Ȯɼ�(1, 2) = �ƥ�d�O���Ȯɼ�(1, 2) + 1
                FormMainMode.tr�q���P_½�P.Enabled = True
                Exit Sub
            End If
        Loop
        If �ƥ�d�O���Ȯɼ�(1, 2) > �ƥ�d�O���Ȯɼ�(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0 Then
            ���ݮɶ���C(2).Add 12
            FormMainMode.���ݮɶ�_2.Enabled = True
        End If
     Case 3
        Do Until �ƥ�d�O���Ȯɼ�(1, 2) > �ƥ�d�O���Ȯɼ�(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * ���εP����d�����j������(1)) + 1
            If Val(pagecardnum(m, 5)) = 2 And Val(pagecardnum(m, 6)) = 1 Then
                �ثe��(17) = 6
                �ثe��(16) = m
                �ƥ�d�O���Ȯɼ�(1, 2) = �ƥ�d�O���Ȯɼ�(1, 2) + 1
                FormMainMode.tr�q���P_½�P.Enabled = True
                Exit Sub
            End If
        Loop
        If �ƥ�d�O���Ȯɼ�(1, 2) > �ƥ�d�O���Ȯɼ�(1, 1) Or Val(FormMainMode.pagecomglead.Caption) <= 0 Then
            ���ݮɶ���C(2).Add 12
            FormMainMode.���ݮɶ�_2.Enabled = True
        End If
     Case 4
        FormMainMode.tr�q���P_��P.Enabled = True
     Case 5
         �@��t����.���ļ��� 7
        '=============�H�U�O�P����(���P)(�ϥΪ�)
'         �԰��t����.�y�Эp��_�ϥΪ̤�P
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(1, 4)
         �ثe��(5) = pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 7)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 13
         FormMainMode.�P����.Enabled = True
        '================�H�U�O�X�P���
        �ثe��(3) = 0
        �԰��t����.�X�P���ǭp��_�ϥΪ�_�X�P
        FormMainMode.�ϥΪ̥X�P_�X�P���_�a�k.Enabled = True
        '=====================
        �ƥ�d�O���Ȯɼ�(1, 2) = 1
    Case 6
        turnpageonin = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        �ƥ�d�O���Ȯɼ�(1, 6) = 0
End Select
End Sub
Sub �A�G�N_�q��(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case �ƥ�d�O���Ȯɼ�(2, 3)
    Case 1
        �ثe��(15) = 14
        �ثe��(17) = 2
        �ƥ�d�O���Ȯɼ�(2, 4) = num
        �ƥ�d�O���Ȯɼ�(2, 1) = tot
        �ƥ�d�O���Ȯɼ�(2, 5) = 2
        �ƥ�d�O���Ȯɼ�(2, 6) = 1
    Case 2
        �ƥ�d�O���Ȯɼ�(2, 2) = 1
        '=======================
        Do Until �ƥ�d�O���Ȯɼ�(2, 2) > �ƥ�d�O���Ȯɼ�(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * ���εP����d�����j������(1)) + 1
            If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                �ثe��(21) = 3
                �ثe��(20) = m
                �ƥ�d�O���Ȯɼ�(2, 2) = �ƥ�d�O���Ȯɼ�(2, 2) + 1
                FormMainMode.tr�ϥΪ�_��P.Enabled = True
                Exit Sub
            End If
        Loop
        If �ƥ�d�O���Ȯɼ�(2, 2) > �ƥ�d�O���Ȯɼ�(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
            ���ݮɶ���C(2).Add 14
            FormMainMode.���ݮɶ�_2.Enabled = True
        End If
     Case 3
        Do Until �ƥ�d�O���Ȯɼ�(2, 2) > �ƥ�d�O���Ȯɼ�(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0
            Randomize
            m = Int(Rnd() * ���εP����d�����j������(1)) + 1
            If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                �ثe��(21) = 3
                �ثe��(20) = m
                �ƥ�d�O���Ȯɼ�(2, 2) = �ƥ�d�O���Ȯɼ�(2, 2) + 1
                FormMainMode.tr�ϥΪ�_��P.Enabled = True
                Exit Sub
            End If
        Loop
        If �ƥ�d�O���Ȯɼ�(2, 2) > �ƥ�d�O���Ȯɼ�(2, 1) Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
            ���ݮɶ���C(2).Add 14
            FormMainMode.���ݮɶ�_2.Enabled = True
        End If
     Case 4
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Width = 810
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Height = 1260
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).cardImage = app_path & "card\" & pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 8) & ".png"
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).CardRotationType = pageonin(�ƥ�d�O���Ȯɼ�(2, 4))
        �@��t����.���ļ��� 7
        ���ݮɶ���C(2).Add 15
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 5
        '=============�H�U�O�P����(���P)(�q��)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(2, 4)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 15
         FormMainMode.�P����.Enabled = True
        '=====================
    Case 6
        ���ݮɶ���C(2).Add 10
        FormMainMode.���ݮɶ�_2.Enabled = True
        �ƥ�d�O���Ȯɼ�(2, 6) = 0
End Select
End Sub
Sub HP�^�__�ϥΪ�(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case �ƥ�d�O���Ȯɼ�(1, 3)
    Case 1
        �ثe��(15) = 16
        �ƥ�d�O���Ȯɼ�(1, 4) = num
        �ƥ�d�O���Ȯɼ�(1, 1) = tot
        �ƥ�d�O���Ȯɼ�(1, 5) = 3
        �ƥ�d�O���Ȯɼ�(1, 6) = 1
        FormMainMode.��������ˬd.Enabled = False
    Case 2
        �ƥ�d�O���Ȯɼ�(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        '=======================
        �԰��t����.�^�_����_�ϥΪ� Val(�ƥ�d�O���Ȯɼ�(1, 1)), 1, 0, True
        ���ݮɶ���C(2).Add 17
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 3
         �@��t����.���ļ��� 7
        '=============�H�U�O�P����(���P)(�ϥΪ�)
'         �԰��t����.�y�Эp��_�ϥΪ̤�P
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(1, 4)
         �ثe��(5) = pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 7)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 17
         FormMainMode.�P����.Enabled = True
        '================�H�U�O�X�P���
        �ثe��(3) = 0
        �԰��t����.�X�P���ǭp��_�ϥΪ�_�X�P
        FormMainMode.�ϥΪ̥X�P_�X�P���_�a�k.Enabled = True
        '=====================
    Case 4
        turnpageonin = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        �ƥ�d�O���Ȯɼ�(1, 6) = 0
End Select
End Sub
Sub HP�^�__�q��(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case �ƥ�d�O���Ȯɼ�(2, 3)
    Case 1
        �ثe��(15) = 18
        �ثe��(17) = 2
        �ƥ�d�O���Ȯɼ�(2, 4) = num
        �ƥ�d�O���Ȯɼ�(2, 1) = tot
        �ƥ�d�O���Ȯɼ�(2, 5) = 3
        �ƥ�d�O���Ȯɼ�(2, 6) = 1
    Case 2
        �ƥ�d�O���Ȯɼ�(2, 2) = 1
        '=======================
        �^�_����_�q�� Val(�ƥ�d�O���Ȯɼ�(2, 1)), 1, 0, True
        ���ݮɶ���C(2).Add 19
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 3
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Width = 810
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Height = 1260
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).cardImage = app_path & "card\" & pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 8) & ".png"
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).CardRotationType = pageonin(�ƥ�d�O���Ȯɼ�(2, 4))
        �@��t����.���ļ��� 7
        ���ݮɶ���C(2).Add 20
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 4
        '=============�H�U�O�P����(���P)(�q��)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(2, 4)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 19
         FormMainMode.�P����.Enabled = True
        '=====================
    Case 5
        ���ݮɶ���C(2).Add 10
        FormMainMode.���ݮɶ�_2.Enabled = True
        �ƥ�d�O���Ȯɼ�(2, 6) = 0
End Select
End Sub
Sub �t��_�ϥΪ�(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case �ƥ�d�O���Ȯɼ�(1, 3)
    Case 1
        �ثe��(15) = 41
        �ƥ�d�O���Ȯɼ�(1, 4) = num
        �ƥ�d�O���Ȯɼ�(1, 1) = tot
        �ƥ�d�O���Ȯɼ�(1, 5) = 4
        �ƥ�d�O���Ȯɼ�(1, 6) = 1
        FormMainMode.��������ˬd.Enabled = False
    Case 2
        �ƥ�d�O���Ȯɼ�(1, 2) = 1
        turnpageonin = 0
        FormMainMode.PEAFInterface.BnOKEnabled False
        '=======================
        �԰��t����.����ʧ@_�M���Ҧ����`���A_�t�� 1, 1
        �԰��t����.��q��s���
        FormMainMode.trgoi1.Enabled = True
        ���ݮɶ���C(2).Add 40
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 3
         �@��t����.���ļ��� 7
        '=============�H�U�O�P����(���P)(�ϥΪ�)
'         �԰��t����.�y�Эp��_�ϥΪ̤�P
         pageqlead(1) = Val(pageqlead(1)) - 1
         FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(1, 4)
         �ثe��(5) = pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 7)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(1, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(1, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 17
         FormMainMode.�P����.Enabled = True
        '================�H�U�O�X�P���
        �ثe��(3) = 0
        �԰��t����.�X�P���ǭp��_�ϥΪ�_�X�P
        FormMainMode.�ϥΪ̥X�P_�X�P���_�a�k.Enabled = True
        '=====================
    Case 4
        turnpageonin = 1
        FormMainMode.PEAFInterface.BnOKStartListen
        �ƥ�d�O���Ȯɼ�(1, 6) = 0
End Select
End Sub
Sub �t��_�q��(ByVal num As Integer, ByVal tot As Integer)
Dim m As Integer
Select Case �ƥ�d�O���Ȯɼ�(2, 3)
    Case 1
        �ثe��(15) = 43
        �ثe��(17) = 2
        �ƥ�d�O���Ȯɼ�(2, 4) = num
        �ƥ�d�O���Ȯɼ�(2, 1) = tot
        �ƥ�d�O���Ȯɼ�(2, 5) = 3
        �ƥ�d�O���Ȯɼ�(2, 6) = 1
    Case 2
        �ƥ�d�O���Ȯɼ�(2, 2) = 1
        '=======================
        �԰��t����.����ʧ@_�M���Ҧ����`���A_�t�� 2, 1
        �԰��t����.��q��s���
        FormMainMode.trgoi2.Enabled = True
        ���ݮɶ���C(2).Add 42
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 3
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Width = 810
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Height = 1260
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).cardImage = app_path & "card\" & pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 8) & ".png"
        FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).CardRotationType = pageonin(�ƥ�d�O���Ȯɼ�(2, 4))
        �@��t����.���ļ��� 7
        ���ݮɶ���C(2).Add 43
        FormMainMode.���ݮɶ�_2.Enabled = True
     Case 4
        '=============�H�U�O�P����(���P)(�q��)
         FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
         pageqlead(2) = Val(pageqlead(2)) - 1
         �P���ʼȮ��ܼ�(1) = 240
         �P���ʼȮ��ܼ�(2) = 960
         �P���ʼȮ��ܼ�(3) = �ƥ�d�O���Ȯɼ�(2, 4)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 9) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Left  '���w�ثeLeft(�y��)
         pagecardnum(�ƥ�d�O���Ȯɼ�(2, 4), 10) = FormMainMode.card(�ƥ�d�O���Ȯɼ�(2, 4)).Top  '���w�ثeTop(�y��)
         �԰��t����.�p��P���ʶZ�����
         �ثe��(15) = 44
         FormMainMode.�P����.Enabled = True
        '=====================
    Case 5
        ���ݮɶ���C(2).Add 10
        FormMainMode.���ݮɶ�_2.Enabled = True
        �ƥ�d�O���Ȯɼ�(2, 6) = 0
End Select
End Sub

