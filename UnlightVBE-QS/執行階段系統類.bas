Attribute VB_Name = "���涥�q�t����"
Option Explicit
Public VBEPersonVS(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 11)  As Variant  'VBE�H���Τ@�ܼ�-VS��
Public atkingpagetotVS(1 To 2, 1 To 5) As Variant  '�C���q�X�P�����μƭȲέp���-VS��
Public VBEPersonBuffVSF() As Variant '���`���A���-VS-F��
Public VBEPersonBuffVSS() As Variant '���`���A���-VS-S��
Public AtkingckVSS(1 To 8, 1 To 3) As Variant  '�ޯ��T�@��-S��(�ޯ�ҰʽX)
Public AtkingckVSF(1 To 8, 1 To 1) As Variant '�ޯ��T�@��-F��(�Ƶ��r��)
Public VBEAtkingVSF(1 To 2, 1 To 3, 1 To 2) As Variant 'VBE>VS�����ܼƲΤ@���-F��
Public VBEAtkingVSS(0 To 20) As Variant 'VBE>VS�����ܼƲΤ@���-S��
Public VBEPageCardNumVS() As Variant '���εP���-VS��
Public VBEVSBuffNum(1 To 2) As Variant '���`���A�M��-���`���A��2�Ӽƭ�-VS��
Public VBEStageNum() As Integer '���涥�q�t��-���涥�q�h�γ~�����Ȯ��ܼ�(0.���涥�q��/1~���N���e)
Public VBEVSStageNum() As Variant '���涥�q�t��-���涥�q�h�γ~�����ܼ�-VS��
Public VBEStageRemoveBuffAllNum() As Boolean '���涥�q�t��-���涥�q73-���`���A��������M��-���`���A�O�_��ĳ�аO�Ȯ��ܼ�
Public VBEActualStatusVS(1 To 2, 1 To 3, 1 To 2) As Variant '�H����ڪ��A���-VS��
Public VBEStage7xAtkingInformation As String '���涥�q7x(76���A�[�J/77�Ѱ�)-�ޯ�ߤ@�ѧO�X�Ȯ��x�s�ܼ�
Public VBEVSStageInfoList As New Collection '���涥�q�t�ΦU�h�Ƭ�����T
Sub ���涥�q�t���`�D�n�{��_�H���D�ʧޯ�(ByVal uscom As Integer, ByVal personnum As Integer, ByVal atkingnum As Integer, ByVal ns As Integer, ByVal PersonBattleNum As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim atkingvssnum As Integer
    If vbecommadtotplayNow > 10 Then Exit Sub '���涥�q�̰�10�h
    If ���涥�q�t����.���涥�q�t��_����(atkingnum, ns, VBEPerson(uscom, personnum, 3, atkingnum, 11), uscom, personnum) = True Then
           ���涥�q�t����.���涥�q�t��_�ǳ��ܼƲΦX��� uscom, VBEStageNumMain, PersonBattleNum
           vbecommadnum(6, vbecommadtotplayNow) = PersonBattleNum
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           atkingvssnum = (uscom - 1) * 12 + (4 * personnum - 4) + atkingnum
           ������O��.������O���`�{�ǰ��� ���涥�q�t��_����}��_�H���D�ʧޯ���(atkingnum, ns, uscom, personnum), atkingvssnum, uscom, atkingnum, ns, vbecommadtotplayNow
    End If
End Sub
Sub ���涥�q�t���`�D�n�{��_�H���Q�ʧޯ�(ByVal uscom As Integer, ByVal personnum As Integer, ByVal atkingnum As Integer, ByVal ns As Integer, ByVal PersonBattleNum As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim passivevssnum As Integer, PassivePersonType As Integer  '�Ȯ��ܼ�
    If vbecommadtotplayNow > 10 Then Exit Sub '���涥�q�̰�10�h
    If ���涥�q�t����.���涥�q�t��_����(atkingnum, ns, VBEPerson(uscom, personnum, 3, atkingnum, 3), uscom, personnum) = True Then
           ���涥�q�t����.���涥�q�t��_�ǳ��ܼƲΦX��� uscom, VBEStageNumMain, PersonBattleNum
           If PersonBattleNum > 1 Then PassivePersonType = 2 Else PassivePersonType = 1
           vbecommadnum(6, vbecommadtotplayNow) = PersonBattleNum
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           passivevssnum = (uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24
           ������O��.������O���`�{�ǰ��� ���涥�q�t��_����}��_�H���Q�ʧޯ���(atkingnum, ns, uscom, personnum, PassivePersonType), passivevssnum, uscom, atkingnum, ns, vbecommadtotplayNow
    End If
End Sub
Sub ���涥�q73_���O_���`���A����_�����M��(ByVal uscom As Integer, ByVal num As Integer, Optional ByVal isHPW As Boolean = False)
Dim vbecommadnumSecond As Integer '���h���涥�q�s����
Dim buffobj As clsStatus
Dim i As Integer
If �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num)).Count > 0 Then
    '=======================
    vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
    '=======================
    Dim VBEStageNumMainSec(0 To 1) As Integer
    VBEStageNumMainSec(0) = 73
    If isHPW = True Then
        VBEStageNumMainSec(1) = 2
    Else
        VBEStageNumMainSec(1) = 1
    End If
    ReDim VBEStageRemoveBuffAllNum(1 To �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num)).Count) As Boolean
    '=======================
    For i = 1 To �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num)).Count
        Set buffobj = �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, num))(i)
        Vss_EventRemoveBuffActionOffNum = 0
        ���涥�q�t���`�D�n�{��_���`���A uscom, ����ݾ��H��������(uscom, num), buffobj.Identifier, 73, num, VBEStageNumMainSec, vbecommadnumSecond
        If Vss_EventRemoveBuffActionOffNum = 1 Then
             VBEStageRemoveBuffAllNum(i) = True
        End If
    Next
    '=======================
    ���涥�q�t��_�ŧi�}�l�ε��� 2
    '=======================
End If
End Sub
Sub ���涥�q73_���O_���`���A����_�S�w�M��(ByVal uscom As Integer, ByVal num As Integer, ByVal akstr As String)
Dim vbecommadnumSecond As Integer '���h���涥�q�s����
'=======================
vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
'=======================
Dim VBEStageNumMainSec(0 To 1) As Integer
VBEStageNumMainSec(0) = 73
VBEStageNumMainSec(1) = 1
'=======================
Vss_EventRemoveBuffActionOffNum = 0
���涥�q�t����.���涥�q�t���`�D�n�{��_���`���A uscom, ����ݾ��H��������(uscom, num), akstr, 73, num, VBEStageNumMainSec, vbecommadnumSecond
'=======================
���涥�q�t��_�ŧi�}�l�ε��� 2
'=======================
End Sub
Sub ���涥�q73_���O_���`���A����_�D�ʲM��(ByVal uscom As Integer, ByVal num As Integer, ByVal akstr As String)
Dim vbecommadnumSecond As Integer '���h���涥�q�s����
'=======================
vbecommadnumSecond = ���涥�q�t��_�ŧi�}�l�ε���(1)
'=======================
Dim VBEStageNumMainSec(0 To 1) As Integer
VBEStageNumMainSec(0) = 73
VBEStageNumMainSec(1) = 0
'=======================
���涥�q�t���`�D�n�{��_���`���A uscom, ����ݾ��H��������(uscom, num), akstr, 73, num, VBEStageNumMainSec, vbecommadnumSecond
'=======================
���涥�q�t��_�ŧi�}�l�ε��� 2
'=======================
End Sub
Sub ���涥�q�t���`�D�n�{��_���涥�q�}�l(ByVal uscomFirst As Integer, ByVal ns As Integer, ByVal nstype As Integer)
    Dim vbecommadtotplayNow As Integer '���h���涥�q�s����
    '==nstype(1.������(����)/2.�u����@��(����)/3.������(������)/4.�u����@��(������)
    Dim i As Integer, k As Integer, w As Integer, atkingnum As Integer
    Dim n As clsStatus
    '=======================
    vbecommadtotplayNow = ���涥�q�t��_�ŧi�}�l�ε���(1)
    '=======================
    Dim VBEStageNumMain() As Integer
    If UBound(VBEStageNum) = 0 Then
        ReDim VBEStageNumMain(1 To 1) As Integer
    Else
        ReDim VBEStageNumMain(0 To UBound(VBEStageNum)) As Integer
        For i = 0 To UBound(VBEStageNum)
           VBEStageNumMain(i) = VBEStageNum(i)
        Next
    End If
    '=======================
    Dim uscom As Integer
    For k = 1 To 2
        If k = 1 Then
            If uscomFirst = 1 Then uscom = 1 Else uscom = 2
        Else
            If uscomFirst = 1 Then uscom = 2 Else uscom = 1
        End If
        '==================�H����ڪ��A
        For w = 1 To 3
            If �H����ڪ��A��Ʈw(uscom, ����ݾ��H��������(uscom, w), 1) <> "" Then
                ���涥�q�t���`�D�n�{��_�H����ڪ��A uscom, ����ݾ��H��������(uscom, w), ns, w, VBEStageNumMain, vbecommadtotplayNow
            End If
        Next
        '==================���`���A
        For w = 1 To 3
            For Each n In �H�����`���A�C��(uscom, ����ݾ��H��������(uscom, w))
                ���涥�q�t���`�D�n�{��_���`���A uscom, ����ݾ��H��������(uscom, w), n.Identifier, ns, w, VBEStageNumMain, vbecommadtotplayNow
            Next
        Next
        '==================�Q�ʧޯ�
        For w = 1 To 3
            For atkingnum = 5 To 8
                If atkingck(uscom, ����ݾ��H��������(uscom, w), atkingnum, 1) = 1 Or Vss_PersonAtkingOffNum(uscom, ����ݾ��H��������(uscom, w), atkingnum) = 0 Then
                    ���涥�q�t���`�D�n�{��_�H���Q�ʧޯ� uscom, ����ݾ��H��������(uscom, w), atkingnum, ns, w, VBEStageNumMain, vbecommadtotplayNow
                End If
            Next
        Next
        '==================�D�ʧޯ�
        For w = 1 To 3
            For atkingnum = 1 To 4
                If (nstype <= 2 And atkingck(uscom, ����ݾ��H��������(uscom, w), atkingnum, 1) = 1) Or _
                    (nstype > 2 And Vss_PersonAtkingOffNum(uscom, ����ݾ��H��������(uscom, w), atkingnum) = 0) Then
                    ���涥�q�t���`�D�n�{��_�H���D�ʧޯ� uscom, ����ݾ��H��������(uscom, w), atkingnum, ns, w, VBEStageNumMain, vbecommadtotplayNow
                End If
            Next
        Next
        '=====================
        If nstype = 2 Or nstype = 4 Then Exit For
    Next
    '=================
    ReDim VBEStageNum(0) As Integer
    ���涥�q�t��_�ŧi�}�l�ε��� 2
    '=================
End Sub
Function ���涥�q�t��_�ŧi�}�l�ε���(ByVal n As Integer) As Integer
    Select Case n
        Case 1 '==�}�l
            vbecommadtotplay = vbecommadtotplay + 1
            ReDim Preserve vbecommadnum(1 To 7, vbecommadtotplay)
            ReDim Preserve vbecommadstr(1 To 3, vbecommadtotplay)
        Case 2 '==����
            vbecommadtotplay = vbecommadtotplay - 1
            ReDim Preserve vbecommadnum(1 To 7, vbecommadtotplay)
            ReDim Preserve vbecommadstr(1 To 3, vbecommadtotplay)
    End Select
    ���涥�q�t��_�ŧi�}�l�ε��� = vbecommadtotplay
End Function
Function ���涥�q�t��_����(ByVal atkingnum As Integer, ByVal ns As Integer, ByVal akstr As String, ByVal uscom As Integer, ByVal personnum As Integer) As Boolean
    If Formsetting.checktest.Value = 0 Then On Error GoTo vscheckerr
    Dim vsstr1 As String, vsstr2 As String, vsstr3() As String, vsstr4 As String
    Dim textlinea As String, str As String
    Dim k As Integer, p As Integer
    '==========================
    If (uscom = 1 And liveus(personnum) <= 0 And ����H����ԤH��(uscom, 2) <> personnum) Or _
       (uscom = 2 And livecom(personnum) <= 0 And ����H����ԤH��(uscom, 2) <> personnum) Then
       ���涥�q�t��_���� = False
       Exit Function
    End If
    '==========================
    Select Case atkingnum
        Case Is <= 4  '==�D�ʧޯ�
            If VBEVSSAtkingStr(uscom, personnum, atkingnum, 1) <> "" Then
                vsstr1 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + atkingnum).Run("main", 1)
                vsstr2 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + atkingnum).Run("main", 2)
                If ����H����ԤH��(uscom, 2) <> personnum Then
                    If ns = 42 Or ns = 43 Or ns = 44 Then '�����\�D���W����ϥΥX�P�ƥ�
                        ���涥�q�t��_���� = False
                        Exit Function
                    End If
                    vsstr4 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + atkingnum).Run("main", 8)
                Else
                    vsstr4 = "ON"
                End If
                '==================
                vsstr3 = Split(vsstr2, "#")
                For k = 0 To UBound(vsstr3)
                    If vsstr1 = akstr And (ns = Val(vsstr3(k))) And vsstr4 = "ON" Then
                        ���涥�q�t��_���� = True
                        Exit Function
                    End If
                Next
                ���涥�q�t��_���� = False
            Else
                ���涥�q�t��_���� = False
            End If
        Case Is <= 8  '==�Q�ʧޯ�
            If VBEVSSAtkingStr(uscom, personnum, atkingnum, 1) <> "" Then
                vsstr1 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("main", 1)
                vsstr2 = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("main", 2)
                '==================
                vsstr3 = Split(vsstr2, "#")
                For k = 0 To UBound(vsstr3)
                    If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                        ���涥�q�t��_���� = True
                        Exit Function
                    End If
                Next
                ���涥�q�t��_���� = False
            Else
                ���涥�q�t��_���� = False
            End If
        Case 9 '==���`���A
             For p = 1 To UBound(VBEVSSBuffStr1)
                If VBEVSSBuffStr1(p) = akstr Then
                    vsstr1 = FormMainMode.PEAFvssc(p + 54).Run("main", 1)
                    vsstr2 = FormMainMode.PEAFvssc(p + 54).Run("main", 2)
                    '==================
                    vsstr3 = Split(vsstr2, "#")
                    For k = 0 To UBound(vsstr3)
                        If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                            ���涥�q�t��_���� = True
                            Exit Function
                        End If
                    Next
                End If
             Next
             ���涥�q�t��_���� = False
        Case 10 '==�H����ڪ��A
             For p = 1 To UBound(VBEVSSActualStatusStr1)
                If VBEVSSActualStatusStr1(p) = akstr Then
                    vsstr1 = FormMainMode.PEAFvssc((uscom - 1) * 3 + personnum + 48).Run("main", 1)
                    vsstr2 = FormMainMode.PEAFvssc((uscom - 1) * 3 + personnum + 48).Run("main", 2)
                    '==================
                    vsstr3 = Split(vsstr2, "#")
                    For k = 0 To UBound(vsstr3)
                        If vsstr1 = akstr And (ns = Val(vsstr3(k))) Then
                            ���涥�q�t��_���� = True
                            Exit Function
                        End If
                    Next
                End If
             Next
             ���涥�q�t��_���� = False
    End Select
Exit Function
    '==============================
vscheckerr:
    ���涥�q�t��_���~�T���q�� 2, "1[" & uscom & "-" & ns & "-" & akstr & "]"
End Function
Function ���涥�q�t��_�ɮ�Ū�J(ByVal atkingnum As Integer, ByVal name As String, ByVal uscom As Integer) As Boolean
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsloaderror
   Select Case atkingnum
        Case Is <= 4
            Dim textlinea As String, str As String
'            Open app_path & "character\" & name & "\" & VBEVSSAtkingStr(uscom, atkingnum, 2) For Input As #1 '������
'            Open App.Path & "\test\input1.txt" For Input As #1
            
            Do Until EOF(1)
               Line Input #1, textlinea
               str = str & textlinea & vbCrLf
            Loop
            Close
            If str <> "" Then
                FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * ����H����ԤH��(uscom, 2) - 4) + atkingnum).AddCode str
                ���涥�q�t��_�ɮ�Ū�J = True
            Else
                ���涥�q�t��_�ɮ�Ū�J = False
            End If
        Case Else
    
    End Select
'=====================================
Exit Function
vsloaderror:
���涥�q�t��_�ɮ�Ū�J = False
'=====================================
End Function
Function ���涥�q�t��_����}��_�H���D�ʧޯ���(ByVal atkingnum As Integer, ByVal ns As Integer, ByVal uscom As Integer, ByVal personnum As Integer) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
    If �@��t����.ProgramIsOnWine = True Then
        Dim wineObj As New clsWineobj
        ���涥�q�t��_wine�ܼƲΦX��ƪ���g�J wineObj, ns, 0
        ���涥�q�t��_����}��_�H���D�ʧޯ��� = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + atkingnum).Run("WineEntryPoint", wineObj)
    Else
        ���涥�q�t��_����}��_�H���D�ʧޯ��� = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + atkingnum).Run("atking", ns, VBEPersonVS, VBEPageCardNumVS, atkingpagetotVS, VBEPersonBuffVSF, VBEPersonBuffVSS, AtkingckVSS, AtkingckVSF, VBEAtkingVSF, VBEAtkingVSS, VBEActualStatusVS, VBEVSStageNum)
    End If
'=====================================
Exit Function
'===========
Dim i As Integer
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
���涥�q�t����.���涥�q�t��_��l_�}��Ū�J�{��
GoTo VssAdminReTry
'===========
vsgoerror:
���涥�q�t��_���~�T���q�� 2, "2[1-" & atkingnum & "]"
'=====================================

End Function
Function ���涥�q�t��_����}��_�H���Q�ʧޯ���(ByVal atkingnum As Integer, ByVal ns As Integer, ByVal uscom As Integer, ByVal personnum As Integer, ByVal PassivePersonType As Integer) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
    Dim PassivePersonTypeVSS As Variant
    PassivePersonTypeVSS = PassivePersonType
    If �@��t����.ProgramIsOnWine = True Then
        Dim wineObj As New clsWineobj
        ���涥�q�t��_wine�ܼƲΦX��ƪ���g�J wineObj, ns, PassivePersonType
        ���涥�q�t��_����}��_�H���Q�ʧޯ��� = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("WineEntryPoint", wineObj)
    Else
        ���涥�q�t��_����}��_�H���Q�ʧޯ��� = FormMainMode.PEAFvssc((uscom - 1) * 12 + (4 * personnum - 4) + (atkingnum - 4) + 24).Run("passive", ns, VBEPersonVS, VBEPageCardNumVS, atkingpagetotVS, VBEPersonBuffVSF, VBEPersonBuffVSS, AtkingckVSS, AtkingckVSF, VBEAtkingVSF, VBEAtkingVSS, VBEActualStatusVS, PassivePersonTypeVSS, VBEVSStageNum)
    End If
'=====================================
Exit Function
'===========
Dim i As Integer
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
���涥�q�t����.���涥�q�t��_��l_�}��Ū�J�{��
GoTo VssAdminReTry
'===========
vsgoerror:
���涥�q�t��_���~�T���q�� 2, "2[2-" & atkingnum & "]"
'=====================================

End Function

Function ���涥�q�t��_����}��_���`���A��(ByVal vssnum As Integer, ByVal ns As Integer, ByVal BuffPersonType As Integer, ByVal akstr As String) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
    Dim BuffPersonTypeVSS As Variant
    BuffPersonTypeVSS = BuffPersonType
    If �@��t����.ProgramIsOnWine = True Then
        Dim wineObj As New clsWineobj
        ���涥�q�t��_wine�ܼƲΦX��ƪ���g�J wineObj, ns, BuffPersonType
        ���涥�q�t��_����}��_���`���A�� = FormMainMode.PEAFvssc(vssnum).Run("WineEntryPoint", wineObj)
    Else
        ���涥�q�t��_����}��_���`���A�� = FormMainMode.PEAFvssc(vssnum).Run("buff", ns, atkingpagetotVS, VBEAtkingVSF, VBEAtkingVSS, VBEVSBuffNum, BuffPersonTypeVSS, VBEVSStageNum)
    End If
'=====================================
Exit Function
'===========
Dim i As Integer
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
���涥�q�t����.���涥�q�t��_��l_�}��Ū�J�{��
GoTo VssAdminReTry
'===========
vsgoerror:
���涥�q�t��_���~�T���q�� 2, "2[3-" & akstr & "]"
'=====================================

End Function
Function ���涥�q�t��_����}��_�H����ڪ��A��(ByVal vssnum As Integer, ByVal ns As Integer, ByVal ActualStatusPersonType As Integer, ByVal akstr As String) As String
   If Formsetting.checktest.Value = 0 Then On Error GoTo vsgoerror
VssAdminReTry:
    Dim ActualStatusPersonTypeVSS As Variant
    ActualStatusPersonTypeVSS = ActualStatusPersonType
    If �@��t����.ProgramIsOnWine = True Then
        Dim wineObj As New clsWineobj
        ���涥�q�t��_wine�ܼƲΦX��ƪ���g�J wineObj, ns, ActualStatusPersonType
        ���涥�q�t��_����}��_�H����ڪ��A�� = FormMainMode.PEAFvssc(vssnum).Run("WineEntryPoint", wineObj)
    Else
        ���涥�q�t��_����}��_�H����ڪ��A�� = FormMainMode.PEAFvssc(vssnum).Run("ActualStatus", ns, VBEPersonVS, VBEPageCardNumVS, atkingpagetotVS, VBEPersonBuffVSF, VBEPersonBuffVSS, VBEAtkingVSF, VBEAtkingVSS, ActualStatusPersonTypeVSS, VBEVSStageNum)
    End If
'=====================================
Exit Function
'===========
Dim i As Integer
For i = 1 To (Val(54) + Val(UBound(VBEVSSBuffStr2)))
   FormMainMode.PEAFvssc(i).Reset
Next
���涥�q�t����.���涥�q�t��_��l_�}��Ū�J�{��
GoTo VssAdminReTry
'===========
vsgoerror:
���涥�q�t��_���~�T���q�� 2, "2[4-" & akstr & "]"
'=====================================

End Function
Sub ���涥�q�t��_�ǳ��ܼƲΦX���(ByVal uscom As Integer, ByRef VBEStageNumMain() As Integer, ByVal PersonBattleNum As Integer)
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
    ReDim VBEPageCardNumVS(1 To ���εP����d�����j������(1), 1 To 6) As Variant '���εP���-VS��
'    Erase VBEVSStageNum '���涥�q�t��-���涥�q�h�γ~�����ܼ�-VS��
    ReDim VBEVSStageNum(1 To UBound(VBEStageNumMain)) As Variant '���涥�q�t��-���涥�q�h�γ~�����ܼ�-VS��
    Erase VBEActualStatusVS '�H����ڪ��A���-VS��
    '===========================
    Dim q As Integer, w As Integer, rr As Integer, tempc As Integer, buffobj As clsStatus
    Dim i As Integer, j As Integer, k As Integer, m As Integer, p As Integer
    
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
            For i = 1 To ���εP����d�����j������(1)
                For j = 1 To 6
                    If j = 1 Or j = 3 Then
                       Select Case pagecardnum(i, j)
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
                            Case "HPW"
                                VBEPageCardNumVS(i, j) = 9
                            Case Else
                                VBEPageCardNumVS(i, j) = 0
                       End Select
                    Else
                       VBEPageCardNumVS(i, j) = Val(pagecardnum(i, j))
                    End If
                Next
            Next
            '======================
            '(1 To 2, 1 To 5)
            For i = 1 To 2
                For j = 1 To 5
                    atkingpagetotVS(i, j) = atkingpagetot(i, j)
                Next
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
            VBEAtkingVSS(0) = PersonBattleNum
            VBEAtkingVSS(1) = pageqlead(1)
            VBEAtkingVSS(2) = pageglead(1)
            VBEAtkingVSS(3) = pageqlead(2)
            VBEAtkingVSS(4) = pageglead(2)
            VBEAtkingVSS(5) = �Y���淾�q�Ȯ��ܼ�(2)
            VBEAtkingVSS(6) = movecp
            VBEAtkingVSS(7) = Val(�������m��l�`��(1))
            VBEAtkingVSS(8) = Val(�������m��l�`��(2))
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            VBEAtkingVSS(14) = �Y���淾�q�Ȯ��ܼ�(5)
            VBEAtkingVSS(15) = �Y���淾�q�Ȯ��ܼ�(6)
            VBEAtkingVSS(16) = moveturn
            VBEAtkingVSS(17) = ����H����ԤH��(1, 1)
            VBEAtkingVSS(18) = ����H����ԤH��(2, 1)
            VBEAtkingVSS(19) = �P�`���q��(1)
            VBEAtkingVSS(20) = �P�`���q��(2)
            Select Case turnatk
                Case 1
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 2
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                    VBEAtkingVSS(16) = 0
                Case 4, 6
                    VBEAtkingVSS(12) = 1
                    VBEAtkingVSS(13) = 0
                    VBEAtkingVSS(16) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
                    VBEAtkingVSS(16) = 0
             End Select
             '=========================
             If LBound(VBEStageNumMain) = 0 Then
                    Select Case VBEStageNumMain(0)
                        Case 41, 46, 48 '���涥�q41/46/48(����洫/�ˮ`/�^�_)
                            For i = 1 To UBound(VBEStageNumMain)
                                    If VBEStageNumMain(i) = -1 Or VBEStageNumMain(i) = -2 Then
                                        VBEVSStageNum(i) = Abs(VBEStageNumMain(i))
                                    Else
                                        VBEVSStageNum(i) = VBEStageNumMain(i)
                                    End If
                            Next
                        Case 76, 77
                            For i = 1 To UBound(VBEStageNumMain) - 1
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                            Next
                            VBEVSStageNum(3) = VBEStage7xAtkingInformation
                        Case 42, 43, 44
                            VBEVSStageNum(1) = Abs(VBEStageNumMain(1))
                        Case Else
                            For i = 1 To UBound(VBEStageNumMain)
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                            Next
                    End Select
             Else
                    VBEVSStageNum(1) = 0 '�L���
             End If
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
            For i = 1 To ���εP����d�����j������(1)
                For j = 1 To 6
                     If j = 1 Or j = 3 Then
                       Select Case pagecardnum(i, j)
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
                            Case "HPW"
                                VBEPageCardNumVS(i, j) = 9
                            Case Else
                                VBEPageCardNumVS(i, j) = 0
                       End Select
                    ElseIf j = 5 Then
                       If Val(pagecardnum(i, j)) = 2 Then
                           VBEPageCardNumVS(i, j) = 1
                        ElseIf Val(pagecardnum(i, j)) = 1 Then
                           VBEPageCardNumVS(i, j) = 2
                        Else
                           VBEPageCardNumVS(i, j) = 0
                        End If
                    Else
                       VBEPageCardNumVS(i, j) = Val(pagecardnum(i, j))
                    End If
                Next
            Next
            '======================
            '(1 To 2, 1 To 5)
            For i = 1 To 2
                If i = 1 Then q = 2 Else q = 1
                For j = 1 To 5
                    atkingpagetotVS(i, j) = atkingpagetot(q, j)
                Next
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
            VBEAtkingVSS(0) = PersonBattleNum
            VBEAtkingVSS(1) = pageqlead(2)
            VBEAtkingVSS(2) = pageglead(2)
            VBEAtkingVSS(3) = pageqlead(1)
            VBEAtkingVSS(4) = pageglead(1)
            VBEAtkingVSS(5) = �Y���淾�q�Ȯ��ܼ�(2)
            VBEAtkingVSS(6) = movecp
            VBEAtkingVSS(7) = Val(�������m��l�`��(2))
            VBEAtkingVSS(8) = Val(�������m��l�`��(1))
            VBEAtkingVSS(9) = BattleTurn
            VBEAtkingVSS(10) = app_path
            VBEAtkingVSS(11) = BattleCardNum
            VBEAtkingVSS(14) = �Y���淾�q�Ȯ��ܼ�(6)
            VBEAtkingVSS(15) = �Y���淾�q�Ȯ��ܼ�(5)
            If moveturn = 2 Then VBEAtkingVSS(16) = 1 Else VBEAtkingVSS(16) = 2
            VBEAtkingVSS(17) = ����H����ԤH��(2, 1)
            VBEAtkingVSS(18) = ����H����ԤH��(1, 1)
            VBEAtkingVSS(19) = �P�`���q��(2)
            VBEAtkingVSS(20) = �P�`���q��(1)
            Select Case turnatk
                Case 1
                    VBEAtkingVSS(12) = 4
                    VBEAtkingVSS(13) = 2
                Case 2
                    VBEAtkingVSS(12) = 3
                    VBEAtkingVSS(13) = 1
                Case 3
                    VBEAtkingVSS(12) = 2
                    VBEAtkingVSS(13) = 0
                    VBEAtkingVSS(16) = 0
                Case 4, 6
                    VBEAtkingVSS(12) = 1
                    VBEAtkingVSS(13) = 0
                    VBEAtkingVSS(16) = 0
                Case Else
                    VBEAtkingVSS(12) = 0
                    VBEAtkingVSS(13) = 0
                    VBEAtkingVSS(16) = 0
             End Select
             '=========================
             If LBound(VBEStageNumMain) = 0 Then
                    Select Case VBEStageNumMain(0)
                        Case 2, 3, 4, 70, 71 '���涥�q2/3/4/70/71(���q-���ʫe)
                            VBEVSStageNum(1) = VBEStageNumMain(2)
                            VBEVSStageNum(2) = VBEStageNumMain(1)
                            VBEVSStageNum(3) = VBEStageNumMain(4)
                            VBEVSStageNum(4) = VBEStageNumMain(3)
                        Case 41, 46, 48 '���涥�q41/46/48(����洫/�ˮ`/�^�_)
                            For i = 1 To UBound(VBEStageNumMain)
                                If VBEStageNumMain(i) = -1 Then
                                    VBEVSStageNum(i) = 2
                                ElseIf VBEStageNumMain(i) = -2 Then
                                    VBEVSStageNum(i) = 1
                                Else
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                                End If
                            Next
                        Case 76, 77
                            If VBEStageNumMain(1) = 1 Then VBEVSStageNum(1) = 2 Else VBEVSStageNum(1) = 1
                            VBEVSStageNum(2) = VBEStageNumMain(2)
                            VBEVSStageNum(3) = VBEStage7xAtkingInformation
                        Case 62 '�ޯ�ĪG�i��h���Y���
                            VBEVSStageNum(1) = VBEStageNumMain(2)
                            VBEVSStageNum(2) = VBEStageNumMain(1)
                            VBEVSStageNum(3) = VBEStageNumMain(4)
                            VBEVSStageNum(4) = VBEStageNumMain(3)
                            VBEVSStageNum(5) = VBEStageNumMain(5)
                        Case 42, 43, 44
                            If VBEStageNumMain(1) = -1 Then VBEVSStageNum(1) = 2 Else VBEVSStageNum(1) = 1
                        Case Else
                            For i = 1 To UBound(VBEStageNumMain)
                                    VBEVSStageNum(i) = VBEStageNumMain(i)
                            Next
                    End Select
             Else
                    VBEVSStageNum(1) = 0 '�L���
             End If
   End Select
End Sub
Sub ���涥�q�t��_��l_�}��Ū�J�{��()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
Dim atknum As Integer, uscomn As Integer, pnnum As Integer, buffnum As Integer
atknum = 1: uscomn = 1: pnnum = 1: buffnum = 1
Dim tot As Integer, textlinea As String, str As String
tot = 1
Do
     textlinea = ""
     str = ""
     Select Case tot
         Case Is <= 24
                If VBEVSSAtkingStr(uscomn, pnnum, atknum, 1) <> "" Then
                    Open app_path & "character\" & VBEPerson(uscomn, pnnum, 1, 1, 2) & "\" & VBEVSSAtkingStr(uscomn, pnnum, atknum, 2) For Input As #1
                    
                    Do Until EOF(1)
                       Line Input #1, textlinea
                       str = str & textlinea & vbCrLf
                    Loop
                    Close
                    If str <> "" Then
                        FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum).AddCode str
                        If �@��t����.ProgramIsOnWine = True Then ���涥�q�t����.���涥�q�t��_�[�JWine�{���i�J�I (uscomn - 1) * 12 + (4 * pnnum - 4) + atknum
                    End If
                End If
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
         Case Is <= 48
                If VBEVSSAtkingStr(uscomn, pnnum, atknum + 4, 1) <> "" Then
                    Open app_path & "character\" & VBEPerson(uscomn, pnnum, 1, 1, 2) & "\" & VBEVSSAtkingStr(uscomn, pnnum, atknum + 4, 2) For Input As #1
                    
                    Do Until EOF(1)
                       Line Input #1, textlinea
                       str = str & textlinea & vbCrLf
                    Loop
                    Close
                    If str <> "" Then
                        FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum + 24).AddCode str
                        If �@��t����.ProgramIsOnWine = True Then ���涥�q�t����.���涥�q�t��_�[�JWine�{���i�J�I (uscomn - 1) * 12 + (4 * pnnum - 4) + atknum + 24
                    End If
                End If
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
         Case Is <= 54
                
         Case Else
                Open VBEVSSBuffStr2(buffnum) For Input As #1
                
                Do Until EOF(1)
                   Line Input #1, textlinea
                   str = str & textlinea & vbCrLf
                Loop
                Close
                If str <> "" Then
                    FormMainMode.PEAFvssc(tot).AddCode str
                    If �@��t����.ProgramIsOnWine = True Then ���涥�q�t����.���涥�q�t��_�[�JWine�{���i�J�I tot
                End If
                buffnum = buffnum + 1
    End Select
    tot = tot + 1
Loop Until tot > (Val(54) + Val(UBound(VBEVSSBuffStr2)))
'===============
Exit Sub
vssyserror:
If tot <= 48 Then
    ���涥�q�t��_���~�T���q�� 1, "3[" & VBEVSSAtkingStr(uscomn, pnnum, atknum, 1) & "]"
ElseIf tot > 48 And tot <= 54 Then
    ���涥�q�t��_���~�T���q�� 1, "3[" & VBEVSSActualStatusStr2(buffnum) & "]"
Else
    ���涥�q�t��_���~�T���q�� 1, "3[" & VBEVSSBuffStr2(buffnum) & "]"
End If
'===============
End Sub
Sub ���涥�q�t�ιC����l�`�{��()
       ���涥�q�t����.���涥�q�t��_���`���A���}���j�M
       ���涥�q�t����.���涥�q�t��_�H����ڪ��A���}���j�M
       ���涥�q�t����.���涥�q�t��_��l_�}������Х� (Val(54) + Val(UBound(VBEVSSBuffStr2)))
       ���涥�q�t����.���涥�q�t��_��l_�}��Ū�J�{��
       ���涥�q�t����.���涥�q�t��_��l_�H���D�ʤγQ�ʧޯ������Ū�J
       ���涥�q�t����.���涥�q�t��_��l_���`���A�����Ū�J
       ���涥�q�t����.���涥�q�t��_��l_�H����ڪ��A�����Ū�J
End Sub
Sub ���涥�q�t��_��l_�}������Х�(ByVal num As Integer)
       If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
       Dim i As Integer
        '==========
        For i = 1 To num
           Load FormMainMode.PEAFvssc(i)
        Next
        '==========
        '==========
        For i = 1 To num
           FormMainMode.PEAFvssc(i).Reset
        Next
        '==========
'===============
Exit Sub
vssyserror:
���涥�q�t��_���~�T���q�� 1, "2"
'===============
End Sub
Sub ���涥�q�t��_���`���A���}���j�M()
  If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
  Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
  mypath = App.Path & "\Buff\"
  mydir = Dir(mypath, vbDirectory) ' ��M�Ĥ@�Ӥl�ؿ��C
  ReDim DirectoryBuff(0)
  ReDim VBEVSSBuffStr1(0)
  ReDim VBEVSSBuffStr2(0)
  Do While True
        Do While mydir <> ""
            ' ���L�ثe���ؿ��ΤW�h�ؿ��C
            If mydir <> "." And mydir <> ".." Then
                ' �ϥΦ줸���ӽT�w MyName �N��@�ؿ��C
                If (GetAttr(mypath & mydir) And vbDirectory) = vbDirectory Then
                    Debug.Print mydir ' �N�ؿ��W����ܥX�ӡC
                    ReDim Preserve DirectoryBuff(UBound(DirectoryBuff) + 1)
                    DirectoryBuff(UBound(DirectoryBuff)) = mypath + mydir
                Else
                    If Utils.GetExtName(mydir) = "ulevsbf" And Index >= 1 Then
                        ���涥�q�t����.���涥�q�t��_��l_���`���A���}���[�J���� mydir, DirectoryBuff(Index) & "\"
                    ElseIf Utils.GetExtName(mydir) = "ulevsbf" And Index = 0 Then
                        ���涥�q�t����.���涥�q�t��_��l_���`���A���}���[�J���� mydir, App.Path & "\Buff\"
                    End If
                End If
            End If
            mydir = Dir()
        Loop
        Index = Index + 1
        If Index > UBound(DirectoryBuff) Then Exit Do
        mypath = DirectoryBuff(Index) + "\"
        mydir = Dir(mypath, vbDirectory)
  Loop
'===============
Exit Sub
vssyserror:
���涥�q�t��_���~�T���q�� 1, "1"
'===============
End Sub
Sub ���涥�q�t��_�H����ڪ��A���}���j�M()
  If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
  Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
  mypath = App.Path & "\CharacterActualStatus\"
  mydir = Dir(mypath, vbDirectory) ' ��M�Ĥ@�Ӥl�ؿ��C
  ReDim DirectoryBuff(0)
  ReDim VBEVSSActualStatusStr1(0)
  ReDim VBEVSSActualStatusStr2(0)
  Do While True
        Do While mydir <> ""
            ' ���L�ثe���ؿ��ΤW�h�ؿ��C
            If mydir <> "." And mydir <> ".." Then
                ' �ϥΦ줸���ӽT�w MyName �N��@�ؿ��C
                If (GetAttr(mypath & mydir) And vbDirectory) = vbDirectory Then
                    Debug.Print mydir ' �N�ؿ��W����ܥX�ӡC
                    ReDim Preserve DirectoryBuff(UBound(DirectoryBuff) + 1)
                    DirectoryBuff(UBound(DirectoryBuff)) = mypath + mydir
                Else
                    If Utils.GetExtName(mydir) = "ulevsc" And Index >= 1 Then
                        ���涥�q�t����.���涥�q�t��_��l_�H����ڪ��A���}���[�J���� mydir, DirectoryBuff(Index) & "\"
                    ElseIf Utils.GetExtName(mydir) = "ulevsc" And Index = 0 Then
                        ���涥�q�t����.���涥�q�t��_��l_�H����ڪ��A���}���[�J���� mydir, App.Path & "\CharacterActualStatus\"
                    End If
                End If
            End If
            mydir = Dir()
        Loop
        Index = Index + 1
        If Index > UBound(DirectoryBuff) Then Exit Do
        mypath = DirectoryBuff(Index) + "\"
        mydir = Dir(mypath, vbDirectory)
  Loop
'===============
Exit Sub
vssyserror:
���涥�q�t��_���~�T���q�� 1, "6"
'===============
End Sub
Sub ���涥�q�t��_��l_���`���A���}���[�J����(ByVal str1 As String, ByVal PersonName As String)
    ReDim Preserve VBEVSSBuffStr2(UBound(VBEVSSBuffStr2) + 1)
    VBEVSSBuffStr2(UBound(VBEVSSBuffStr2)) = PersonName & str1
End Sub
Sub ���涥�q�t��_��l_�H����ڪ��A���}���[�J����(ByVal str1 As String, ByVal PersonName As String)
    ReDim Preserve VBEVSSActualStatusStr2(UBound(VBEVSSActualStatusStr2) + 1)
    VBEVSSActualStatusStr2(UBound(VBEVSSActualStatusStr2)) = PersonName & str1
End Sub
Sub ���涥�q�t��_��l_�H���D�ʤγQ�ʧޯ������Ū�J()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
Dim vsstr As String, ���r��() As String
Dim atknum As Integer, uscomn As Integer, pnnum As Integer
Dim tot As Integer, i As Integer, k As Integer
atknum = 1: uscomn = 1: pnnum = 1
tot = 1
Do
    vsstr = ""
    Select Case tot
         Case Is <= 24
                If VBEVSSAtkingStr(uscomn, pnnum, atknum, 1) <> "" Then
                    For i = 3 To 7
                        vsstr = FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum).Run("main", i)
                        ���r�� = Split(vsstr, "#")
                        '==================
                        Select Case i
                            Case 3
                                VBEPerson(uscomn, pnnum, 3, atknum, 1) = ���r��(0)
                            Case 4
                                VBEPerson(uscomn, pnnum, 3, atknum, 2) = ���r��(0)
                                VBEPerson(uscomn, pnnum, 3, atknum, 8) = ���r��(1)
                            Case 5
                                VBEPerson(uscomn, pnnum, 3, atknum, 3) = ���r��(0)
                                VBEPerson(uscomn, pnnum, 3, atknum, 9) = ���r��(1)
                            Case 6
                                VBEPerson(uscomn, pnnum, 3, atknum, 4) = ���r��(0)
                                VBEPerson(uscomn, pnnum, 3, atknum, 10) = ���r��(1)
                            Case 7
                                VBEPerson(uscomn, pnnum, 3, atknum, 5) = ""
                                For k = 0 To UBound(���r��)
                                     VBEPerson(uscomn, pnnum, 3, atknum, 5) = VBEPerson(uscomn, pnnum, 3, atknum, 5) & ���r��(k)
                                Next
                        End Select
                    Next
                End If
                '=================================================
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
         Case Is <= 48
                If VBEVSSAtkingStr(uscomn, pnnum, atknum + 4, 1) <> "" Then
                    For i = 3 To 4
                        vsstr = FormMainMode.PEAFvssc((uscomn - 1) * 12 + (4 * pnnum - 4) + atknum + 24).Run("main", i)
                        ���r�� = Split(vsstr, "#")
                        '==================
                        Select Case i
                            Case 3
                                VBEPerson(uscomn, pnnum, 3, atknum + 4, 1) = ���r��(0)
                            Case 4
                                VBEPerson(uscomn, pnnum, 3, atknum + 4, 2) = ""
                                For k = 0 To UBound(���r��)
                                     VBEPerson(uscomn, pnnum, 3, atknum + 4, 2) = VBEPerson(uscomn, pnnum, 3, atknum + 4, 2) & ���r��(k)
                                Next
                        End Select
                    Next
                End If
                '=================================================
                atknum = atknum + 1
                If atknum > 4 Then
                    atknum = 1
                    pnnum = pnnum + 1
                End If
                If pnnum > 3 Then
                    pnnum = 1
                    uscomn = uscomn + 1
                End If
                If uscomn > 2 Then
                    atknum = 1: uscomn = 1: pnnum = 1
                End If
    End Select
    tot = tot + 1
Loop Until tot > 48
'===============
Exit Sub
vssyserror:
If tot <= 24 Then
    ���涥�q�t��_���~�T���q�� 1, "4[" & uscomn & "," & atknum & "]"
Else
    ���涥�q�t��_���~�T���q�� 1, "4[" & uscomn & "," & atknum + 4 & "]"
End If
'===============
End Sub
Sub ���涥�q�t��_��l_���`���A�����Ū�J()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
ReDim VBEVSSBuffStr1(UBound(VBEVSSBuffStr2))
Dim vsstr As String
Dim i As Integer

For i = 1 To UBound(VBEVSSBuffStr2)
    vsstr = FormMainMode.PEAFvssc(i + 54).Run("main", 1)
    VBEVSSBuffStr1(i) = vsstr
Next
'===============
Exit Sub
vssyserror:
���涥�q�t��_���~�T���q�� 1, "5[" & VBEVSSBuffStr2(i) & "]"
'===============
End Sub
Sub ���涥�q�t��_��l_�H����ڪ��A�����Ū�J()
If Formsetting.checktest.Value = 0 Then On Error GoTo vssyserror
ReDim VBEVSSActualStatusStr1(UBound(VBEVSSActualStatusStr2))
Dim vsstr As String, textlinea As String, str As String
Dim i As Integer

For i = 1 To UBound(VBEVSSActualStatusStr2)
    Open VBEVSSActualStatusStr2(i) For Input As #1
    Do Until EOF(1)
       Line Input #1, textlinea
       str = str & textlinea & vbCrLf
    Loop
    Close
    If str <> "" Then
        FormMainMode.PEAFvssc(49).AddCode str
    End If
    vsstr = FormMainMode.PEAFvssc(49).Run("main", 1)
    VBEVSSActualStatusStr1(i) = vsstr
    FormMainMode.PEAFvssc(49).Reset
Next
'===============
Exit Sub
vssyserror:
���涥�q�t��_���~�T���q�� 1, "7[" & VBEVSSActualStatusStr2(i) & "]"
'===============
End Sub
Sub ���涥�q�t���`�D�n�{��_���`���A(ByVal uscom As Integer, ByVal personnum As Integer, ByVal akstr As String, ByVal ns As Integer, ByVal PersonBattleNum As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim buffvssnum As Integer, BuffPersonType As Integer, buffobj As clsStatus '�Ȯ��ܼ�
    Dim p As Integer
    
    If vbecommadtotplayNow > 10 Then Exit Sub '���涥�q�̰�10�h
    If ���涥�q�t����.���涥�q�t��_����(9, ns, akstr, uscom, personnum) = True Then
           Set buffobj = �H�����`���A�C��(uscom, personnum)(akstr)
           ���涥�q�t����.���涥�q�t��_�ǳ��ܼƲΦX��� uscom, VBEStageNumMain, PersonBattleNum
           If PersonBattleNum > 1 Then BuffPersonType = 2 Else BuffPersonType = 1
           vbecommadnum(6, vbecommadtotplayNow) = PersonBattleNum
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           Erase VBEVSBuffNum '���`���A�M��-���`���A��2�Ӽƭ�-VS��
           For p = 1 To UBound(VBEVSSBuffStr1)
                 If VBEVSSBuffStr1(p) = akstr Then
                     buffvssnum = p + 54
                     VBEVSBuffNum(1) = buffobj.Value
                     VBEVSBuffNum(2) = buffobj.Total
                     Exit For
                 End If
            Next
           ������O��.������O���`�{�ǰ��� ���涥�q�t��_����}��_���`���A��(buffvssnum, ns, BuffPersonType, VBEVSSBuffStr1(p)), buffvssnum, uscom, 9, ns, vbecommadtotplayNow
    End If
End Sub
Sub ���涥�q�t���`�D�n�{��_�H����ڪ��A(ByVal uscom As Integer, ByVal personnum As Integer, ByVal ns As Integer, ByVal PersonBattleNum As Integer, ByRef VBEStageNumMain() As Integer, ByVal vbecommadtotplayNow As Integer)
    Dim ActualStatusvssnum As Integer, ActualStatusPersonType As Integer '�Ȯ��ܼ�
    If vbecommadtotplayNow > 10 Then Exit Sub '���涥�q�̰�10�h
    If ���涥�q�t����.���涥�q�t��_����(10, ns, �H����ڪ��A��Ʈw(uscom, personnum, 1), uscom, personnum) = True Then
           ���涥�q�t����.���涥�q�t��_�ǳ��ܼƲΦX��� uscom, VBEStageNumMain, PersonBattleNum
           If PersonBattleNum > 1 Then ActualStatusPersonType = 2 Else ActualStatusPersonType = 1
           vbecommadnum(6, vbecommadtotplayNow) = PersonBattleNum
           vbecommadnum(7, vbecommadtotplayNow) = personnum
           ActualStatusvssnum = (((uscom - 1) * 3) + personnum) + 48
           ������O��.������O���`�{�ǰ��� ���涥�q�t��_����}��_�H����ڪ��A��(ActualStatusvssnum, ns, ActualStatusPersonType, �H����ڪ��A��Ʈw(uscom, personnum, 1)), ActualStatusvssnum, uscom, 10, ns, vbecommadtotplayNow
    End If
End Sub
Function ���涥�q�t��_�j�M���b���椧���涥�q(ByVal vscmdname As String) As Integer
    Dim i As Integer
    For i = 1 To vbecommadtotplay
         If vbecommadstr(1, i) = vscmdname Then
             ���涥�q�t��_�j�M���b���椧���涥�q = i
             Exit Function
         End If
    Next
    '==========�p�G�䤣���
    ���涥�q�t��_�j�M���b���椧���涥�q = 0
End Function
Sub ���涥�q�t��_���~�T���q��(ByVal num As Integer, ByVal num1 As String)
MsgBox "���涥�q���~(03-" & num & "-" & num1 & ")�G" & Chr(10) & "�t�Ω�Ū���θ����}�����O�ɵo�Ϳ��~�C" & Chr(10) & Chr(10) & "(" & Err.Number & "):" & Err.Description, vbCritical
End
End Sub
Sub ���涥�q�t��_�[�JWine�{���i�J�I(ByVal num As Integer)
Dim strcode As String
Select Case num
         Case Is <= 24
            strcode = "Function WineEntryPoint(wineObj)" & vbCrLf & "WineEntryPoint = atking(wineObj.oNs, wineObj.GetArray(""VBEPersonVS""), wineObj.GetArray(""VBEPageCardNumVS""), wineObj.GetArray(""AtkingPagetotVS""), wineObj.GetArray(""VBEPersonBuffVSF""), wineObj.GetArray(""VBEPersonBuffVSS""), wineObj.GetArray(""AtkingckVSS""), wineObj.GetArray(""AtkingckVSF""), wineObj.GetArray(""VBEAtkingVSF""), wineObj.GetArray(""VBEAtkingVSS""), wineObj.GetArray(""VBEActualStatusVS""), wineObj.GetArray(""VBEVSStageNum""))" & vbCrLf & "End Function"
         Case Is <= 48
            strcode = "Function WineEntryPoint(wineObj)" & vbCrLf & "WineEntryPoint = passive(wineObj.oNs, wineObj.GetArray(""VBEPersonVS""), wineObj.GetArray(""VBEPageCardNumVS""), wineObj.GetArray(""AtkingPagetotVS""), wineObj.GetArray(""VBEPersonBuffVSF""), wineObj.GetArray(""VBEPersonBuffVSS""), wineObj.GetArray(""AtkingckVSS""), wineObj.GetArray(""AtkingckVSF""), wineObj.GetArray(""VBEAtkingVSF""), wineObj.GetArray(""VBEAtkingVSS""), wineObj.GetArray(""VBEActualStatusVS""), wineObj.oPersonType, wineObj.GetArray(""VBEVSStageNum""))" & vbCrLf & "End Function"
         Case Is <= 54
            strcode = "Function WineEntryPoint(wineObj)" & vbCrLf & "WineEntryPoint = ActualStatus(wineObj.oNs, wineObj.GetArray(""VBEPersonVS""), wineObj.GetArray(""VBEPageCardNumVS""), wineObj.GetArray(""AtkingPagetotVS""), wineObj.GetArray(""VBEPersonBuffVSF""), wineObj.GetArray(""VBEPersonBuffVSS""), wineObj.GetArray(""VBEAtkingVSF""), wineObj.GetArray(""VBEAtkingVSS""), wineObj.oPersonType, wineObj.GetArray(""VBEVSStageNum""))" & vbCrLf & "End Function"
         Case Else
            strcode = "Function WineEntryPoint(wineObj)" & vbCrLf & "WineEntryPoint = buff(wineObj.oNs, wineObj.GetArray(""AtkingPagetotVS""), wineObj.GetArray(""VBEAtkingVSF""), wineObj.GetArray(""VBEAtkingVSS""), wineObj.GetArray(""VBEVSBuffNum""), wineObj.oPersonType, wineObj.GetArray(""VBEVSStageNum""))" & vbCrLf & "End Function"
End Select
FormMainMode.PEAFvssc(num).AddCode strcode
End Sub
Sub ���涥�q�t��_wine�ܼƲΦX��ƪ���g�J(ByRef wineObj As clsWineobj, ByVal ns As Integer, ByVal persontype As Integer)
wineObj.oNs = ns
wineObj.oPersonType = persontype
wineObj.AddInformation "VBEAtkingVSF", ���涥�q�t����.VBEAtkingVSF
wineObj.AddInformation "VBEAtkingVSS", ���涥�q�t����.VBEAtkingVSS
wineObj.AddInformation "AtkingPagetotVS", ���涥�q�t����.atkingpagetotVS
wineObj.AddInformation "VBEPersonVS", ���涥�q�t����.VBEPersonVS
wineObj.AddInformation "VBEPageCardNumVS", ���涥�q�t����.VBEPageCardNumVS
wineObj.AddInformation "AtkingckVSS", ���涥�q�t����.AtkingckVSS
wineObj.AddInformation "AtkingckVSF", ���涥�q�t����.AtkingckVSF
wineObj.AddInformation "VBEPersonBuffVSF", ���涥�q�t����.VBEPersonBuffVSF
wineObj.AddInformation "VBEPersonBuffVSS", ���涥�q�t����.VBEPersonBuffVSS
wineObj.AddInformation "VBEActualStatusVS", ���涥�q�t����.VBEActualStatusVS
wineObj.AddInformation "VBEVSBuffNum", ���涥�q�t����.VBEVSBuffNum
wineObj.AddInformation "VBEVSStageNum", ���涥�q�t����.VBEVSStageNum
End Sub
