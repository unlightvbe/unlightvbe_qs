Attribute VB_Name = "�@��t����"
Option Explicit
Public app_path As String  '���|�]�w�X
Public ����H����ԤH��(1 To 2, 1 To 2) As Integer '�����Ԩ���H�Ƭ�����(1.�ϥΪ�/2.�q��,1.�`�@�H��/2.�ثe�ĴX��)
Public ����ݾ��H��������(1 To 2, 1 To 3) As Integer '����ݾ�����H���s��������(1.�ϥΪ�/2.�q��,1.���W����/2~3.�ݾ������n��s��)
Public tr1num As Integer 'PEStartForm�p�ƾ��Ȯ��ܼ�
Public PEAEtr1num As Integer 'PEAttackingEndingForm�p�ƾ��Ȯ��ܼ�
Public st As Integer, sq As Integer, swq As Integer, cardusq As Integer, cardcomq As Integer   'PEAttackingStartForm�p�ƾ��Ȯ��ܼ�
Public �Ĥ@���Ұ�Ū�J�{�ǼаO As Boolean '�Ĥ@���Ұʵ{��Ū�J�{�ǼаO��
Public ����Ū�J���� As String 'PEStartForm����Ū�J���Ȯɬ�����
Public �����ˬd����ؼм� As Integer '�����ˬd����p�ƾ��ؼм�
Public �q�����O�_�w�X�{ As Boolean '���ҳq�����O�_�w�g�X�{�Ȯ��ܼ�
Public ProgramIsOnWine As Boolean '�{���O�_�B��Wine���ҤU����

Sub �P�_�r��_FormMainMode()
Dim i, a As Integer
a = 14
If FormMainMode.PEStext1.FontName <> "Bradley Gratis" Then
    '===========PEAttackingForm
    FormMainMode.pageul.FontSize = 24
    FormMainMode.bloodnumcom1.FontSize = 20
    FormMainMode.bloodnumcom2.FontSize = 10
    FormMainMode.bloodnumus1.FontSize = 20
    FormMainMode.bloodnumus2.FontSize = 10
    '===========PEGameFreeModeSettingForm
    For i = 1 To 3
        FormMainMode.PEGFusbi1(i).FontSize = a
        FormMainMode.PEGFusbi2(i).FontSize = a
        FormMainMode.PEGFusbi3(i).FontSize = a
        FormMainMode.PEGFcardcompi1(i).FontSize = a
        FormMainMode.PEGFcardcompi2(i).FontSize = a
        FormMainMode.PEGFcardcompi3(i).FontSize = a
    Next
    '===========PEAttackingStartForm
    For i = 1 To 3
        FormMainMode.PEASusbi1(i).FontSize = a
        FormMainMode.PEASusbi2(i).FontSize = a
        FormMainMode.PEASusbi3(i).FontSize = a
        FormMainMode.PEAScardcompi1(i).FontSize = a
        FormMainMode.PEAScardcompi2(i).FontSize = a
        FormMainMode.PEAScardcompi3(i).FontSize = a
    Next
End If
End Sub
     
Sub �}�l�C���i��{��()
Dim i As Integer, m As Integer, n As Integer, u As Integer, personvsp As Integer, perosntempCount As Integer
Dim personnameg(1 To 2, 1 To 3) As String '�H���H����ܬ�����
�@��t����.���ļ��� 11
���ϥΪ̨ƥ� = False
���q���ƥ� = False
�q����ƥ�d�O�_�X����ܼ� = False
�O�_���ʶ��q����p�P�_�{�� = False
'==============�M���Ҧ��ܼƭȻP��_�]�w
�@��t����.�M���԰��t�ΩҦ��ܼƭ�
�@��t����.�M���԰��t�ζ}�l���]�w��
'==============������(�H��)
For i = 1 To 3
    If FormMainMode.personnameus(i).Text = "�m�H���n" Then
        personnameg(1, i) = ""
    Else
        personnameg(1, i) = FormMainMode.personnameus(i).Text
    End If
    If FormMainMode.personnamecom(i).Text = "�m�H���n" Then
        personnameg(2, i) = ""
    Else
        personnameg(2, i) = FormMainMode.personnamecom(i).Text
    End If
Next
If FormMainMode.opnpersonvs(1).Value = True Then
    personvsp = 1
    If FormMainMode.personnameus(1).Text = "�m�H���n" Then
        Randomize
        perosntempCount = 0
        Do
            m = Int(Rnd() * (FormMainMode.personnameus(1).ListCount - 1)) + 1
            �H���t����.�d���H����TŪ�J_�G���q_�ϥΪ� FormMainMode.personnameus(1).List(m), 1
            If Formsetting.chkusesimilarlevel.Value = 0 Then
                n = Int(Rnd() * (FormMainMode.personlevelus(1).ListCount - 1 + 1)) + 0
                Exit Do
            End If
            n = ����H���H���ۦ����ſ��(1, 1)
            If n <> -1 Then
                Exit Do
            Else
                perosntempCount = perosntempCount + 1
                If perosntempCount > 3 Then
                    n = Int(Rnd() * (FormMainMode.personlevelus(1).ListCount - 1 + 1)) + 0
                End If
            End If
        Loop Until (perosntempCount > 3)
'        �H���t����.�d���H����TŪ�J_�T���q_�ϥΪ� FormMainMode.personnameus(1).List(m), FormMainMode.personlevelus(1).List(n), 1, 1
        �H���t����.�M������H����T�ܼ� 1, 1
        �H���t����.�d���H����TŪ�J_�T���q FormMainMode.personnameus(1).List(m), FormMainMode.personlevelus(1).List(n), 1, 1
'        �H���t����.�d���H����TŪ�J_�|���q_�ϥΪ� FormMainMode.personnameus(1).List(m), 1
        �H���t����.�d���H����TŪ�J_�|���q FormMainMode.personnameus(1).List(m), 1, 1
    End If
    If FormMainMode.personnamecom(1).Text = "�m�H���n" Then
        Randomize
        perosntempCount = 0
        Do
            m = Int(Rnd() * (FormMainMode.personnamecom(1).ListCount - 1)) + 1
            �H���t����.�d���H����TŪ�J_�G���q_�q�� FormMainMode.personnamecom(1).List(m), 1
            If Formsetting.chkusesimilarlevel.Value = 0 Then
                n = Int(Rnd() * (FormMainMode.personlevelcom(1).ListCount - 1 + 1)) + 0
                Exit Do
            End If
            n = ����H���H���ۦ����ſ��(2, 1)
            If n <> -1 Then
                Exit Do
            Else
                perosntempCount = perosntempCount + 1
                If perosntempCount > 3 Then
                    n = Int(Rnd() * (FormMainMode.personlevelcom(1).ListCount - 1 + 1)) + 0
                End If
            End If
        Loop Until (perosntempCount > 3)
'        �H���t����.�d���H����TŪ�J_�T���q_�q�� FormMainMode.personnamecom(1).List(m), FormMainMode.personlevelcom(1).List(n), 1, 2
        �H���t����.�M������H����T�ܼ� 2, 1
        �H���t����.�d���H����TŪ�J_�T���q FormMainMode.personnamecom(1).List(m), FormMainMode.personlevelcom(1).List(n), 1, 2
'        �H���t����.�d���H����TŪ�J_�|���q_�q�� formmainmode.personnamecom(1).List(m), 1
    End If
Else
    personvsp = 3
    For i = 1 To 3
        If FormMainMode.personnameus(i).Text = "�m�H���n" Then
            Randomize
            perosntempCount = 0
            Do
                m = Int(Rnd() * (FormMainMode.personnameus(i).ListCount - 1)) + 1
                personnameg(1, i) = FormMainMode.personnameus(i).List(m)
                �H���t����.�d���H����TŪ�J_�G���q_�ϥΪ� FormMainMode.personnameus(i).List(m), i
                If Formsetting.chkusesimilarlevel.Value = 0 Then
                    n = Int(Rnd() * (FormMainMode.personlevelus(i).ListCount - 1 + 1)) + 0
                    Exit Do
                End If
                n = ����H���H���ۦ����ſ��(1, i)
                If n <> -1 Then
                    Exit Do
                Else
                    perosntempCount = perosntempCount + 1
                    If perosntempCount > 3 Then
                        n = Int(Rnd() * (FormMainMode.personlevelus(i).ListCount - 1 + 1)) + 0
                    End If
                End If
            Loop Until (perosntempCount > 3)
'            �H���t����.�d���H����TŪ�J_�T���q_�ϥΪ� FormMainMode.personnameus(i).List(m), FormMainMode.personlevelus(i).List(n), i, 1
'            �H���t����.�d���H����TŪ�J_�|���q_�ϥΪ� FormMainMode.personnameus(i).List(m), i
            �H���t����.�M������H����T�ܼ� 1, i
            �H���t����.�d���H����TŪ�J_�T���q FormMainMode.personnameus(i).List(m), FormMainMode.personlevelus(i).List(n), i, 1
            �H���t����.�d���H����TŪ�J_�|���q FormMainMode.personnameus(i).List(m), i, 1
            ��s�H���M��_�ϥΪ̤�_�ܧ�_�}�l�H�� i, personnameg(1, 1), personnameg(1, 2), personnameg(1, 3)
        End If
        If FormMainMode.personnamecom(i).Text = "�m�H���n" Then
            Randomize
            perosntempCount = 0
            Do
                m = Int(Rnd() * (FormMainMode.personnamecom(i).ListCount - 1)) + 1
                personnameg(2, i) = FormMainMode.personnamecom(i).List(m)
                �H���t����.�d���H����TŪ�J_�G���q_�q�� FormMainMode.personnamecom(i).List(m), i
                If Formsetting.chkusesimilarlevel.Value = 0 Then
                    n = Int(Rnd() * (FormMainMode.personlevelcom(i).ListCount - 1 + 1)) + 0
                    Exit Do
                End If
                n = ����H���H���ۦ����ſ��(2, i)
                If n <> -1 Then
                    Exit Do
                Else
                    perosntempCount = perosntempCount + 1
                    If perosntempCount > 3 Then
                        n = Int(Rnd() * (FormMainMode.personlevelcom(i).ListCount - 1 + 1)) + 0
                    End If
                End If
            Loop Until (perosntempCount > 3)
'            �H���t����.�d���H����TŪ�J_�T���q_�q�� FormMainMode.personnamecom(i).List(m), FormMainMode.personlevelcom(i).List(n), i, 2
            �H���t����.�M������H����T�ܼ� 2, i
            �H���t����.�d���H����TŪ�J_�T���q FormMainMode.personnamecom(i).List(m), FormMainMode.personlevelcom(i).List(n), i, 2
'            �H���t����.�d���H����TŪ�J_�|���q_�q�� formmainmode.personnamecom(i).List(m), i
            ��s�H���M��_�q����_�ܧ�_�}�l�H�� i, personnameg(2, 1), personnameg(2, 2), personnameg(2, 3)
        End If
    Next
End If
'=================�j�ð��ﶵ
If Formsetting.�j�ð��ﶵ.Value = 1 Then
   If personvsp = 1 Then
        VBEPerson(1, 1, 1, 3, 1) = 99
        VBEPerson(2, 1, 1, 3, 1) = 99
   ElseIf personvsp = 3 Then
        For i = 1 To 3
            VBEPerson(1, i, 1, 3, 1) = 99
            VBEPerson(2, i, 1, 3, 1) = 99
        Next
   End If
End If
'=======�ˬd�]�w
If ����H����ܪŭ��ˬd(personvsp) = False Then
   FormHint.Left = FormMainMode.Left + 1185
   FormHint.Top = FormMainMode.Top + 3030
   FormHint.Show 1
Else
'-------------
    Select Case personvsp
       Case 1
            For i = 1 To 1
                FormMainMode.cardusname(i).Caption = VBEPerson(1, 1, 1, 1, 1)
                FormMainMode.cardusspname(i).Caption = VBEPerson(1, 1, 1, 1, 3)
                FormMainMode.cardcomname(i).Caption = VBEPerson(2, 1, 1, 1, 1)
                FormMainMode.cardcomspname(i).Caption = VBEPerson(2, 1, 1, 1, 3)
                FormMainMode.PEAScardus(i).Picture = LoadPicture(VBEPerson(1, 1, 1, 5, 5))
                FormMainMode.PEAScardcom(i).Picture = LoadPicture(VBEPerson(2, 1, 1, 5, 5))
                FormMainMode.PEASusbi1(i).Caption = VBEPerson(1, 1, 1, 3, 1)
                FormMainMode.PEASusbi2(i).Caption = VBEPerson(1, 1, 1, 3, 2)
                FormMainMode.PEASusbi3(i).Caption = VBEPerson(1, 1, 1, 3, 3)
                FormMainMode.PEAScardcompi1(i).Caption = VBEPerson(2, 1, 1, 3, 1)
                FormMainMode.PEAScardcompi2(i).Caption = VBEPerson(2, 1, 1, 3, 2)
                FormMainMode.PEAScardcompi3(i).Caption = VBEPerson(2, 1, 1, 3, 3)
            Next
       Case 3
            For i = 1 To ����H����ԤH��(1, 1)
                FormMainMode.cardusname(i).Caption = VBEPerson(1, i, 1, 1, 1)
                FormMainMode.cardusspname(i).Caption = VBEPerson(1, i, 1, 1, 3)
                FormMainMode.PEAScardus(i).Picture = LoadPicture(VBEPerson(1, i, 1, 5, 5))
                FormMainMode.PEASusbi1(i).Caption = VBEPerson(1, i, 1, 3, 1)
                FormMainMode.PEASusbi2(i).Caption = VBEPerson(1, i, 1, 3, 2)
                FormMainMode.PEASusbi3(i).Caption = VBEPerson(1, i, 1, 3, 3)
            Next
            For i = 1 To ����H����ԤH��(2, 1)
                FormMainMode.cardcomname(i).Caption = VBEPerson(2, i, 1, 1, 1)
                FormMainMode.cardcomspname(i).Caption = VBEPerson(2, i, 1, 1, 3)
                FormMainMode.PEAScardcom(i).Picture = LoadPicture(VBEPerson(2, i, 1, 5, 5))
                FormMainMode.PEAScardcompi1(i).Caption = VBEPerson(2, i, 1, 3, 1)
                FormMainMode.PEAScardcompi2(i).Caption = VBEPerson(2, i, 1, 3, 2)
                FormMainMode.PEAScardcompi3(i).Caption = VBEPerson(2, i, 1, 3, 3)
            Next
    End Select
    '===============================
    For i = 1 To ����H����ԤH��(1, 1)
        If Val(VBEPerson(1, i, 1, 3, 5)) = 1 Then
            FormMainMode.PEASusbi1(i).Left = 300
            FormMainMode.PEASusbi1(i).Top = 3220
            FormMainMode.PEASusbi2(i).Left = 960
            FormMainMode.PEASusbi2(i).Top = 3220
            FormMainMode.PEASusbi3(i).Left = 1820
            FormMainMode.PEASusbi3(i).Top = 3220
        Else
            FormMainMode.PEASusbi1(i).Left = 555
            FormMainMode.PEASusbi1(i).Top = 3240
            FormMainMode.PEASusbi2(i).Left = 1200
            FormMainMode.PEASusbi2(i).Top = 3240
            FormMainMode.PEASusbi3(i).Left = 1920
            FormMainMode.PEASusbi3(i).Top = 3240
        End If
    Next
    For i = 1 To ����H����ԤH��(2, 1)
        If Val(VBEPerson(2, i, 1, 3, 5)) = 1 Then
            FormMainMode.PEAScardcompi1(i).Left = 230
            FormMainMode.PEAScardcompi1(i).Top = 3220
            FormMainMode.PEAScardcompi2(i).Left = 960
            FormMainMode.PEAScardcompi2(i).Top = 3220
            FormMainMode.PEAScardcompi3(i).Left = 1820
            FormMainMode.PEAScardcompi3(i).Top = 3220
        Else
            FormMainMode.PEAScardcompi1(i).Left = 480
            FormMainMode.PEAScardcompi1(i).Top = 3240
            FormMainMode.PEAScardcompi2(i).Left = 1200
            FormMainMode.PEAScardcompi2(i).Top = 3240
            FormMainMode.PEAScardcompi3(i).Left = 1920
            FormMainMode.PEAScardcompi3(i).Top = 3240
        End If
    Next
    '===========================�԰��t�ΥD���Ū�J(��)
    ���涥�q�t����.���涥�q�t�ιC����l�`�{��
    �԰��t����.�C������d������Х�
    '===============================��ԹϤ����J(�H���զX)-�e�m���q
    If Formsetting.BGM���.Text = "�m�H��-�a�ϲզX�n" Then
        If Formsetting.��Ԧa�Ͽ��.Text = "�m�H���n" Then
            Randomize
            m = Int(Rnd() * Val(Formsetting.��Ԧa�Ͽ��.ListCount - 1)) + 1
            Formsetting.��Ԧa�Ͽ��.ListIndex = m
        Else
            Formsetting.��Ԧa�Ͽ��_Click
        End If
     End If
    '===============================��ԭ��ָ��J
    FormMainMode.cMusicPlayer(0).MusicStop
    If Formsetting.BGM���.Text = "�m�H���n" Then
        Randomize
        m = Int(Rnd() * Val(Formsetting.BGM���.ListCount - 1)) + 1
        Formsetting.BGM���.ListIndex = m
    End If
       Select Case Formsetting.BGM���.Text
         Case "�ª�"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulbgm04.mp3"
         Case "�B�ʴ�`(�s)"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm003.mp3"
         Case "�H��Ӧa"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm004.mp3"
         Case "�ܤB���|�櫰��"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm000.mp3"
         Case "���b�˪L"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm001.mp3"
         Case "�U������"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm002.mp3"
         Case "���Y����"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm005.mp3"
         Case "���ɯ"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm006.mp3"
         Case "�]��ù�e�����J"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm008.mp3"
         Case "�Q�i�����´�"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm010.mp3"
         Case "ÿ�e�઺���"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm007.mp3"
         Case "�ƨg�s��"
           FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\bgm009.mp3"
         Case "�m��L�n"
           FormMainMode.cMusicPlayer(0).Filepath = Formsetting.lopnmusictext.Caption
        End Select
      '===========================��ԹϤ����J
    If Formsetting.��Ԧa�Ͽ��.Text = "�m�H���n" Then
        Randomize
        m = Int(Rnd() * Val(Formsetting.��Ԧa�Ͽ��.ListCount - 1)) + 1
        Formsetting.��Ԧa�Ͽ��.ListIndex = m
    End If
    If Formsetting.��Ԧa�Ͽ��.Text = "�m��L�n" Then
        FormMainMode.PEAttackingForm.Picture = LoadPicture(Formsetting.lopnmapjpgtext.Caption)
        FormMainMode.PEAttackingStartForm.Picture = LoadPicture(Formsetting.lopnmapjpgtext.Caption)
        FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(Formsetting.lopnmapjpgtext.Caption)
    Else
        FormMainMode.PEAttackingForm.Picture = LoadPicture(app_path & "gif\system\map\" & Formsetting.��Ԧa�Ͽ��.ListIndex & ".jpg")
        FormMainMode.PEAttackingStartForm.Picture = LoadPicture(app_path & "gif\system\map\" & Formsetting.��Ԧa�Ͽ��.ListIndex & ".jpg")
        FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(app_path & "gif\system\map\" & Formsetting.��Ԧa�Ͽ��.ListIndex & ".jpg")
    End If
    '=============================================
    If Formsetting.chkusenewpage.Value = 1 Then
        �԰��t����.���εP�a�ϵP�����t�m Formsetting.��Ԧa�Ͽ��.Text
    Else
        �԰��t����.���εP�a�ϵP�����t�m 0
    End If
    '=============================================
    FormMainMode.cMusicPlayer(0).IsLoop = True
    '=================================================
    �԰��Y�뤶���H����ø�ϸ��|������(1) = VBEPerson(1, 1, 1, 5, 3)
    �԰��Y�뤶���H����ø�ϸ��|������(2) = VBEPerson(2, 1, 1, 5, 3)
    '=================================================
    If Formsetting.chkusenewinterface.Value = 1 Then
        �t����ܬɭ������� = 2
    Else
        �t����ܬɭ������� = 1
    End If
    '=======================================�ƥ�d/����d����]�w
    �@��t����.�ۥѾ԰��Ҧ��]�w���U���]�wŪ�J�{��
    �԰��t����.�C������P����ŧi�{��
    �԰��t����.�ƥ�d�B�z_�p��i��
    �԰��t����.�ƥ�d�B�z_��l_�ϥΪ̤�
    �԰��t����.�ƥ�d�B�z_��l_�q����
    �԰��t����.�ƥ�d�B�z_���w_�ϥΪ̤�
    �԰��t����.�ƥ�d�B�z_���w_�q����
    '===============================================
    �@��t����.�԰��t�Ϊ��Ū�J�{��
    �@��t����.�԰��t�ζ}�l���Ū�J�{��
    '===================
    �@��t����.�D���_PEAttackingStartForm���
    FormMainMode.cMusicPlayer(0).MusicPlay
    FormMainMode.PEGameFreeModeSettingForm.Visible = False
End If
End Sub
Function ����H���H���ۦ����ſ��(ByVal uscom As Integer, ByVal num As Integer) As Integer
Dim levelmark As Integer
Dim tempnum As Integer
Dim personlist As ComboBox
Dim temppass As Boolean
Dim i As Integer
levelmark = Formsetting.cbsimilarlevel.ListIndex
tempnum = 1
Select Case uscom
    Case 1
        Set personlist = FormMainMode.personlevelus(num)
    Case 2
        Set personlist = FormMainMode.personlevelcom(num)
End Select
temppass = False
Do
    If temppass = False Then
        For i = 0 To personlist.ListCount - 1
            If personlist.List(i) = Formsetting.cbsimilarlevel.List(levelmark) Then
                ����H���H���ۦ����ſ�� = i
                Exit Function
            End If
        Next
    End If
    Select Case tempnum
        Case 1
            levelmark = Formsetting.cbsimilarlevel.ListIndex - 1
        Case 2
            levelmark = Formsetting.cbsimilarlevel.ListIndex + 1
        Case 3
            levelmark = Formsetting.cbsimilarlevel.ListIndex - 2
        Case 4
            levelmark = Formsetting.cbsimilarlevel.ListIndex + 2
    End Select
    tempnum = tempnum + 1
    If levelmark > Formsetting.cbsimilarlevel.ListCount - 1 Or levelmark < 0 Then temppass = True Else temppass = False
Loop Until (tempnum > 4)

'=======��������ۦ�������
����H���H���ۦ����ſ�� = -1
End Function
Function ����H����ܪŭ��ˬd(ByVal personvs As Integer) As Boolean
����H����ܪŭ��ˬd = True
Dim i As Integer
Select Case personvs
   Case 1
      If FormMainMode.personnameus(1).Text = "" Or FormMainMode.personnamecom(1).Text = "" Then
          ����H����ܪŭ��ˬd = False
      Else
          ����H����ܪŭ��ˬd = True
          ����H����ԤH��(1, 1) = 1
          ����H����ԤH��(2, 1) = 1
      End If
   Case 3
      ����H����ԤH��(1, 1) = 3
      ����H����ԤH��(2, 1) = 3
      For i = 1 To 3
          If FormMainMode.personnameus(i).Text = "" Or FormMainMode.personnamecom(i).Text = "" Then
              ����H����ܪŭ��ˬd = False
'              Exit Function
          End If
       Next
       If i > 3 And ����H����ܪŭ��ˬd = False Then
          If FormMainMode.personnameus(1).Text = "" Or FormMainMode.personnamecom(1).Text = "" Then
              '======���Ŧܤ֦�1�줧�W�w
              Exit Function
          End If
          '===========�ˬd�ŭȨ�����
           For i = 2 To 3
              If FormMainMode.personnameus(i).Text = "" Then
                  ����H����ԤH��(1, 1) = ����H����ԤH��(1, 1) - 1
              End If
              If FormMainMode.personnamecom(i).Text = "" Then
                  ����H����ԤH��(2, 1) = ����H����ԤH��(2, 1) - 1
              End If
           Next
           ����H����ܪŭ��ˬd = True
        Else
           ����H����ܪŭ��ˬd = True
        End If
End Select
End Function
Sub �d���H����T���J_�j�M�ɮ�()
Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
  mypath = App.Path & "\character\"
  mydir = Dir(mypath, vbDirectory) ' ��M�Ĥ@�Ӥl�ؿ��C
  ReDim DirectoryBuff(0)
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
                    If Utils.GetExtName(mydir) = "uleci" Then
                        �H���t����.�d���H����TŪ�J_�춥�q mypath + mydir
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
'MsgBox "1-5-2-4"
End Sub
Function �ƥ�d��Ʈw(ByVal name As String, ByVal �ؼ� As Integer) As String
'�ؼ�:1-�Ǧ^������/2-�Ǧ^�ƥ�d�ɮצW��/3-�Ǧ^�����ƭ�
Select Case name
    Case "�C1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=1=ATK-�C=1"
        End Select
    Case "�C2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks2"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=2=ATK-�C=2"
        End Select
    Case "�C3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks3"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=3=ATK-�C=3"
        End Select
   Case "�C4"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks4"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=4=ATK-�C=4"
        End Select
    Case "�C5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks5"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=5=ATK-�C=5"
        End Select
    Case "�C6"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks6"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=6=ATK-�C=6"
        End Select
    Case "�C7"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks7"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=7=ATK-�C=7"
        End Select
    Case "�C8"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks8"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=8=ATK-�C=8"
        End Select
    Case "�j1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atkg1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=1=ATK-�j=1"
        End Select
    Case "�j2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atkg2"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=2=ATK-�j=2"
        End Select
    Case "�j3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atkg3"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=3=ATK-�j=3"
        End Select
    Case "�j4"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg4"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=4=ATK-�j=4"
        End Select
    Case "�j5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg5"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=5=ATK-�j=5"
        End Select
    Case "�j6"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg6"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=6=ATK-�j=6"
        End Select
    Case "�j7"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg7"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=7=ATK-�j=7"
        End Select
    Case "�j8"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg8"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=8=ATK-�j=8"
        End Select
    Case "��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-def1"
           Case 3
              �ƥ�d��Ʈw = "DEF=1=DEF=1"
        End Select
    Case "��2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-def2"
           Case 3
              �ƥ�d��Ʈw = "DEF=2=DEF=2"
        End Select
    Case "��3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-def3"
           Case 3
              �ƥ�d��Ʈw = "DEF=3=DEF=3"
        End Select
    Case "��4"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-def4"
           Case 3
              �ƥ�d��Ʈw = "DEF=4=DEF=4"
        End Select
    Case "��5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-def5"
           Case 3
              �ƥ�d��Ʈw = "DEF=5=DEF=5"
        End Select
    Case "��7"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-def7"
           Case 3
              �ƥ�d��Ʈw = "DEF=7=DEF=7"
        End Select
    Case "�S1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-spe1"
           Case 3
              �ƥ�d��Ʈw = "SPE=1=SPE=1"
        End Select
    Case "�S2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-spe2"
           Case 3
              �ƥ�d��Ʈw = "SPE=2=SPE=2"
        End Select
    Case "�S3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 7
           Case 2
              �ƥ�d��Ʈw = "7-spe3"
           Case 3
              �ƥ�d��Ʈw = "SPE=3=SPE=3"
        End Select
    Case "�S4"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 7
           Case 2
              �ƥ�d��Ʈw = "7-spe4"
           Case 3
              �ƥ�d��Ʈw = "SPE=4=SPE=4"
        End Select
    Case "�S5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 7
           Case 2
              �ƥ�d��Ʈw = "7-spe5"
           Case 3
              �ƥ�d��Ʈw = "SPE=5=SPE=5"
        End Select
    Case "��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-mov1"
           Case 3
              �ƥ�d��Ʈw = "MOV=1=MOV=1"
        End Select
    Case "��2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 4
           Case 2
              �ƥ�d��Ʈw = "4-mov2"
           Case 3
              �ƥ�d��Ʈw = "MOV=2=MOV=2"
        End Select
    Case "��3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 4
           Case 2
              �ƥ�d��Ʈw = "4-mov3"
           Case 3
              �ƥ�d��Ʈw = "MOV=3=MOV=3"
        End Select
    Case "��4"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 4
           Case 2
              �ƥ�d��Ʈw = "4-mov4"
           Case 3
              �ƥ�d��Ʈw = "MOV=4=MOV=4"
        End Select
    Case "��5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 4
           Case 2
              �ƥ�d��Ʈw = "4-mov5"
           Case 3
              �ƥ�d��Ʈw = "MOV=5=MOV=5"
        End Select
    Case "���|1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-draw1"
           Case 3
              �ƥ�d��Ʈw = "DRAW=1=DRAW=1"
        End Select
    Case "���|2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 5
           Case 2
              �ƥ�d��Ʈw = "5-draw2"
           Case 3
              �ƥ�d��Ʈw = "DRAW=2=DRAW=2"
        End Select
    Case "���|3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 5
           Case 2
              �ƥ�d��Ʈw = "5-draw3"
           Case 3
              �ƥ�d��Ʈw = "DRAW=3=DRAW=3"
        End Select
    Case "���|4"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 5
           Case 2
              �ƥ�d��Ʈw = "5-draw4"
           Case 3
              �ƥ�d��Ʈw = "DRAW=4=DRAW=4"
        End Select
    Case "���|5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 5
           Case 2
              �ƥ�d��Ʈw = "5-draw5"
           Case 3
              �ƥ�d��Ʈw = "DRAW=5=DRAW=5"
        End Select
    Case "�A�G�N1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 6
           Case 2
              �ƥ�d��Ʈw = "6-brk1"
           Case 3
              �ƥ�d��Ʈw = "BRK=1=BRK=1"
        End Select
    Case "�A�G�N2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 6
           Case 2
              �ƥ�d��Ʈw = "6-brk2"
           Case 3
              �ƥ�d��Ʈw = "BRK=2=BRK=2"
        End Select
    Case "�A�G�N3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 6
           Case 2
              �ƥ�d��Ʈw = "6-brk3"
           Case 3
              �ƥ�d��Ʈw = "BRK=3=BRK=3"
        End Select
    Case "�A�G�N5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 6
           Case 2
              �ƥ�d��Ʈw = "6-brk5"
           Case 3
              �ƥ�d��Ʈw = "BRK=5=BRK=5"
        End Select
    Case "HP�^�_1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-hp1"
           Case 3
              �ƥ�d��Ʈw = "HPL=1=HPL=1"
        End Select
    Case "HP�^�_2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-hp2"
           Case 3
              �ƥ�d��Ʈw = "HPL=2=HPL=2"
        End Select
    Case "HP�^�_3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-hp3"
           Case 3
              �ƥ�d��Ʈw = "HPL=3=HPL=3"
        End Select
    Case "�C3/�j1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks3g1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=3=ATK-�j=1"
        End Select
    Case "�C4/�j2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks4g2"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=4=ATK-�j=2"
        End Select
    Case "�C5/�j3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks5g3"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=5=ATK-�j=3"
        End Select
    Case "�j3/�C1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atkg3s1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=3=ATK-�C=1"
        End Select
    Case "�j4/�C2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg4s2"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=4=ATK-�C=2"
        End Select
    Case "�j5/�C3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg5s3"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=5=ATK-�C=3"
        End Select
    Case "��3/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-def3mov1"
           Case 3
              �ƥ�d��Ʈw = "DEF=3=MOV=1"
        End Select
    Case "��4/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-def4mov1"
           Case 3
              �ƥ�d��Ʈw = "DEF=4=MOV=1"
        End Select
    Case "��5/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 3
           Case 2
              �ƥ�d��Ʈw = "3-def5mov1"
           Case 3
              �ƥ�d��Ʈw = "DEF=5=MOV=1"
        End Select
    Case "�S1/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-spe1def1"
           Case 3
              �ƥ�d��Ʈw = "SPE=1=DEF=1"
        End Select
    Case "�S2/��2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 7
           Case 2
              �ƥ�d��Ʈw = "7-spe2def2"
           Case 3
              �ƥ�d��Ʈw = "SPE=2=DEF=2"
        End Select
    Case "�S3/��3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 7
           Case 2
              �ƥ�d��Ʈw = "7-spe3def3"
           Case 3
              �ƥ�d��Ʈw = "SPE=3=DEF=3"
        End Select
    Case "�C3/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks3mov1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=3=MOV=1"
        End Select
    Case "�C4/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks4mov1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=4=MOV=1"
        End Select
    Case "�C5/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 1
           Case 2
              �ƥ�d��Ʈw = "1-atks5mov1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=5=MOV=1"
        End Select
    Case "�j3/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atkg3mov1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=3=MOV=1"
        End Select
    Case "�j4/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg4mov1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=4=MOV=1"
        End Select
    Case "�j5/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 2
           Case 2
              �ƥ�d��Ʈw = "2-atkg5mov1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=5=MOV=1"
        End Select
    Case "�C3/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks3def1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=3=DEF=1"
        End Select
    Case "�j3/��1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atkg3def1"
           Case 3
              �ƥ�d��Ʈw = "ATK-�j=3=DEF=1"
        End Select
    Case "��1/�S1"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-mov1spe1"
           Case 3
              �ƥ�d��Ʈw = "MOV=1=SPE=1"
        End Select
    Case "��2/�S2"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 4
           Case 2
              �ƥ�d��Ʈw = "4-mov2spe2"
           Case 3
              �ƥ�d��Ʈw = "MOV=2=SPE=2"
        End Select
    Case "��3/�S3"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 4
           Case 2
              �ƥ�d��Ʈw = "4-mov3spe3"
           Case 3
              �ƥ�d��Ʈw = "MOV=3=SPE=3"
        End Select
    Case "�C5/�j5"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-atks5g5"
           Case 3
              �ƥ�d��Ʈw = "ATK-�C=5=ATK-�j=5"
        End Select
    Case "�t��"
       Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 0
           Case 2
              �ƥ�d��Ʈw = "0-hpw"
           Case 3
              �ƥ�d��Ʈw = "HPW=1=HPW=1"
        End Select
    Case Else
        Select Case �ؼ�
           Case 1
              �ƥ�d��Ʈw = 99
        End Select
End Select
End Function
Sub �԰��t�Ϊ��Ū�J�{��()
Dim �Ȯɼ�(2) As Integer '���s�y�мȮ��ܼ�
Dim i As Integer, j As Integer, ckl As Integer, mm As Integer, w As Integer, n As Integer '�Ȯ��ܼ�
'------------
goidefus = 0
movecp = 2
turnpageonin = 0
trend�Ȯ��ܼ� = 0
FormMainMode.PEAFInterface.MessageClear
FormMainMode.PEAFInterface.BnOKStopListen
FormMainMode.PEAFInterface.BnOKVisable False
'----------------------------�g�J�ޯ���y��
atkinghelpxy(1, 1, 1) = 2520
atkinghelpxy(1, 2, 1) = 4730
atkinghelpxy(1, 3, 1) = 6930
atkinghelpxy(1, 4, 1) = 9140

atkinghelpxy(2, 1, 1) = 7560
atkinghelpxy(2, 2, 1) = 5400
atkinghelpxy(2, 3, 1) = 3240
atkinghelpxy(2, 4, 1) = 1080
For i = 1 To 4
   atkinghelpxy(1, i, 2) = 3480
   atkinghelpxy(2, i, 2) = 960
Next
'-----------------------
For ckl = 1 To ���εP����d�����j������(1)
    FormMainMode.card(ckl).Visible = False
    FormMainMode.card(ckl).CardEnabledType = True
Next
'-------�H�U�O�]�p�������
If Formsetting.checktest.Value = 0 Then
    FormMainMode.pageusglead.Visible = False
    FormMainMode.pageusqlead.Visible = False
    FormMainMode.pagecomglead.Visible = False
    FormMainMode.pagecomqlead.Visible = False
    FormMainMode.Command1.Visible = False
    FormMainMode.Command2.Visible = False
    FormMainMode.�v�l�]�w.Visible = False
End If

If Formsetting.checktestpersondown.Value = 1 Then
    FormMainMode.�v�l�]�w.Visible = True
End If

If Formsetting.�j�ð��ﶵ.Value = 0 Then
    �Ȯɼ�(1) = 8400
    �Ȯɼ�(2) = 7200
Else
   �Ȯɼ�(1) = 9360
   �Ȯɼ�(2) = 6720
End If
FormMainMode.cn1.Top = �Ȯɼ�(1)
FormMainMode.cnmove.Top = �Ȯɼ�(1)
FormMainMode.cnmove2.Top = �Ȯɼ�(1)
FormMainMode.cn2.Top = �Ȯɼ�(1)
FormMainMode.cn22.Top = �Ȯɼ�(1)
FormMainMode.cn3.Top = �Ȯɼ�(1)
FormMainMode.cn32.Top = �Ȯɼ�(1)
FormMainMode.cn4.Top = �Ȯɼ�(1)
FormMainMode.cn1.Left = �Ȯɼ�(2)
FormMainMode.cnmove.Left = �Ȯɼ�(2)
FormMainMode.cnmove2.Left = �Ȯɼ�(2)
FormMainMode.cn2.Left = �Ȯɼ�(2)
FormMainMode.cn22.Left = �Ȯɼ�(2)
FormMainMode.cn3.Left = �Ȯɼ�(2)
FormMainMode.cn32.Left = �Ȯɼ�(2)
FormMainMode.cn4.Left = �Ȯɼ�(2)
'=====================
FormMainMode.��ܦC1.goi1��� = False
FormMainMode.��ܦC1.goi2��� = False
FormMainMode.cn2.Visible = False
FormMainMode.cnmove.Visible = False
FormMainMode.cnmove2.Visible = False
FormMainMode.cn22.Visible = False
FormMainMode.cn3.Visible = False
FormMainMode.cn32.Visible = False
FormMainMode.cn4.Visible = False
'===========================�t����ܤ����]�w
If �t����ܬɭ������� = 1 Then
    FormMainMode.PEAFInterface.Passive_������� = False
    FormMainMode.cardpagejpg.Visible = True
    FormMainMode.pageul.Visible = True
    FormMainMode.draw1.Visible = True
    FormMainMode.move1.Visible = True
    FormMainMode.move3.Visible = True
    FormMainMode.move4.Visible = True
    FormMainMode.draw2.Visible = False
    FormMainMode.atkdef1.Visible = False
    FormMainMode.atkdef2.Visible = False
    FormMainMode.move2.Visible = False
Else
    FormMainMode.PEAFInterface.Passive_������� = True
    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\stageblack.gif"
    FormMainMode.cardpagejpg.Visible = False
    FormMainMode.pageul.Visible = False
    FormMainMode.draw1.Visible = False
    FormMainMode.move1.Visible = False
    FormMainMode.move3.Visible = False
    FormMainMode.move4.Visible = False
    FormMainMode.draw2.Visible = False
    FormMainMode.atkdef1.Visible = False
    FormMainMode.atkdef2.Visible = False
    FormMainMode.move2.Visible = False
End If
FormMainMode.��ܦC1.��ܦC�Ϥ� = app_path & "gif\system\DRAW.png"
'-----------�H�U���P�ժ�l��
If Formsetting.�j�ð��ﶵ.Value = 1 Then
    For mm = 1 To 3
       �P�`���q��(mm) = Val(Formsetting.�j�ð��Ҧ��ﶵ_�P��.Text)
       If �P�`���q��(mm) < 1 Then �P�`���q��(mm) = 1
    Next
ElseIf Formsetting.�D�ԼҦ��ﶵ.Value = 1 Then
   For mm = 2 To 3
       �P�`���q��(mm) = 5 + Val(Formsetting.�D�ԼҦ��ﶵ_�P��.Text)
   Next
    �P�`���q��(1) = 5
Else
    For mm = 1 To 3
       �P�`���q��(mm) = 5
    Next
End If

'=============���եγ~
'�P�`���q��(1) = 57
'�P�`���q��(2) = 0
'�P�`���q��(3) = 57
'==================�P�]�w(���])
For mm = 1 To ���εP����d�����j������(1)
   pagecardnum(mm, 6) = 4
Next
�԰��t����.���εP���ϥ��ˬd
BattleCardNum = ���εP�U�P����������(0, 2)
�԰��t����.����ʧ@_�t���`�d�P�i�Ƨ�s
���q���A�� = 0
FormMainMode.pageusqlead.Caption = 0
FormMainMode.pageusglead.Caption = 0
FormMainMode.pagecomqlead.Caption = 0
FormMainMode.pagecomglead.Caption = 0
'=======�H�U������H���]�w(�`��])
FormMainMode.PEAFpersoncardus(1).Left = 0
FormMainMode.PEAFpersoncardus(1).Visible = False
FormMainMode.cardus(1).Left = 0
FormMainMode.cardus(1).Top = 6240
FormMainMode.cardus(1).ZOrder
FormMainMode.cardus(1).Visible = True
FormMainMode.PEAFpersoncardcom(1).Left = 0
�԰��t����.PersonCardShowOnMode(1, 1) = True
�԰��t����.PersonCardShowOnMode(2, 1) = True
FormMainMode.PEAFpersoncardus(1).ShowOnMode = True
FormMainMode.PEAFpersoncardcom(1).ShowOnMode = True
'=======
For i = 2 To 3
   If ����H����ԤH��(1, 1) >= i Then
       FormMainMode.PEAFpersoncardus(i).ShowOnMode = True
       FormMainMode.PEAFpersoncardus(i).Left = 2520 * (i - 1)
       FormMainMode.PEAFpersoncardus(i).Visible = True
       FormMainMode.cardus(i).Visible = False
       �԰��t����.PersonCardShowOnMode(1, i) = True
   Else
       FormMainMode.PEAFpersoncardus(i).Visible = False
       FormMainMode.cardus(i).CardMain_����HP = 0
   End If
   If ����H����ԤH��(2, 1) >= i Then
       �԰��t����.PersonCardShowOnMode(2, i) = False
       FormMainMode.PEAFpersoncardcom(i).ShowOnMode = False
       FormMainMode.PEAFpersoncardcom(i).Left = 2520 * (i - 1)
       FormMainMode.PEAFpersoncardcom(i).Visible = True
   Else
       FormMainMode.PEAFpersoncardcom(i).Visible = False
       FormMainMode.cardcom(i).CardMain_����HP = 0
   End If
Next
If ����H����ԤH��(1, 1) > 1 Or ����H����ԤH��(2, 1) > 1 Then
   FormMainMode.��ܦC1.�H���԰��H�� = 3
Else
   FormMainMode.��ܦC1.�H���԰��H�� = 1
End If
For w = 1 To 3
   ����ݾ��H��������(1, w) = w
   ����ݾ��H��������(2, w) = w
Next
'=======�H�U������H���]�w(�ϥΪ�)
����H����ԤH��(1, 2) = 1
For w = 1 To ����H����ԤH��(1, 1)
    FormMainMode.cardus(w).CardMain_����Ϥ� = VBEPerson(1, w, 1, 5, 5)
    atkus(w) = Val(VBEPerson(1, w, 1, 3, 2))
    defus(w) = Val(VBEPerson(1, w, 1, 3, 3))
    liveus(w) = Val(VBEPerson(1, w, 1, 3, 1))
    uslevel(w) = Val(VBEPerson(1, w, 1, 2, 2))
    nameus(w) = VBEPerson(1, w, 1, 1, 1)
    FormMainMode.cardus(w).CardMain_����HP = liveus(w)
    FormMainMode.cardus(w).CardMain_����ATK = atkus(w)
    FormMainMode.cardus(w).CardMain_����DEF = defus(w)
    liveusmax(w) = liveus(w)
    liveus41(w) = liveusmax(w) \ 3
    '=================
    �԰��t����.��������p�d��T�g�J 1, w
    '=================
    FormMainMode.cardus(w).CardMain_����HPMAX = liveusmax(w)
    FormMainMode.cardus(w).CardMain_�O�_���s�˦���T = CBool(Val(VBEPerson(1, w, 1, 3, 5)) = 1)
    '=================
    FormMainMode.cardus(w).���`���A�����]
    FormMainMode.cardus(w).CardBack�����]
    �԰��t����.�ޯ໡�����J_�H���d���I��_�ϥΪ� w
Next
FormMainMode.bloodnumus1.Caption = liveus(1)
FormMainMode.bloodnumus2.Caption = liveusmax(1)
FormMainMode.personusminijpg.�p�H�����] = True
FormMainMode.personusminijpg.�p�H���Ϥ� = VBEPerson(1, 1, 1, 5, 1)
FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = VBEPerson(1, 1, 1, 5, 2)
FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = VBEPerson(1, 1, 1, 5, 4)
FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width
FormMainMode.personusminijpg.�p�H���v�lLeft = Val(VBEPerson(1, 1, 2, 1, 5))
FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(VBEPerson(1, 1, 2, 1, 6))
FormMainMode.personusminijpg.�p�H���v������ = False
'=======�H�U������H���]�w(�q��)
����H����ԤH��(2, 2) = 1
For w = 1 To ����H����ԤH��(2, 1)
    FormMainMode.cardcom(w).CardMain_����Ϥ� = VBEPerson(2, w, 1, 5, 5)
    atkcom(w) = Val(VBEPerson(2, w, 1, 3, 2))
    defcom(w) = Val(VBEPerson(2, w, 1, 3, 3))
    livecom(w) = Val(VBEPerson(2, w, 1, 3, 1))
    comlevel(w) = Val(VBEPerson(2, w, 1, 2, 2))
    namecom(w) = VBEPerson(2, w, 1, 1, 1)
    livecommax(w) = livecom(w)
    '=================
    �԰��t����.��������p�d��T�g�J 2, w
    '=================
    livecom41(w) = livecommax(w) \ 3
    FormMainMode.cardcom(w).CardMain_����HP = livecom(w)
    FormMainMode.cardcom(w).CardMain_����HPMAX = livecommax(w)
    FormMainMode.cardcom(w).CardMain_����ATK = atkcom(w)
    FormMainMode.cardcom(w).CardMain_����DEF = defcom(w)
    FormMainMode.cardcom(w).CardMain_�O�_���s�˦���T = CBool(Val(VBEPerson(2, w, 1, 3, 5)) = 1)
    '=================
    FormMainMode.cardcom(w).���`���A�����]
    FormMainMode.cardcom(w).CardBack�����]
    �԰��t����.�ޯ໡�����J_�H���d���I��_�q�� w
Next
FormMainMode.bloodnumcom1.Caption = livecom(1)
FormMainMode.bloodnumcom2.Caption = livecommax(1)
FormMainMode.personcomminijpg.�p�H�����] = True
FormMainMode.personcomminijpg.�p�H���Ϥ� = VBEPerson(2, 1, 1, 5, 1)
FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = VBEPerson(2, 1, 1, 5, 2)
FormMainMode.��ܦC1.�q����p�H���Ϥ� = VBEPerson(2, 1, 1, 5, 4)
FormMainMode.��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
FormMainMode.personcomminijpg.�p�H���v�lLeft = Val(VBEPerson(2, 1, 2, 1, 5))
FormMainMode.personcomminijpg.�p�H���v�ltop�t = Val(VBEPerson(2, 1, 2, 1, 6))
FormMainMode.personcomminijpg.�p�H���v������ = True
'==================�g�J�ޯ���y��
atkinghelpxy(1, 1, 1) = 2520
atkinghelpxy(1, 2, 1) = 4730
atkinghelpxy(1, 3, 1) = 6930
atkinghelpxy(1, 4, 1) = 9140

atkinghelpxy(2, 1, 1) = 6840
atkinghelpxy(2, 2, 1) = 4560
atkinghelpxy(2, 3, 1) = 2280
atkinghelpxy(2, 4, 1) = 0
For i = 1 To 4
    atkinghelpxy(1, i, 2) = 3000
    atkinghelpxy(2, i, 2) = 960
Next
'===========================�ޯ໡�����J
�԰��t����.�ޯ໡�����J_�ϥΪ�
�԰��t����.�ޯ໡�����J_�q��
FormMainMode.PEAFInterface.Passive_�ޯ�@������] 1
FormMainMode.PEAFInterface.Passive_�ޯ�@������] 2
For n = 5 To 8
    If VBEPerson(1, 1, 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ����� n - 4
    Else
       FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ�W�� VBEPerson(1, 1, 3, n, 1), n - 4
       FormMainMode.PEAFInterface.Passive_�ϥΪ�_�ޯ���� n - 4
    End If
    '=============
    If VBEPerson(2, 1, 3, n, 1) = "" Then
       FormMainMode.PEAFInterface.Passive_�q��_�ޯ����� n - 4
    Else
       FormMainMode.PEAFInterface.Passive_�q��_�ޯ�W�� VBEPerson(2, 1, 3, n, 1), n - 4
       FormMainMode.PEAFInterface.Passive_�q��_�ޯ���� n - 4
    End If
Next
FormMainMode.PEAFatkinghelpc.Visible = False
'=====================�H�U�O�ޯ����C�����
For i = 1 To 2
    For j = 1 To 4
        FormMainMode.PEAFInterface.ActiveSkillLight i, j, False
    Next
Next
'==================����p�H����ø���w�ζZ�����w
����ʧ@_�Z���ܧ� movecp, False
'================���ԼҦ��]�w
If Formsetting.chkpersonvsmode.Value = 1 Then
    For i = 2 To 3
        FormMainMode.PEAFpersoncardcom(i).ShowOnMode = False
        FormMainMode.cardcom(i).CardMain_����HP = -99
        FormMainMode.cardcom(i).CardMain_����ATK = -99
        FormMainMode.cardcom(i).CardMain_����DEF = -99
        FormMainMode.cardcom(i).CardMain_����Ϥ� = app_path & "gif\system\personunknown.jpg"
        FormMainMode.cardcom(i).CardBack�����]
    Next
End If
'--------------------------�p��Z�����(HP���)
�Z�����(1, 1, 1) = 5295 \ liveusmax(1)
�Z�����(1, 2, 1) = (11340 - 6060) \ livecommax(1)
'==============��q���J�ʵe�]�w
FormMainMode.bloodlineout1.Width = 1
FormMainMode.bloodlineout2.Left = 11325
Erase ��q�p�ƾ��ʵe�Ȯ��ܼ�
��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 1) = 106
��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 1) = 106
'==============�ɶ��b�C��]�w
�԰��t����.�ɶ��b_���]
'===============���]���貧�`���A�]�w
For i = 1 To 2
    For j = 1 To UBound(�H�����`���A�C��, 2)
        Set �H�����`���A�C��(i, j) = New Collection
    Next
Next
'==================
BattleTurn = 1
FormMainMode.PEAFInterface.turn = BattleTurn
FormMainMode.PEAFAnimateInterface.MusicPlayerObj = FormMainMode.cMusicPlayer(5)
FormMainMode.PEAFAnimateInterface.ImageMaskUse = Formsetting.chkImageMaskUse.Value
End Sub
Sub �ۥѾ԰��Ҧ��]�w���Ū�J�{��()
Dim i, j As Integer
���ϥΪ̨ƥ� = True
���q���ƥ� = True
FormMainMode.cdgpersonus.Filter = "UnlightVBE �d���H����T��(*.uleci)|*.uleci"
�d���H����T�ɮ�Ū�����Ѭ����� = ""
'======================�M���H�������O���ܼ�
Erase VBEPerson
Erase VBEVSSAtkingStr
totpersonnumber = 0
�`�@�H���W�� = ""
�`�@�H���ɮצW = ""
'======================
For i = 1 To 3
    FormMainMode.personnameus(i).Clear
    FormMainMode.personnamecom(i).Clear
    FormMainMode.personlevelus(i).Clear
    FormMainMode.personlevelcom(i).Clear
Next
�@��t����.�d���H����T���J_�j�M�ɮ�
'===============�վ�w�]
If FormMainMode.personnameus(1).ListCount > 0 Then
    For i = 1 To 3
       FormMainMode.personnameus(i).ListIndex = 0
       FormMainMode.personnamecom(i).ListIndex = 0
    Next
End If
FormMainMode.opnpersonvs(2).Value = True
FormMainMode.cMusicPlayer(0).MusicPlay
FormMainMode.personreadifus.Visible = False
End Sub
Sub �C����lŪ�J�{��()
    Dim i As Integer
    '=====�H�U�O�I�����֤�SE��l�]�w
    For i = FormMainMode.cMusicPlayer.UBound + 1 To 11
        Load FormMainMode.cMusicPlayer(i)
    Next
    FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulbgm03.mp3"
    FormMainMode.cMusicPlayer(0).Volume = 50
    FormMainMode.cMusicPlayer(0).IsLoop = True
    �@��t����.���Ī�l�]�w
    For i = 1 To FormMainMode.cMusicPlayer.UBound
          FormMainMode.cMusicPlayer(i).Volume = 45
    Next
End Sub
Sub �D���_PEStartForm���()
FormMainMode.PEStartForm.Left = 0
FormMainMode.PEStartForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
FormMainMode.PEStartForm.Visible = True
FormMainMode.PEStartForm.ZOrder
'====================
FormMainMode.PEStext1.Visible = False
tr1num = 0
FormMainMode.tr1.Enabled = True
End Sub
Sub �D���_PEGameFreeModeSettingForm���()
FormMainMode.PEGameFreeModeSettingForm.Left = 0
FormMainMode.PEGameFreeModeSettingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
FormMainMode.PEGameFreeModeSettingForm.Visible = True
FormMainMode.PEGameFreeModeSettingForm.Enabled = True
FormMainMode.PEGameFreeModeSettingForm.ZOrder
'===================�{�ɳq�����
If �q�����O�_�w�X�{ = False Then
    �q�����O�_�w�X�{ = True
    If �d���H����T�ɮ�Ū�����Ѭ����� <> "" Then
        �@��t����.�q�������� 1
    End If
End If
If Formsetting.chkautocontinuemode.Value = 1 Then
    FormMainMode.PEGFbnstart_Click
End If
End Sub
Sub �D���_PEAttackingForm���()
FormMainMode.PEAttackingForm.Left = 0
FormMainMode.PEAttackingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
FormMainMode.PEAttackingForm.Visible = True
FormMainMode.PEAttackingForm.ZOrder
End Sub
Sub �D���_PEAttackingStartForm���()
FormMainMode.PEAttackingStartForm.Left = 0
FormMainMode.PEAttackingStartForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
FormMainMode.PEASpersontalk.Visible = False
FormMainMode.PEAttackingStartForm.Visible = True
FormMainMode.PEAttackingStartForm.ZOrder
'=============
FormMainMode.PEASpersontalk.Visible = False
End Sub
Sub �D���_PEAttackingEndingForm���()
FormMainMode.PEAttackingEndingForm.Left = 0
FormMainMode.PEAttackingEndingForm.Top = 0
FormMainMode.Width = 11430
FormMainMode.Height = 10325
FormMainMode.PEAttackingEndingForm.Visible = True
FormMainMode.PEAttackingEndingForm.ZOrder
'================
FormMainMode.bnreturnt.Visible = False
FormMainMode.bnreturn.Visible = False
FormMainMode.bn.Visible = False
FormMainMode.bnt.Visible = False
End Sub
Sub �԰��t�ζ}�l���Ū�J�{��()
Dim i As Integer '�Ȯ��ܼ�
'======================
FormMainMode.downjpg.Top = Val(FormMainMode.Height)
FormMainMode.upjpg_2.�j�H���Ϥ� = App.Path & "\gif\system\startupjpg.png"
FormMainMode.upjpg_2.Top = -Val(FormMainMode.upjpg.Height)
For i = 1 To 3
   FormMainMode.PEAScardus(i).Top = -Val(FormMainMode.PEAScardus(i).Height)
   FormMainMode.PEAScardcom(i).Top = -Val(FormMainMode.PEAScardcom(i).Height)
Next
FormMainMode.�j�H���ι�_�ϥΪ�.�j�H���Ϥ� = VBEPerson(1, 1, 1, 5, 3)
FormMainMode.�j�H���ι�_�ϥΪ�.�j�H���v������ = False
FormMainMode.�j�H���ι�_�ϥΪ�.Top = 8400 - FormMainMode.�j�H���ι�_�ϥΪ�.�j�H���Ϥ�height
FormMainMode.�j�H���ι�_�ϥΪ�.Width = FormMainMode.�j�H���ι�_�ϥΪ�.�j�H���Ϥ�width
FormMainMode.�j�H���ι�_�ϥΪ�.Left = -FormMainMode.�j�H���ι�_�ϥΪ�.�j�H���Ϥ�width
FormMainMode.�j�H���ι�_�q��.�j�H���Ϥ� = VBEPerson(2, 1, 1, 5, 3)
FormMainMode.�j�H���ι�_�q��.�j�H���v������ = True
FormMainMode.�j�H���ι�_�q��.Top = 8400 - FormMainMode.�j�H���ι�_�q��.�j�H���Ϥ�height
FormMainMode.�j�H���ι�_�q��.Width = FormMainMode.�j�H���ι�_�q��.�j�H���Ϥ�width
FormMainMode.�j�H���ι�_�q��.Left = FormMainMode.ScaleWidth
st = 0
sq = 0
FormMainMode.start1.Enabled = True
End Sub
Sub �ۥѾ԰��Ҧ��]�w���򥻳]�w�{��()
Dim i As Integer '�Ȯ��ܼ�
Formsetting.��Ԧa�Ͽ��.ListIndex = 0
Formsetting.BGM���.ListIndex = 0
Formsetting.comboeventcarrdus.ListIndex = 2
Formsetting.comboeventcarrdcom.ListIndex = 2
Formsetting.cbsimilarlevel.ListIndex = 4
For i = 1 To 18
    Formsetting.personus(i).ListIndex = 0
    Formsetting.personcom(i).ListIndex = 0
Next
Formsetting.lopnmusictext.Visible = False
Formsetting.lopnmapjpgtext.Visible = False
Formsetting.ckendturnnum.Text = 18
Formsetting.t1.Tab = 0
Formsetting.chkusenewai.Value = 1
Formsetting.chkusenewpage.Value = 1
Formsetting.chkusenewinterface.Value = 1
Formsetting.chkpersonvsmode.Value = 1
Formsetting.ckendturn.Value = 1
End Sub
Sub �M���԰��t�ΩҦ��ܼƭ�()
Dim i As Integer, j As Integer '�Ȯ��ܼ�
'Erase atkingno '�ޯ�o�ʱƧǼȮɹϤ����|�x�s�ܼ�(�ޯ�o�ʶ���8~1,1.�Ϥ����|/2.(1)�ϥΪ�/(2)�q����/3.Left/4.Top(�y��)/5.�����e��(Width)/6.��������(Height)/7.�ޯ�s��/8.�ޯ���椤�ɱҰʭ�/9.�ޯ���椤���Ϥ��ˬd��/10.��2�i�Ϥ����|)
Erase goicheck   '����/���m�Ҧ��[��ƭ��ˬd�X
Erase liveus
Erase livecom
Erase liveusmax
Erase livecommax
Erase pageusleadmax  '�ϥΪ̵P���ǭp�ƪ�(0.��P/1.�X�P)
Erase pagecomleadmax   '�q���P���ǭp�ƪ�(0.��P/1.�X�P)
Erase pageqlead   '�X�P�p���ܼ�(1.�ϥΪ�/2.�q��)
Erase pageglead   '��P�p���ܼ�(1.�ϥΪ�/2.�q��)
movedsus = 0   '�ϥΪ̲��ʶ��q�M�w���ܼ�
turnpageonin = 0 '���q�O�_�i�X�P�ܼ�(�@��)
turnpageoninatking = 0 '���q�O�_�i�X�P�ܼ�(�ޯ�ϥ�)
goickus = 0 '�P�Ȥ@���ˬd�X
Erase atkingck   '�ޯඥ�q�ҰʽX(x.�H���ޯ�s��,1.�ޯ���涥�q/2.�ޯ�Ұ��ˬd��)
'Erase atkingckai   'AI�ޯඥ�q�ҰʽX(x.�H���ޯ�s��,1.�ޯ���涥�q/2.�ޯ�Ұ��ˬd��)
Erase atkingtrn  '�ޯ�p�ƾ��Ȯ��x�s�ܼ�(1.�ϥΪ�(�{)/2.�q��(�{)/3.�ϥΪ�(�ƥ�)/4.�q��(�ƥ�))
HP�ˬd�ܼ� = False 'HP�ˬd���q�O�_�w�ˬd�ܼ�
HP�ˬd���q�� = 0  'HP�ˬd���q�ܼ�(1.���ʶ��q��,2.����/���m���q�e,3.��/���m���q��)
Erase �Z�����  '�Z�����Ȯ��x�s���(1.HP���/2.�P����,1.�ϥΪ�/2.�q��,1.Left���/2.Top���)
Erase personminixy '�p�H���Ϥ��y�Ы��w���(1.�ϥΪ�/2.�q��,��n��,1.��Z��/2.���Z��/3.���Z��,1.Left/2.Top(�y��))
'Erase �H�����`���A��Ʈw '���`���A���(1.�ϥΪ�/2.�q��,��x�Ӳ��`���A,1.���A�ƭ�/2.���A�έp��(�Ѿl�^�X/�֭p)/3.���A�s��)
Erase ���`���A�ˬd�� '���`���A�ҰʽX(x.���`���A�s��,1.���A���涥�q/2.���A�Ұ��ˬd��)
�ޯ�ʵe��ܶ��q�� = 0 '�ޯ�ʵe�p�ƾ����q�X(1.����/���m���q-���q,2.���ʶ��q-���q/3.�o�P���q��B���ʶ��q�e/4.���ʶ��q��/5.�������q��/6.���m���q��/7.�^�X������)
Erase �������m��l�`�� '����/���m�Ҧ���l�ƶq���(1.�ϥΪ�(�`)/2.�q��(�`)/3.�ϥΪ�(��)/4.�q��(��))
Erase atkingpagetot  '�C���q�X�P�����μƭȲέp���(1.�ϥΪ�/2.�q��,1.�C/2.��/3.��/4.�S/5.�j)
Erase ��ƹs�ˬd��  '��e���q��l�ƶq�O�_���s�ˬd��(1.�ϥΪ�/2.�q��)
Erase pagecardnum '���εP���(��x�s��(1~70-���P/71~88-�ϥΪ̨ƥ�P/89~106-�q���ƥ�P),1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.(1)�ϥΪ�-(2)�q��/6.(1)��P-(2)�X�P-(3)�õP-(4)�P��/7.�X�P����/8.�Ϥ��s��/9.�ثeLeft(�y��)/10.�ثeTop(�y��)/11.(1)�q����X�P(��)-(2)�q���o�X�P(�~))
Erase �P�`���q�� '�P�֦��`���q��(1.�ϥΪ�/2.�q��/3.�`�p)
Erase �P���ʼȮ��ܼ� '�P���ʭp�ƾ��Ȯ��ܼ�(1.Left���/2.Top���/3.�P�i�s��)
Erase �ثe��  '�`�Ȯ��ܼ�
Erase �X�P���ǲέp�Ȯ��ܼ� '�X�P���ǲέp�`�Ȯɸ��(1.�ϥΪ̥X�P/2.�ϥΪ̤�P/3.�q���X�P/4.�q����P,��x����,1.�ثe�P�X�P����/2.�P�i�s��)
Erase �Z�����_���P�Ȯɼ�  '���P�ӧO�Z�����Ȯ��x�s�ܼ�(��x����,1.Left���/2.Top���/3.�P�i�s��)
���q���A�� = 0 '�C���q�}�l�������A�ˬd��(1.�}�l���q/2.�������q)
Erase �p�H���Y�����ʤ�V��   '�p�H���Y�����ʤ�V���A��(1.�ϥΪ�/2.�q��[1.�V��,2.�V�~])
Erase ��q�p�ƾ��ʵe�Ȯ��ܼ� '�}�l��l���q-��q�ʵe�p�ƾ��Ȯ��ܼ�(1.�ϥΪ̦��/2.�q�����,1.�C�����ʶq/2.�O�_�w����)
Erase �ɶ��b�C���ܤƬ����Ȯ��ܼ� '�ɶ��b�i���C���ܤƶ��q�����Ȯ��ܼ�(1~3(1)����ܤƶq(1(1).�ɶ��b(���~))/2.�ثe�֭p�q/3.�ثe�C��(R,G,B),4.(1)�ɶ��b(�~)���q��-(1)���ܬ�-(2)���ܶ�/2.�ثe�֭p�q/3.�ثe�C��(R))
Erase �}�l�d�����ʰʵe������   '�}�l�ɨC�i�d�����ʰʵe����������(1.�ϥΪ�/2.�q��,1~3.�d��/4.�ثe�ĴX�i)
Erase �洫��������Ȯ��ܼ� '�洫������������Ȯɼ�(1.�ϥΪ�/2.�q��/3.�O�_��U����/4.�洫���⧹���涥�q��)
Erase pageeventnum '�ƥ�d�ƦC�������(1.�ϥΪ�/2.�q��,1~18-�s��,1.�ƥ�d�W��/2.�ƥ�d�ɮצW��)
'�Y����ˮ`�� = 0 '�԰��t�Ϊ��-�Y���淾�q�Ȯ��ܼ�(2)���ܼƪ��
�԰��Ҧ��ӱѬ����� = 0 '�԰��t�η�e�ӱѬ����Ȯ��ܼ�(1.�ϥΪ̤�ӧQ/2.�ϥΪ̤�ѥ_/3.����)
�q���貾�ʶ��q��ܼ� = 0
�q����ƥ�d�O�_�X����ܼ� = False
�O�_�t�Τ��� = False
Erase �H���d���I���s��������  '�H���d���I���ޯ໡���H���s���Ȯ��ܼ�(1.(1).�ϥΪ�/(2).�q��,2.��n��)
Erase �Y���淾�q�Ȯ��ܼ� 'Form6���ȷ��q�Ȯ��ܼ�(1.�@�^�X������P�_(1.�e/2.��),2.��l���(�ϥΪ�)-�Y��ᦳ�Ķˮ`��,3.��l���(�q��)-�Y���ˮ`��H(1.�ϥΪ�/2.�q��),4.(1.�ϥΪ̥���/2.�q������))
Erase ���εP�U�P���������� '�U�������εP�P���������Ȯ��ܼ�(0.(1)�ثe�w�o�P�`�ƶq/(2)�ثe�����P�`�ƶq,1~31.(1)�ثe�w�ϥΤ��P��/(2)�ӵP����ϥΤ��`�ƶq)
Erase ���εP����d�����j������ '�԰��t�ι���P����������(1.�`�@�P��/2.���P�P��/3.�ϥΪ̨ƥ�d�̩��s��/4.�q���ƥ�d�̩��s��)
Erase �԰��Y�뤶���H����ø�ϸ��|������ '�԰��t���Y�뤶������H����ø�ϸ��|������(1.�ϥΪ̤�/2.�q����)
Erase �H����ڪ��A��Ʈw '�H����ڪ��A���
'===================
Erase �ƥ�d�O���Ȯɼ� '�ƥ�d�ϥά����Ȯ��ܼ�(0.(1)�`�@�����^�X��,1.�ϥΪ�/2.�q��,1.�`�@�ƭ�/2.�ثe�B�z�ƭ�/3.�ثe���q/4.�ƥ�d�P�s��/5.�ƥ����/6.�O�_�Ұ�)
''===================
vbecommadtotplay = 0
ReDim vbecommadnum(1 To 7, vbecommadtotplay)
ReDim vbecommadstr(1 To 3, vbecommadtotplay)
Erase Vss_PersonAtkingOffNum
Erase Vss_AtkingInformationRecordStr
ReDim VBEStageNum(0) As Integer
For i = 1 To 2
    For j = 1 To UBound(�H�����`���A�C��, 2)
        Set �H�����`���A�C��(i, j) = Nothing
    Next
Next
End Sub
Sub �M���԰��t�ζ}�l���]�w��()
Dim i As Integer, j As Integer
For i = 1 To 3
    FormMainMode.PEAScardus(i).Picture = LoadPicture()
    FormMainMode.PEASusbi1(i).Caption = 0
    FormMainMode.PEAScardcom(i).Picture = LoadPicture()
    FormMainMode.PEAScardcompi1(i).Caption = 0
    FormMainMode.cardusname(i).Visible = True
    FormMainMode.cardusspname(i).Visible = True
    FormMainMode.cardcomname(i).Visible = True
    FormMainMode.cardcomspname(i).Visible = True
    For j = 1 To 3
        FormMainMode.PEAFpersoncardcom(j).Left = 2520 * (j - 1)
        FormMainMode.PEAFpersoncardus(j).Left = 2520 * (j - 1)
    Next
    '================
    If i >= 2 Then
        Formchangeperson.card(i - 1).Visible = True
        Formchangeperson.bnok(i - 1).Visible = True
    End If
Next
End Sub
Sub �ۥѾ԰��Ҧ��]�w���U���]�wŪ�J�{��()
Dim i As Integer, ne As Integer, nd As Integer  '�Ȯ��ܼ�
'========������Ū�J
For i = 1 To 18
    ne = i Mod 6
    nd = i \ 6
    If ne = 0 Then ne = 6
    If i = 6 Or i = 12 Or i = 18 Then
       nd = (i \ 6) - 1
    End If
    Formsetting.persontgus(i).Caption = Mid(VBEPerson(1, nd + 1, 1, 3, 4), ne, 1)
    Formsetting.persontgcom(i).Caption = Mid(VBEPerson(2, nd + 1, 1, 3, 4), ne, 1)
Next
End Sub
Sub �q��������(ByVal num As Integer)
Dim pstr() As String
Dim k As Integer
Select Case num
    Case 1
        FormMessage.Label2 = "�j�p�j�z�n" & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & "�bŪ���Y�ǥd���H����T�ɮ׮ɵo�ͤF���~�A" & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & "�Фj�p�j��H�U�ɮ׶i��ӧO�ˬd�G" & Chr(10)
        FormMessage.Label2 = FormMessage.Label2 & Chr(10)
        pstr = Split(�d���H����T�ɮ�Ū�����Ѭ�����, "=")
        For k = 1 To UBound(pstr)
             FormMessage.Label2 = FormMessage.Label2 & pstr(k) & Chr(10)
        Next
        For k = 1 To 6
             FormMessage.Label2 = FormMessage.Label2 & Chr(10)
        Next
        FormMessage.Label2 = FormMessage.Label2 & "����" & Chr(10)
        FormMessage.Label2.Visible = True
        FormMessage.Text1.Visible = False
        FormMessage.Show 1
End Select
End Sub
Sub ���ļ���(ByVal num As Integer)
'Select Case num
'    Case 1
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse06.mp3"
'    Case 2
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse09.mp3"
'    Case 3
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse08.mp3"
'    Case 4
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse29.mp3"
'    Case 5
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse13.mp3"
'    Case 6
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse12.mp3"
'    Case 7
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse11.mp3"
'    Case 8
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse10_f.mp3"
'    Case 9
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse23.mp3"
'    Case 10
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse22.mp3"
'    Case 11
'        FormMainMode.cMusicPlayer(num).Filepath = app_path & "mp3\ulse01.mp3"
'End Select
FormMainMode.cMusicPlayer(num).MusicStop
FormMainMode.cMusicPlayer(num).MusicPlay
End Sub
Sub ���Ī�l�]�w()
FormMainMode.cMusicPlayer(1).Filepath = app_path & "mp3\ulse06.mp3"
FormMainMode.cMusicPlayer(2).Filepath = app_path & "mp3\ulse09.mp3"
FormMainMode.cMusicPlayer(3).Filepath = app_path & "mp3\ulse08.mp3"
FormMainMode.cMusicPlayer(4).Filepath = app_path & "mp3\ulse29.mp3"
FormMainMode.cMusicPlayer(5).Filepath = app_path & "mp3\ulse13.mp3"
FormMainMode.cMusicPlayer(6).Filepath = app_path & "mp3\ulse12.mp3"
FormMainMode.cMusicPlayer(7).Filepath = app_path & "mp3\ulse11.mp3"
FormMainMode.cMusicPlayer(8).Filepath = app_path & "mp3\ulse10_f.mp3"
FormMainMode.cMusicPlayer(9).Filepath = app_path & "mp3\ulse23.mp3"
FormMainMode.cMusicPlayer(10).Filepath = app_path & "mp3\ulse22.mp3"
FormMainMode.cMusicPlayer(11).Filepath = app_path & "mp3\ulse01.mp3"
End Sub
Sub ���}�C������(Cancel As Integer, UnloadMode As Integer)
Dim YesNo As VbMsgBoxResult
If UnloadMode = 0 Then
   YesNo = MsgBox("�T�w���}�C��?", 36, "UnlightVBE-�t�δ���")
   If YesNo = VbMsgBoxResult.vbYes Then
    End
   Else
    Cancel = 1
   End If
End If

End Sub
