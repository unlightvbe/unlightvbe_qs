VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWineobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private oVBEPersonVS() As Variant    'VBE�H���Τ@�ܼ�-VS��
Private oatkingpagetotVS() As Variant  '�C���q�X�P�����μƭȲέp���-VS��
Private oVBEPersonBuffVSF() As Variant    '���`���A���-VS-F��
Private oVBEPersonBuffVSS() As Variant    '���`���A���-VS-S��
Private oAtkingckVSS() As Variant  '�ޯ��T�@��-S��(�ޯ�ҰʽX)
Private oAtkingckVSF() As Variant    '�ޯ��T�@��-F��(�Ƶ��r��)
Private oVBEAtkingVSF() As Variant    'VBE>VS�����ܼƲΤ@���-F��
Private oVBEAtkingVSS() As Variant    'VBE>VS�����ܼƲΤ@���-S��
Private oVBEPageCardNumVS() As Variant    '���εP���-VS��
Private oVBEVSBuffNum() As Variant    '���`���A�M��-���`���A��2�Ӽƭ�-VS��
Private oVBEVSStageNum() As Variant    '���涥�q�t��-���涥�q�h�γ~�����ܼ�-VS��
Private oVBEActualStatusVS() As Variant    '�H����ڪ��A���-VS��
Public oPersonType As Variant  '�Ө���ثe���W���A
Public oNs As Variant    '���涥�q��

Public Sub AddInformation(ByVal str As Variant, ByRef Varray() As Variant)
    Select Case str
        Case "VBEAtkingVSF"
            oVBEAtkingVSF = Varray
        Case "VBEAtkingVSS"
            oVBEAtkingVSS = Varray
        Case "AtkingPagetotVS"
            oatkingpagetotVS = Varray
        Case "VBEPersonVS"
            oVBEPersonVS = Varray
        Case "VBEPageCardNumVS"
            oVBEPageCardNumVS = Varray
        Case "AtkingckVSS"
            oAtkingckVSS = Varray
        Case "AtkingckVSF"
            oAtkingckVSF = Varray
        Case "VBEPersonBuffVSF"
            oVBEPersonBuffVSF = Varray
        Case "VBEPersonBuffVSS"
            oVBEPersonBuffVSS = Varray
        Case "VBEActualStatusVS"
            oVBEActualStatusVS = Varray
        Case "VBEVSBuffNum"
            oVBEVSBuffNum = Varray
        Case "VBEVSStageNum"
            oVBEVSStageNum = Varray
    End Select
End Sub
Public Function GetArray(ByVal str As Variant)
    Select Case str
        Case "VBEAtkingVSF"
            GetArray = oVBEAtkingVSF
        Case "VBEAtkingVSS"
            GetArray = oVBEAtkingVSS
        Case "AtkingPagetotVS"
            GetArray = oatkingpagetotVS
        Case "VBEPersonVS"
            GetArray = oVBEPersonVS
        Case "VBEPageCardNumVS"
            GetArray = oVBEPageCardNumVS
        Case "AtkingckVSS"
            GetArray = oAtkingckVSS
        Case "AtkingckVSF"
            GetArray = oAtkingckVSF
        Case "VBEPersonBuffVSF"
            GetArray = oVBEPersonBuffVSF
        Case "VBEPersonBuffVSS"
            GetArray = oVBEPersonBuffVSS
        Case "VBEActualStatusVS"
            GetArray = oVBEActualStatusVS
        Case "VBEVSBuffNum"
            GetArray = oVBEVSBuffNum
        Case "VBEVSStageNum"
            GetArray = oVBEVSStageNum
    End Select
End Function