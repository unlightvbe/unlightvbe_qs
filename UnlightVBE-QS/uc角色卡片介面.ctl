VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl uc����d������ 
   Appearance      =   0  '����
   BackColor       =   &H00000000&
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ClipBehavior    =   0  '�L
   ClipControls    =   0   'False
   HitBehavior     =   2  '�ϥ�ø�ϰϰ�
   ScaleHeight     =   5250
   ScaleWidth      =   8160
   Begin VB.PictureBox PEAFcardbackpassive 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   5280
      Picture         =   "uc����d������.ctx":0000
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   60
      Top             =   0
      Width           =   2535
      Begin VB.Image cardactiveChickimage 
         Height          =   300
         Left            =   0
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label PEAFpersoncardback_passivetext 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   65
         Top             =   1740
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_passivetext 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   1280
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_passivetext 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   840
         Width           =   2295
      End
      Begin ImageX.aicAlphaImage PEAFcardbackpassiveBR 
         Height          =   390
         Index           =   4
         Left            =   90
         Top             =   1720
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   688
         Image           =   "uc����d������.ctx":7B89
         Opacity         =   50
         Props           =   5
      End
      Begin ImageX.aicAlphaImage PEAFcardbackpassiveBR 
         Height          =   390
         Index           =   3
         Left            =   90
         Top             =   1280
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   688
         Image           =   "uc����d������.ctx":8044
         Opacity         =   50
         Props           =   5
      End
      Begin ImageX.aicAlphaImage PEAFcardbackpassiveBR 
         Height          =   390
         Index           =   2
         Left            =   90
         Top             =   820
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   688
         Image           =   "uc����d������.ctx":84FF
         Opacity         =   50
         Props           =   5
      End
      Begin VB.Label PEAFpersoncardback_passivetext 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   380
         Width           =   2295
      End
      Begin ImageX.aicAlphaImage PEAFcardbackpassiveBR 
         Height          =   390
         Index           =   1
         Left            =   90
         Top             =   360
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   688
         Image           =   "uc����d������.ctx":89BA
         Opacity         =   50
         Props           =   5
      End
      Begin VB.Label PEAFpersoncardback_passivemain 
         BackStyle       =   0  '�z��
         Caption         =   "DEF+7�C���m���\�ɡA������P�ҶW�L�����m�P�Ȫ��ˮ`"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   8.25
            Charset         =   136
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   61
         Top             =   2280
         Width           =   2295
      End
   End
   Begin VB.PictureBox PEAFcardback 
      Appearance      =   0  '����
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   2640
      Picture         =   "uc����d������.ctx":8E75
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Image cardpassiveChickimage 
         Height          =   300
         Left            =   1250
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label PEAFpersoncardback_main 
         BackStyle       =   0  '�z��
         Caption         =   "DEF+7�C���m���\�ɡA������P�ҶW�L�����m�P�Ȫ��ˮ`"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   8.25
            Charset         =   136
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   59
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   57
         Top             =   1245
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '�z��
         Caption         =   "��K�g��"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   55
         Top             =   315
         Width           =   2295
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num4 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   54
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num4 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   53
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num4 
         Height          =   255
         Index           =   3
         Left            =   1635
         TabIndex        =   52
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num4 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   51
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num4 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   50
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num3 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   49
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num3 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   48
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num3 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   47
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num3 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   46
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num3 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   45
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num2 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   44
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num2 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   43
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num2 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   42
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num2 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   41
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num2 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   40
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num1 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   39
         Top             =   600
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num1 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   38
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num1 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   37
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num1 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   36
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range4 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   35
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range4 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   34
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range4 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   33
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range3 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   32
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range3 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   31
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range3 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   30
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range2 
         Height          =   255
         Index           =   3
         Left            =   885
         TabIndex        =   29
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range2 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   28
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range2 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   27
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range1 
         Height          =   255
         Index           =   3
         Left            =   885
         TabIndex        =   26
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range1 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   25
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_turn 
         Height          =   135
         Index           =   4
         Left            =   100
         TabIndex        =   24
         Top             =   1960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_turn 
         Height          =   135
         Index           =   3
         Left            =   100
         TabIndex        =   23
         Top             =   1530
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_turn 
         Height          =   135
         Index           =   2
         Left            =   100
         TabIndex        =   22
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_turn 
         Height          =   135
         Index           =   1
         Left            =   100
         TabIndex        =   21
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_range1 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   20
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc�d���I�� PEAFpersoncardback_num1 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   19
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   1
         Left            =   120
         Top             =   340
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "uc����d������.ctx":E8AA
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   2
         Left            =   120
         Top             =   800
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "uc����d������.ctx":E97F
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   3
         Left            =   120
         Top             =   1280
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "uc����d������.ctx":EA54
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   4
         Left            =   120
         Top             =   1710
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "uc����d������.ctx":EB29
         Props           =   13
      End
   End
   Begin VB.PictureBox card 
      Appearance      =   0  '����
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin ImageX.aicAlphaImage PEAFcardusbackclick 
         Height          =   795
         Left            =   480
         Top             =   1320
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1402
         Image           =   "uc����d������.ctx":EBFE
         Props           =   13
      End
      Begin ImageX.aicAlphaImage cardback 
         Height          =   3600
         Left            =   0
         Top             =   0
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   6350
         Image           =   "uc����d������.ctx":11333
         Props           =   9
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   8
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   9
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   10
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   11
         Left            =   1440
         TabIndex        =   7
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   12
         Left            =   1440
         TabIndex        =   6
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   13
         Left            =   1440
         TabIndex        =   5
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc���`���A personspe 
         Height          =   375
         Index           =   14
         Left            =   1440
         TabIndex        =   4
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin VB.Label personlabeldef 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label personlabelatk 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label personlabelhp 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   555
         TabIndex        =   1
         Top             =   3240
         Width           =   375
      End
   End
End
Attribute VB_Name = "uc����d������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_cardmain_jpg As String
Dim m_cardmain_personhp As Integer, m_cardmain_personhpmax As Integer, m_cardmain_personhp41 As Integer
Dim m_cardmain_personatk As Integer
Dim m_cardmain_persondef As Integer
Dim m_cardback_activehelp(1 To 4) As String, m_cardback_passivehelp(1 To 4) As String
Dim m_cardback_activecheck As Integer, m_cardback_passivecheck As Integer, m_cardbackcheck As Integer
Dim m_cardmain_isnewtype As Boolean
Public Sub ��ﲧ�`���A���(ByVal buffnum As Integer, ByVal ImagePath As String, ByVal num As Integer, ByVal tot As Integer, ByVal isVisible As Boolean)
If buffnum >= 1 And buffnum <= 14 Then
    If isVisible = False Then
        personspe(buffnum).Visible = False
    Else
        personspe(buffnum).person_num = num
        personspe(buffnum).person_turn = tot
        personspe(buffnum).���`���A�Ϥ� = ImagePath
        personspe(buffnum).Visible = True
    End If
End If
End Sub
Public Sub ���`���A�����]()
For i = 1 To 14
    personspe(i).Visible = False
Next
End Sub
Public Sub CardBack�����]()
Erase m_cardback_activehelp
m_cardback_activecheck = 0
m_cardback_passivecheck = 0
m_cardbackcheck = 0
For i = 1 To 4
      PEAFpersoncardback_turn(i).Visible = False
      PEAFpersoncardback_text(i).Visible = False
      PEAFpersoncardback_passivetext(i).Visible = False
      PEAFpersoncardback_main.Caption = ""
      PEAFpersoncardback_passivemain.Caption = ""
      '==========
      Select Case i
          Case 1
                 For k = 1 To 5
                     PEAFpersoncardback_num1(k).Visible = False
                 Next
               '================
                 For k = 1 To 3
                       PEAFpersoncardback_range1(k).�������O = 2
                       PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\system\cardback\CBrge.png"
                       PEAFpersoncardback_range1(k).���ؽs�� = 2
                 Next
               '================
          Case 2
                 For k = 1 To 5
                     PEAFpersoncardback_num2(k).Visible = False
                 Next
               '================
                 For k = 1 To 3
                       PEAFpersoncardback_range2(k).�������O = 2
                       PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\system\cardback\CBrge.png"
                       PEAFpersoncardback_range2(k).���ؽs�� = 2
                 Next
               '================
          Case 3
                 For k = 1 To 5
                     PEAFpersoncardback_num3(k).Visible = False
                 Next
               '================
                 For k = 1 To 3
                       PEAFpersoncardback_range3(k).�������O = 2
                       PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\system\cardback\CBrge.png"
                       PEAFpersoncardback_range3(k).���ؽs�� = 2
                 Next
               '================
          Case 4
                 For k = 1 To 5
                     PEAFpersoncardback_num4(k).Visible = False
                 Next
               '================
                 For k = 1 To 3
                       PEAFpersoncardback_range4(k).�������O = 2
                       PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\system\cardback\CBrge.png"
                       PEAFpersoncardback_range4(k).���ؽs�� = 2
                 Next
               '================
      End Select
Next
End Sub
Public Property Get CardMain_����Ϥ�() As String
   CardMain_����Ϥ� = m_cardmain_jpg
End Property
Public Property Let CardMain_����Ϥ�(ByVal New_CardMain_����Ϥ� As String)
   m_cardmain_jpg = New_CardMain_����Ϥ�
   PropertyChanged "CardMain_����Ϥ�"
   If Me.CardMain_����Ϥ� <> "" Then
       card.Picture = LoadPicture(Me.CardMain_����Ϥ�)
   End If
End Property
Public Property Get CardMain_����HP() As Integer
   CardMain_����HP = m_cardmain_personhp
End Property
Public Property Let CardMain_����HP(ByVal New_CardMain_����HP As Integer)
   m_cardmain_personhp = New_CardMain_����HP
   PropertyChanged "CardMain_����HP"
'   If Me.CardMain_����HP < 0 And Me.CardMain_����HP <> -99 Then Me.CardMain_����HP = 0
   '========================
   If Me.CardMain_����HP = -99 Then
       personlabelhp.Caption = "?"
   Else
       personlabelhp.Caption = Me.CardMain_����HP
   End If
   If Me.CardMain_����HP = Me.CardMain_����HPMAX Or Me.CardMain_����HP = -99 Then
        personlabelhp.ForeColor = RGB(255, 255, 255)
        personlabelhp.ForeColor = RGB(255, 255, 255)
        cardback.Opacity = 0
   ElseIf Me.CardMain_����HP < Me.CardMain_����HPMAX And Me.CardMain_����HP > m_cardmain_personhp41 Then
        personlabelhp.ForeColor = RGB(255, 255, 128)
        personlabelhp.ForeColor = RGB(255, 255, 128)
        cardback.Opacity = 0
   ElseIf Me.CardMain_����HP <= m_cardmain_personhp41 Then
        personlabelhp.ForeColor = RGB(255, 0, 0)
        personlabelhp.ForeColor = RGB(255, 0, 0)
        cardback.Opacity = 0
   End If
   If Me.CardMain_����HP = 0 Then
        cardback.Opacity = 100
        cardback.ZOrder
        PEAFcardusbackclick.Visible = False
   End If
End Property
Public Property Get CardMain_����HPMAX() As Integer
   CardMain_����HPMAX = m_cardmain_personhpmax
End Property
Public Property Let CardMain_����HPMAX(ByVal New_CardMain_����HPMAX As Integer)
   m_cardmain_personhpmax = New_CardMain_����HPMAX
   PropertyChanged "CardMain_����HPMAX"
   If Me.CardMain_����HPMAX < 0 Then Me.CardMain_����HPMAX = 0
   '==================
   If Me.CardMain_����HPMAX > 1 Then
       m_cardmain_personhp41 = Int(Me.CardMain_����HPMAX / 3 + 0.9)
   Else
       m_cardmain_personhp41 = 0
   End If
   '================================���HP���A
   Me.CardMain_����HP = m_cardmain_personhp
End Property
Public Property Get CardMain_����ATK() As Integer
   CardMain_����ATK = m_cardmain_personatk
End Property
Public Property Let CardMain_����ATK(ByVal New_CardMain_����ATK As Integer)
   m_cardmain_personatk = New_CardMain_����ATK
   PropertyChanged "CardMain_����ATK"
   If Me.CardMain_����ATK = -99 Then
       personlabelatk.Caption = "?"
   ElseIf Me.CardMain_����ATK < 0 Then
       Me.CardMain_����ATK = 0
       personlabelatk.Caption = Me.CardMain_����ATK
   Else
       personlabelatk.Caption = Me.CardMain_����ATK
   End If
End Property
Public Property Get CardMain_����DEF() As Integer
   CardMain_����DEF = m_cardmain_persondef
End Property
Public Property Let CardMain_����DEF(ByVal New_CardMain_����DEF As Integer)
   m_cardmain_persondef = New_CardMain_����DEF
   PropertyChanged "CardMain_����DEF"
   If Me.CardMain_����DEF = -99 Then
       personlabeldef.Caption = "?"
   ElseIf Me.CardMain_����DEF < 0 Then
       Me.CardMain_����DEF = 0
       personlabeldef.Caption = Me.CardMain_����DEF
   Else
       personlabeldef.Caption = Me.CardMain_����DEF
   End If
End Property
Public Property Get CardMain_�O�_���s�˦���T() As Boolean
   CardMain_�O�_���s�˦���T = m_cardmain_isnewtype
End Property
Public Property Let CardMain_�O�_���s�˦���T(ByVal New_CardMain_�O�_���s�˦���T As Boolean)
   m_cardmain_isnewtype = New_CardMain_�O�_���s�˦���T
   PropertyChanged "CardMain_�O�_���s�˦���T"
   If m_cardmain_isnewtype = False Then
        personlabelhp.Left = 555
        personlabelhp.Top = 3240
        personlabelatk.Left = 1200
        personlabelatk.Top = 3240
        personlabeldef.Left = 1920
        personlabeldef.Top = 3240
   Else
        personlabelhp.Left = 300
        personlabelhp.Top = 3220
        personlabelatk.Left = 960
        personlabelatk.Top = 3220
        personlabeldef.Left = 1820
        personlabeldef.Top = 3220
   End If
End Property
Public Property Get CardBack_�D�ʧ�_�ޯ�W��() As String
   CardBack_�D�ʧ�_�ޯ�W�� = ""
End Property
Public Property Let CardBack_�D�ʧ�_�ޯ�W��(ByVal New_CardBack_�D�ʧ�_�ޯ�W�� As String)
   Dim buffstr() As String
   buffstr = Split(New_CardBack_�D�ʧ�_�ޯ�W��, "#")
   If buffstr(0) <> "" And Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
       PEAFpersoncardback_text(Val(buffstr(1))).Caption = buffstr(0)
       PEAFpersoncardback_text(Val(buffstr(1))).Visible = True
   Else
       PEAFpersoncardback_text(Val(buffstr(1))).Visible = False
   End If
   PropertyChanged "CardBack_�D�ʧ�_�ޯ�W��"
End Property
Public Property Get CardBack_�D�ʧ�_�ޯ໡��() As String
   CardBack_�D�ʧ�_�ޯ໡�� = ""
End Property
Public Property Let CardBack_�D�ʧ�_�ޯ໡��(ByVal New_CardBack_�D�ʧ�_�ޯ໡�� As String)
   Dim buffstr() As String
   buffstr = Split(New_CardBack_�D�ʧ�_�ޯ໡��, "#")
   If buffstr(0) <> "" And Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
       For i = 1 To Len(buffstr(0))
            If Mid(buffstr(0), i, 1) = "&" Then
                Mid(buffstr(0), i, 1) = Chr(10)
            End If
       Next
       m_cardback_activehelp(Val(buffstr(1))) = buffstr(0)
   End If
   PropertyChanged "CardBack_�D�ʧ�_�ޯ໡��"
End Property
Public Property Get CardBack_�D�ʧ�_���q�N�X() As String
   CardBack_�D�ʧ�_���q�N�X = ""
End Property
Public Property Let CardBack_�D�ʧ�_���q�N�X(ByVal New_CardBack_�D�ʧ�_���q�N�X As String)
   Dim buffstr() As String
   buffstr = Split(New_CardBack_�D�ʧ�_���q�N�X, "#")
   If Val(buffstr(0)) >= 1 And Val(buffstr(0)) <= 3 And Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
        PEAFpersoncardback_turn(Val(buffstr(1))).�������O = 3
        PEAFpersoncardback_turn(Val(buffstr(1))).�Ϥ� = App.Path & "\gif\system\cardback\CBturn.png"
        PEAFpersoncardback_turn(Val(buffstr(1))).���ؽs�� = Val(buffstr(0))
        PEAFpersoncardback_turn(Val(buffstr(1))).Visible = True
   Else
        PEAFpersoncardback_turn(Val(buffstr(1))).Visible = False
   End If
   PropertyChanged "CardBack_�D�ʧ�_���q�N�X"
End Property
Public Property Get CardBack_�D�ʧ�_�Z���N�X() As String
   CardBack_�D�ʧ�_�Z���N�X = ""
End Property
Public Property Let CardBack_�D�ʧ�_�Z���N�X(ByVal New_CardBack_�D�ʧ�_�Z���N�X As String)
   Dim buffstr() As String
   Dim k As Integer
   buffstr = Split(New_CardBack_�D�ʧ�_�Z���N�X, "#")
   If Len(buffstr(0)) = 3 And Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
        Select Case Val(buffstr(1))
            Case 1
                For k = 1 To 3
                     PEAFpersoncardback_range1(k).�������O = 2
                     PEAFpersoncardback_range1(k).�Ϥ� = App.Path & "\gif\system\cardback\CBrge.png"
                     If Mid(buffstr(0), k, 1) = 1 Then
                         If k < 3 Then
                             PEAFpersoncardback_range1(k).���ؽs�� = 1
                         Else
                             PEAFpersoncardback_range1(k).���ؽs�� = 3
                         End If
                     Else
                         PEAFpersoncardback_range1(k).���ؽs�� = 2
                     End If
                Next
            Case 2
                For k = 1 To 3
                     PEAFpersoncardback_range2(k).�������O = 2
                     PEAFpersoncardback_range2(k).�Ϥ� = App.Path & "\gif\system\cardback\CBrge.png"
                     If Mid(buffstr(0), k, 1) = 1 Then
                         If k < 3 Then
                             PEAFpersoncardback_range2(k).���ؽs�� = 1
                         Else
                             PEAFpersoncardback_range2(k).���ؽs�� = 3
                         End If
                     Else
                         PEAFpersoncardback_range2(k).���ؽs�� = 2
                     End If
                Next
            Case 3
                For k = 1 To 3
                     PEAFpersoncardback_range3(k).�������O = 2
                     PEAFpersoncardback_range3(k).�Ϥ� = App.Path & "\gif\system\cardback\CBrge.png"
                     If Mid(buffstr(0), k, 1) = 1 Then
                         If k < 3 Then
                             PEAFpersoncardback_range3(k).���ؽs�� = 1
                         Else
                             PEAFpersoncardback_range3(k).���ؽs�� = 3
                         End If
                     Else
                         PEAFpersoncardback_range3(k).���ؽs�� = 2
                     End If
                Next
            Case 4
                For k = 1 To 3
                     PEAFpersoncardback_range4(k).�������O = 2
                     PEAFpersoncardback_range4(k).�Ϥ� = App.Path & "\gif\system\cardback\CBrge.png"
                     If Mid(buffstr(0), k, 1) = 1 Then
                         If k < 3 Then
                             PEAFpersoncardback_range4(k).���ؽs�� = 1
                         Else
                             PEAFpersoncardback_range4(k).���ؽs�� = 3
                         End If
                     Else
                         PEAFpersoncardback_range4(k).���ؽs�� = 2
                     End If
                Next
        End Select
   End If
   PropertyChanged "CardBack_�D�ʧ�_�Z���N�X"
End Property
Public Property Get CardBack_�D�ʧ�_�d���N�X() As String
   CardBack_�D�ʧ�_�d���N�X = ""
End Property
Public Property Let CardBack_�D�ʧ�_�d���N�X(ByVal New_CardBack_�D�ʧ�_�d���N�X As String)
   Dim buffstr() As String, strw() As String
   Dim k As Integer, n As Integer
   buffstr = Split(New_CardBack_�D�ʧ�_�d���N�X, "#")
   If Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
        strw = Split(buffstr(0), "&")
        Select Case Val(buffstr(1))
            Case 1
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               PEAFpersoncardback_num1(k + 1).�������O = 1
                               PEAFpersoncardback_num1(k + 1).�Ϥ� = App.Path & "\gif\system\cardback\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               If Mid(strw(k), 2, 1) = "a" Then
                                    n = 10
                               ElseIf Mid(strw(k), 2, 1) = "b" Then
                                    n = 11
                               Else
                                    n = Val(Mid(strw(k), 2, 1))
                               End If
                               PEAFpersoncardback_num1(k + 1).���ؽs�� = n
                               PEAFpersoncardback_num1(k + 1).Visible = True
                        Else
                               PEAFpersoncardback_num1(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        PEAFpersoncardback_num1(k + 1).Visible = False
                Next
            Case 2
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               PEAFpersoncardback_num2(k + 1).�������O = 1
                               PEAFpersoncardback_num2(k + 1).�Ϥ� = App.Path & "\gif\system\cardback\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               If Mid(strw(k), 2, 1) = "a" Then
                                    n = 10
                               ElseIf Mid(strw(k), 2, 1) = "b" Then
                                    n = 11
                               Else
                                    n = Val(Mid(strw(k), 2, 1))
                               End If
                               PEAFpersoncardback_num2(k + 1).���ؽs�� = n
                               PEAFpersoncardback_num2(k + 1).Visible = True
                        Else
                               PEAFpersoncardback_num2(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        PEAFpersoncardback_num2(k + 1).Visible = False
                Next
            Case 3
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               PEAFpersoncardback_num3(k + 1).�������O = 1
                               PEAFpersoncardback_num3(k + 1).�Ϥ� = App.Path & "\gif\system\cardback\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               If Mid(strw(k), 2, 1) = "a" Then
                                    n = 10
                               ElseIf Mid(strw(k), 2, 1) = "b" Then
                                    n = 11
                               Else
                                    n = Val(Mid(strw(k), 2, 1))
                               End If
                               PEAFpersoncardback_num3(k + 1).���ؽs�� = n
                               PEAFpersoncardback_num3(k + 1).Visible = True
                        Else
                               PEAFpersoncardback_num3(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        PEAFpersoncardback_num3(k + 1).Visible = False
                Next
            Case 4
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               PEAFpersoncardback_num4(k + 1).�������O = 1
                               PEAFpersoncardback_num4(k + 1).�Ϥ� = App.Path & "\gif\system\cardback\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               If Mid(strw(k), 2, 1) = "a" Then
                                    n = 10
                               ElseIf Mid(strw(k), 2, 1) = "b" Then
                                    n = 11
                               Else
                                    n = Val(Mid(strw(k), 2, 1))
                               End If
                               PEAFpersoncardback_num4(k + 1).���ؽs�� = n
                               PEAFpersoncardback_num4(k + 1).Visible = True
                        Else
                               PEAFpersoncardback_num4(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        PEAFpersoncardback_num4(k + 1).Visible = False
                Next
        End Select
   End If
   PropertyChanged "CardBack_�D�ʧ�_�d���N�X"
End Property
Public Property Get CardBack_�Q�ʧ�_�ޯ�W��() As String
   CardBack_�Q�ʧ�_�ޯ�W�� = ""
End Property
Public Property Let CardBack_�Q�ʧ�_�ޯ�W��(ByVal New_CardBack_�Q�ʧ�_�ޯ�W�� As String)
   Dim buffstr() As String
   buffstr = Split(New_CardBack_�Q�ʧ�_�ޯ�W��, "#")
   If buffstr(0) <> "" And Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
       PEAFpersoncardback_passivetext(Val(buffstr(1))).Caption = buffstr(0)
       PEAFpersoncardback_passivetext(Val(buffstr(1))).Visible = True
   Else
       PEAFpersoncardback_passivetext(Val(buffstr(1))).Visible = False
   End If
   PropertyChanged "CardBack_�Q�ʧ�_�ޯ�W��"
End Property
Public Property Get CardBack_�Q�ʧ�_�ޯ໡��() As String
   CardBack_�Q�ʧ�_�ޯ໡�� = ""
End Property
Public Property Let CardBack_�Q�ʧ�_�ޯ໡��(ByVal New_CardBack_�Q�ʧ�_�ޯ໡�� As String)
   Dim buffstr() As String
   buffstr = Split(New_CardBack_�Q�ʧ�_�ޯ໡��, "#")
   If buffstr(0) <> "" And Val(buffstr(1)) >= 1 And Val(buffstr(1)) <= 4 Then
       For i = 1 To Len(buffstr(0))
            If Mid(buffstr(0), i, 1) = "&" Then
                Mid(buffstr(0), i, 1) = Chr(10)
            End If
       Next
       m_cardback_passivehelp(Val(buffstr(1))) = buffstr(0)
   End If
   PropertyChanged "CardBack_�Q�ʧ�_�ޯ໡��"
End Property
Private Sub cardactiveChickimage_Click()
PEAFcardback.Visible = False
PEAFcardback.Left = 0
PEAFcardback.Top = 0
PEAFcardback.Visible = True
PEAFcardback.ZOrder
PEAFcardbackpassive.Visible = False
m_cardbackcheck = 1
End Sub

Private Sub cardback_Click(ByVal Button As Integer)
Call PEAFcardusbackclick_Click(Button)
End Sub

Private Sub cardback_MouseExit()
PEAFcardusbackclick.Visible = False
End Sub

Private Sub cardback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PEAFcardusbackclick.Visible = True
PEAFcardusbackclick.ZOrder
End Sub

Private Sub cardpassiveChickimage_Click()
PEAFcardbackpassive.Visible = False
PEAFcardbackpassive.Left = 0
PEAFcardbackpassive.Top = 0
PEAFcardbackpassive.Visible = True
PEAFcardbackpassive.ZOrder
PEAFcardback.Visible = False
m_cardbackcheck = 2
End Sub

Private Sub PEAFcardback_Click()
'wmp1.Controls.stop
'wmp1.Controls.play
card.Visible = False
card.Left = 0
card.Top = 0
card.Visible = True
card.ZOrder
PEAFcardback.Visible = False
m_cardbackcheck = 1
End Sub

Private Sub PEAFcardback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 4
    If i <> m_cardback_activecheck Then
        PEAFcardbackBR(i).Opacity = 0
    End If
Next
End Sub

Private Sub PEAFcardbackBR_Click(Index As Integer, ByVal Button As Integer)
m_cardback_activecheck = Index
PEAFpersoncardback_main.Caption = m_cardback_activehelp(Index)
End Sub

Private Sub PEAFcardbackBR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PEAFcardbackBR(Index).Opacity = 100
For i = 1 To 4
    If i <> m_cardback_activecheck And i <> Index Then
        PEAFcardbackBR(i).Opacity = 0
    End If
Next
End Sub

Private Sub PEAFcardbackpassive_Click()
'wmp1.Controls.stop
'wmp1.Controls.play
card.Visible = False
card.Left = 0
card.Top = 0
card.Visible = True
card.ZOrder
PEAFcardbackpassive.Visible = False
m_cardbackcheck = 2
End Sub

Private Sub PEAFcardbackpassive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 4
    If i <> m_cardback_passivecheck And i <> Index Then
        PEAFcardbackpassiveBR(i).Opacity = 0
    End If
Next
End Sub

Private Sub PEAFcardbackpassiveBR_Click(Index As Integer, ByVal Button As Integer)
m_cardback_passivecheck = Index
PEAFpersoncardback_passivemain.Caption = m_cardback_passivehelp(Index)
End Sub

Private Sub PEAFcardbackpassiveBR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PEAFcardbackpassiveBR(Index).Opacity = 50
For i = 1 To 4
    If i <> m_cardback_passivecheck And i <> Index Then
        PEAFcardbackpassiveBR(i).Opacity = 0
    End If
Next
End Sub

Private Sub PEAFcardusbackclick_Click(ByVal Button As Integer)
'wmp1.Controls.stop
'wmp1.Controls.play
If m_cardbackcheck <= 1 Then
    PEAFcardback.Visible = False
    PEAFcardback.Left = 0
    PEAFcardback.Top = 0
    PEAFcardback.Visible = True
    PEAFcardback.ZOrder
    m_cardbackcheck = 1
Else
    PEAFcardbackpassive.Visible = False
    PEAFcardbackpassive.Left = 0
    PEAFcardbackpassive.Top = 0
    PEAFcardbackpassive.Visible = True
    PEAFcardbackpassive.ZOrder
    m_cardbackcheck = 2
End If
card.Visible = False
End Sub

Private Sub PEAFpersoncardback_main_Click()
Call PEAFcardback_Click
End Sub

Private Sub PEAFpersoncardback_main_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 4
    If i <> m_cardback_activecheck Then
        PEAFcardbackBR(i).Opacity = 0
    End If
Next
End Sub

Private Sub PEAFpersoncardback_passivemain_Click()
Call PEAFcardbackpassive_Click
End Sub

Private Sub PEAFpersoncardback_passivemain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 4
    If i <> m_cardback_passivecheck Then
        PEAFcardbackpassiveBR(i).Opacity = 0
    End If
Next
End Sub

Private Sub PEAFpersoncardback_passivetext_Click(Index As Integer)
Call PEAFcardbackpassiveBR_Click(Index, 0)
End Sub

Private Sub PEAFpersoncardback_passivetext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PEAFcardbackpassiveBR_MouseMove(Index, 0, 0, 0, 0)
End Sub

Private Sub PEAFpersoncardback_text_Click(Index As Integer)
Call PEAFcardbackBR_Click(Index, 0)
End Sub

Private Sub PEAFpersoncardback_text_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PEAFcardbackBR_MouseMove(Index, 0, 0, 0, 0)
End Sub

Private Sub UserControl_Initialize()
'wmp1.settings.volume = 20
'wmp1.settings.playCount = 1
'wmp1.URL = App.Path & "\mp3\ulse23.mp3"
'wmp1.Controls.stop
If personlabelhp.FontName <> "Bradley Gratis" Then
    personlabelhp.FontSize = 14
    personlabelatk.FontSize = 14
    personlabeldef.FontSize = 14
End If
End Sub
