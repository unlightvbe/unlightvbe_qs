VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form FormMainMode 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "UnlightVBE-QS Origin"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20400
   BeginProperty Font 
      Name            =   "�L�n������"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMainMode.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   20400
   StartUpPosition =   2  '�ù�����
   Tag             =   "UnlightVBE-QS Origin"
   Begin VB.PictureBox PEAttackingForm 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   9910
      Left            =   4800
      Picture         =   "FormMainMode.frx":0CCA
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   11340
      Begin VB.CommandButton �v�l�]�w 
         Caption         =   "�v�l�]�w"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8760
         TabIndex        =   125
         Top             =   9360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   124
         Top             =   9360
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "���}"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   123
         Top             =   9360
         Width           =   1215
      End
      Begin VB.PictureBox PEAFtoolbox 
         Height          =   2175
         Left            =   4920
         ScaleHeight     =   2115
         ScaleWidth      =   6075
         TabIndex        =   111
         Top             =   6240
         Visible         =   0   'False
         Width           =   6135
         Begin VB.Timer �������q_���q��l 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   4680
            Top             =   240
         End
         Begin VB.Timer ���ʶ��q_���q��l 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   1920
            Top             =   1200
         End
         Begin VB.Timer ���m���q_���q��l 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   4680
            Top             =   840
         End
         Begin VB.Timer NextTurn_���q2 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   3840
            Top             =   1200
         End
         Begin VB.CommandButton cn1 
            Caption         =   "�o�P"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   119
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cnmove2 
            Caption         =   "OK"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   118
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cnmove 
            Caption         =   "�U�@�B"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   117
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn32 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   116
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn22 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   115
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn3 
            Caption         =   "�U�@�B"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   114
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn2 
            Caption         =   "�U�@�B"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   113
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cn4 
            Caption         =   "Next Turn"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   112
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Timer OK���s�P���������ˬd 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   1200
            Top             =   0
         End
         Begin VB.Timer ��������ˬd 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   600
            Top             =   0
         End
         Begin VB.Timer �������q_���q1 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   5040
            Top             =   240
         End
         Begin VB.Timer �������q_���q2 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   5400
            Top             =   240
         End
         Begin VB.Timer �ϥΪ̥X�P_��P��� 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   0
            Top             =   120
         End
      End
      Begin MSScriptControlCtl.ScriptControl PEAFvssc 
         Index           =   0
         Left            =   2640
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         UseSafeSubset   =   -1  'True
      End
      Begin VB.Timer �ϥΪ̥X�P_AI�X�P����_�ƥ�d 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3720
         Top             =   5640
      End
      Begin VB.Timer �ϥΪ̥X�P_AI�X�P���� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3240
         Top             =   5640
      End
      Begin VB.Timer �H�������ˬd 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   2400
         Top             =   2640
      End
      Begin VB.Timer tr�P��_�^�P_�q�� 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1320
         Top             =   3840
      End
      Begin VB.Timer atkingtrus 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   10920
         Top             =   5400
      End
      Begin VB.Timer trend 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   10920
         Top             =   4920
      End
      Begin VB.Timer trnextend 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   5040
      End
      Begin VB.Timer �P���� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   960
         Top             =   2760
      End
      Begin VB.Timer �o�P_�ϥΪ̶��q 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   480
         Top             =   2520
      End
      Begin VB.Timer �o�P_�q�����q 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   480
         Top             =   3000
      End
      Begin VB.Timer �o�P�ˬd 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   2760
      End
      Begin VB.Timer �P����_���P 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   1680
      End
      Begin VB.Timer �ϥΪ̥X�P_�X�P���_�a�� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   5520
         Top             =   5520
      End
      Begin VB.Timer �ϥΪ̥X�P_�X�P���_�a�k 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   5760
         Top             =   5520
      End
      Begin VB.Timer atkingtrcom 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   10920
         Top             =   3120
      End
      Begin VB.Timer �q���X�P 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8280
         Top             =   120
      End
      Begin VB.Timer �q���X�P_�X�P���_�a�� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7320
         Top             =   1080
      End
      Begin VB.Timer �q���X�P_��P��� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7800
         Top             =   120
      End
      Begin VB.Timer �q���X�P_�G�P 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   7440
         Top             =   1560
      End
      Begin VB.Timer ���P���q_�p�� 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   2160
      End
      Begin VB.Timer ��l���槹�Ұ� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   720
         Top             =   5040
      End
      Begin VB.Timer ���ݮɶ� 
         Enabled         =   0   'False
         Interval        =   375
         Left            =   10920
         Top             =   2640
      End
      Begin VB.Timer �p�H���Y������_�ϥΪ� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3840
         Top             =   1080
      End
      Begin VB.Timer trgoi1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3000
         Top             =   3120
      End
      Begin VB.Timer trgoi2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   8160
         Top             =   3120
      End
      Begin VB.Timer �p�H���Y������_�q�� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4200
         Top             =   1080
      End
      Begin VB.Timer ���ʹϤ������ˬd 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1680
         Top             =   1920
      End
      Begin VB.Timer tr�q���P_½�P 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   8280
         Top             =   1080
      End
      Begin VB.Timer tr�q���P_���P 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   8280
         Top             =   1560
      End
      Begin VB.Timer tr�P��_�^�P_�ϥΪ� 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1080
         Top             =   3840
      End
      Begin VB.Timer tr�ϥΪ�_��P 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1080
         Top             =   4440
      End
      Begin VB.Timer tr�q���P_��P 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   7920
         Top             =   1560
      End
      Begin VB.Timer tr�P��_��P_�ϥΪ� 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1200
         Top             =   4440
      End
      Begin VB.Timer tr�P��_��P_�q�� 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1920
         Top             =   1440
      End
      Begin VB.Timer trtimeline 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   5520
         Top             =   4920
      End
      Begin VB.Timer ��q���J�ʵe 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7440
         Top             =   5640
      End
      Begin VB.Timer ���ݮɶ�_2 
         Enabled         =   0   'False
         Interval        =   187
         Left            =   10560
         Top             =   2640
      End
      Begin VB.Timer tr�ϥΪ̵P_���P 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   5520
      End
      Begin VB.Timer �q���X�P_�X�P���_�a�k 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7560
         Top             =   1080
      End
      Begin UnlightVBE.uc����d������ cardcom 
         Height          =   3615
         Index           =   0
         Left            =   600
         TabIndex        =   129
         Top             =   6120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   2355
         _ExtentY        =   3625
      End
      Begin UnlightVBE.uc����d������ cardus 
         Height          =   3615
         Index           =   0
         Left            =   0
         TabIndex        =   128
         Top             =   6120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   2355
         _ExtentY        =   3625
      End
      Begin UnlightVBE.uc�ޯ�ʵe���� PEAFAnimateInterface 
         Height          =   9910
         Left            =   0
         Top             =   0
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   17489
      End
      Begin UnlightVBE.uc����p�d PEAFpersoncardcom 
         Height          =   495
         Index           =   3
         Left            =   5040
         TabIndex        =   138
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc����p�d PEAFpersoncardcom 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   137
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc����p�d PEAFpersoncardcom 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc����p�d PEAFpersoncardus 
         Height          =   495
         Index           =   3
         Left            =   5040
         TabIndex        =   135
         Top             =   9360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc����p�d PEAFpersoncardus 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   134
         Top             =   9360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc����p�d PEAFpersoncardus 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   133
         Top             =   9360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
      End
      Begin UnlightVBE.uc�ޯ໡�� PEAFatkinghelpc 
         Height          =   3255
         Left            =   2640
         TabIndex        =   132
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5741
      End
      Begin UnlightVBE.uc�԰��t�εP������ PEAFInterface 
         Height          =   9915
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   11340
         _ExtentX        =   2143
         _ExtentY        =   2778
      End
      Begin VB.Label bloodnumcom2 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11040
         TabIndex        =   2
         Top             =   5850
         Width           =   300
      End
      Begin VB.Label bloodnumcom1 
         Alignment       =   1  '�a�k���
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   10560
         TabIndex        =   3
         Top             =   5600
         Width           =   375
      End
      Begin VB.Image PEAFbloodbackimage2 
         Height          =   690
         Left            =   10080
         Picture         =   "FormMainMode.frx":22BA9
         Top             =   5440
         Width           =   1275
      End
      Begin VB.Label bloodnumus2 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   590
         TabIndex        =   4
         Top             =   5820
         Width           =   300
      End
      Begin VB.Label bloodnumus1 
         Alignment       =   1  '�a�k���
         BackStyle       =   0  '�z��
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   30
         TabIndex        =   5
         Top             =   5600
         Width           =   450
      End
      Begin VB.Image PEAFbloodbackimage1 
         Height          =   690
         Left            =   0
         Picture         =   "FormMainMode.frx":232F0
         Top             =   5440
         Width           =   1290
      End
      Begin UnlightVBE.uc�Y�뤶�� PEAFDiceInterface 
         Height          =   9910
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Width           =   11340
         _ExtentX        =   2566
         _ExtentY        =   1296
      End
      Begin VB.Label pageusglead 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   122
         Top             =   6600
         Width           =   135
      End
      Begin UnlightVBE.ucCard card 
         Height          =   1335
         Index           =   0
         Left            =   5280
         TabIndex        =   121
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2355
      End
      Begin VB.Label pagecomglead 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   120
         Top             =   720
         Width           =   135
      End
      Begin UnlightVBE.��ܦC ��ܦC1 
         Height          =   1215
         Left            =   0
         TabIndex        =   1
         Top             =   3520
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2143
      End
      Begin VB.Image atkdef2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":23A80
         Top             =   1860
         Width           =   2280
      End
      Begin VB.Image atkdef1 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":241CA
         Top             =   1590
         Width           =   2280
      End
      Begin VB.Image draw2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":24920
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Image move2 
         Height          =   270
         Left            =   9120
         Picture         =   "FormMainMode.frx":24F87
         Top             =   1320
         Width           =   2280
      End
      Begin VB.Image PEAFtest 
         Height          =   2865
         Index           =   1
         Left            =   2040
         Picture         =   "FormMainMode.frx":256A9
         Tag             =   "personusminijpg"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label pageusqlead 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   5880
         Width           =   135
      End
      Begin VB.Label pagecomqlead 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "0"
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
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   2160
         Width           =   135
      End
      Begin VB.Image Image2 
         Height          =   120
         Left            =   5280
         Picture         =   "FormMainMode.frx":271D7
         Top             =   6120
         Width           =   780
      End
      Begin VB.Shape bloodlineout1 
         BorderStyle     =   0  '�z��
         FillColor       =   &H000000FF&
         FillStyle       =   0  '���
         Height          =   80
         Left            =   0
         Top             =   6160
         Width           =   5295
      End
      Begin VB.Shape bloodlineout2 
         BorderStyle     =   0  '�z��
         FillColor       =   &H000000FF&
         FillStyle       =   0  '���
         Height          =   75
         Left            =   6060
         Top             =   6160
         Width           =   5295
      End
      Begin VB.Image timeup 
         Height          =   105
         Left            =   5290
         Picture         =   "FormMainMode.frx":2726A
         Top             =   4720
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Line timelineout1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   0
         X2              =   5250
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Line timelineout2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   6060
         X2              =   11310
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Image PEAFtest 
         Height          =   2880
         Index           =   2
         Left            =   9960
         Picture         =   "FormMainMode.frx":272D6
         Tag             =   "personcomminijpg"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape bloodlinein1 
         BorderStyle     =   6  '����u
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  '���
         Height          =   90
         Left            =   0
         Top             =   6150
         Width           =   5295
      End
      Begin VB.Shape bloodlinein2 
         BorderStyle     =   6  '����u
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  '���
         Height          =   90
         Left            =   6060
         Top             =   6150
         Width           =   5295
      End
      Begin VB.Shape timelinein1 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  '����u
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  '���
         Height          =   90
         Left            =   0
         Top             =   4720
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Shape timelinein2 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  '����u
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  '���
         Height          =   90
         Left            =   6050
         Top             =   4720
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Image draw1 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":28D00
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Image move1 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":28E7F
         Top             =   1340
         Width           =   2040
      End
      Begin VB.Image move3 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":2901B
         Top             =   1610
         Width           =   2040
      End
      Begin VB.Image move4 
         Height          =   240
         Left            =   9360
         Picture         =   "FormMainMode.frx":292A4
         Top             =   1880
         Width           =   2040
      End
      Begin UnlightVBE.�p�H���ζH personusminijpg 
         Height          =   4935
         Left            =   0
         TabIndex        =   9
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8705
      End
      Begin VB.Image cardpagejpg 
         Height          =   915
         Left            =   0
         Picture         =   "FormMainMode.frx":29521
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label pageul 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "57"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   1100
         Width           =   855
      End
      Begin UnlightVBE.�p�H���ζH personcomminijpg 
         Height          =   4935
         Left            =   5520
         TabIndex        =   10
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8705
      End
      Begin ImageX.aicAlphaImage PEAFMoveRange 
         Height          =   1080
         Left            =   2880
         Top             =   2160
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   1905
         Image           =   "FormMainMode.frx":29D84
         Scaler          =   3
         Props           =   13
      End
   End
   Begin VB.PictureBox PEGameFreeModeSettingForm 
      Appearance      =   0  '����
      BackColor       =   &H80000000&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   2760
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   11340
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   11415
         TabIndex        =   36
         Top             =   4320
         Width           =   11415
         Begin VB.ComboBox personlevelus 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox personlevelus 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2760
            TabIndex        =   50
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox personlevelus 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5400
            TabIndex        =   49
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox personnameus 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   48
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox personnameus 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3720
            TabIndex        =   47
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox personnameus 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6360
            TabIndex        =   46
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox personlevelcom 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3360
            TabIndex        =   45
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox personlevelcom 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   6000
            TabIndex        =   44
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox personlevelcom 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   8640
            TabIndex        =   43
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox personnamecom 
            Appearance      =   0  '����
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4320
            TabIndex        =   42
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox personnamecom 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   6960
            TabIndex        =   41
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox personnamecom 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   9600
            TabIndex        =   40
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton opnpersonvs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3v3"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton personreadifus 
            Caption         =   "Ū�J..."
            Height          =   495
            Left            =   2040
            TabIndex        =   37
            Top             =   720
            Width           =   975
         End
         Begin MSComDlg.CommonDialog cdgpersonus 
            Left            =   3000
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "UnlightVBE-�d���H����T-�}���ɮ�"
         End
         Begin VB.OptionButton opnpersonvs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1v1"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   855
         End
         Begin ImageX.aicAlphaImage aicAlphaImage1 
            Height          =   660
            Left            =   5160
            Top             =   600
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   1164
            Image           =   "FormMainMode.frx":2E650
            Props           =   5
         End
         Begin UnlightVBE.�j�H���ι� personfus 
            Height          =   1215
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2143
         End
         Begin VB.Image PEGFbnstart 
            Height          =   510
            Left            =   9600
            Picture         =   "FormMainMode.frx":2F53D
            Top             =   600
            Width           =   1440
         End
         Begin VB.Image bnabout 
            Height          =   390
            Left            =   8280
            Picture         =   "FormMainMode.frx":30093
            Top             =   720
            Width           =   1320
         End
         Begin VB.Image bnconfig 
            Height          =   390
            Left            =   7080
            Picture         =   "FormMainMode.frx":306FE
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label personresetus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   59
            Top             =   360
            Width           =   495
         End
         Begin VB.Label personresetus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   58
            Top             =   360
            Width           =   495
         End
         Begin VB.Label personresetus 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   6480
            TabIndex        =   57
            Top             =   360
            Width           =   495
         End
         Begin VB.Label personresetcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   56
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label personresetcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   6960
            TabIndex        =   55
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label personresetcom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���]"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   9600
            TabIndex        =   54
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label PEGFplayertext 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '�z��
            Caption         =   "1P"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   8160
            TabIndex        =   53
            Top             =   0
            Width           =   375
         End
         Begin VB.Label PEGFcomtext 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '�z��
            Caption         =   "COM"
            BeginProperty Font 
               Name            =   "Bradley Gratis"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   2400
            TabIndex        =   52
            Top             =   1395
            Width           =   855
         End
      End
      Begin VB.PictureBox PEGFcardus 
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
         Index           =   1
         Left            =   120
         Picture         =   "FormMainMode.frx":30D0E
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   32
         Top             =   600
         Width           =   2535
         Begin VB.Label PEGFusbi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   550
            TabIndex        =   35
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEGFusbi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   34
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFusbi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   33
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardus 
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
         Index           =   2
         Left            =   2760
         Picture         =   "FormMainMode.frx":34BB1
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   28
         Top             =   600
         Width           =   2535
         Begin VB.Label PEGFusbi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   550
            TabIndex        =   31
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEGFusbi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1200
            TabIndex        =   30
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFusbi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1920
            TabIndex        =   29
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardus 
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
         Index           =   3
         Left            =   5400
         Picture         =   "FormMainMode.frx":38A54
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   24
         Top             =   600
         Width           =   2535
         Begin VB.Label PEGFusbi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   550
            TabIndex        =   27
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEGFusbi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1200
            TabIndex        =   26
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFusbi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1920
            TabIndex        =   25
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardcom 
         Appearance      =   0  '����
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   1
         Left            =   3360
         Picture         =   "FormMainMode.frx":3C8F7
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   20
         Top             =   6240
         Width           =   2535
         Begin VB.Label PEGFcardcompi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   23
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   22
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   21
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardcom 
         Appearance      =   0  '����
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   2
         Left            =   6000
         Picture         =   "FormMainMode.frx":4079A
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   16
         Top             =   6240
         Width           =   2535
         Begin VB.Label PEGFcardcompi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   19
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1200
            TabIndex        =   18
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1920
            TabIndex        =   17
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEGFcardcom 
         Appearance      =   0  '����
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   3
         Left            =   8640
         Picture         =   "FormMainMode.frx":4463D
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   12
         Top             =   6240
         Width           =   2535
         Begin VB.Label PEGFcardcompi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   15
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1200
            TabIndex        =   14
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEGFcardcompi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1920
            TabIndex        =   13
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '�z��
         Caption         =   "GameSetting"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '�z��
         Caption         =   "�ۥѾ԰��Ҧ��C���޾ɳ]�w"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   61
         Top             =   195
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   465
         Left            =   0
         Picture         =   "FormMainMode.frx":484E0
         Top             =   0
         Width           =   11400
      End
   End
   Begin VB.PictureBox PEStartForm 
      Appearance      =   0  '����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   3480
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   63
      Top             =   960
      Visible         =   0   'False
      Width           =   11340
      Begin VB.Timer tr1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   9720
         Top             =   8400
      End
      Begin VB.Label PEStext1 
         Alignment       =   1  '�a�k���
         BackStyle       =   0  '�z��
         Caption         =   "Now  Loading..."
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   8280
         TabIndex        =   64
         Top             =   9120
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.PictureBox PEMusicForm 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   1320
      ScaleHeight     =   7935
      ScaleWidth      =   11895
      TabIndex        =   130
      Top             =   1200
      Visible         =   0   'False
      Width           =   11895
      Begin VB.Timer PEMFtr1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   240
         Top             =   1680
      End
      Begin UnlightVBE.ucMusicPlayer cMusicPlayer 
         Height          =   855
         Index           =   0
         Left            =   840
         TabIndex        =   131
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1508
      End
   End
   Begin VB.PictureBox PEAttackingStartForm 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   3120
      Picture         =   "FormMainMode.frx":4A8CF
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   65
      Top             =   240
      Visible         =   0   'False
      Width           =   11340
      Begin VB.Timer PEASpke 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   240
         Top             =   2880
      End
      Begin VB.Timer start1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   45
         Top             =   4680
      End
      Begin VB.Timer start2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   525
         Top             =   4680
      End
      Begin VB.PictureBox PEAScardcom 
         Appearance      =   0  '����
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   1
         Left            =   6285
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEAScardcompi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "0"
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
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   69
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   68
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   67
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEAScardcom 
         Appearance      =   0  '����
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   2
         Left            =   7485
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   96
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEAScardcompi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1920
            TabIndex        =   99
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEAScardcompi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1200
            TabIndex        =   98
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "0"
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
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   97
            Top             =   3240
            Width           =   495
         End
      End
      Begin VB.PictureBox PEAScardus 
         Appearance      =   0  '����
         BackColor       =   &H00404040&
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
         Index           =   1
         Left            =   2805
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   70
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEASusbi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "0"
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
            Index           =   1
            Left            =   550
            TabIndex        =   73
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label PEASusbi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   72
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEASusbi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   71
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.PictureBox PEAScardus 
         Appearance      =   0  '����
         BackColor       =   &H00404040&
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
         Index           =   2
         Left            =   1485
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   88
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEASusbi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1920
            TabIndex        =   91
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEASusbi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   2
            Left            =   1200
            TabIndex        =   90
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEASusbi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "0"
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
            Index           =   2
            Left            =   550
            TabIndex        =   89
            Top             =   3240
            Width           =   375
         End
      End
      Begin VB.PictureBox PEAScardcom 
         Appearance      =   0  '����
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   3615
         Index           =   3
         Left            =   8565
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   100
         Top             =   3960
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEAScardcompi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1920
            TabIndex        =   103
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEAScardcompi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1200
            TabIndex        =   102
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEAScardcompi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "0"
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
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   101
            Top             =   3240
            Width           =   495
         End
      End
      Begin VB.PictureBox PEAScardus 
         Appearance      =   0  '����
         BackColor       =   &H00404040&
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
         Index           =   3
         Left            =   360
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   92
         Top             =   3960
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label PEASusbi3 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1920
            TabIndex        =   95
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label PEASusbi2 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
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
            Index           =   3
            Left            =   1200
            TabIndex        =   94
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label PEASusbi1 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "0"
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
            Index           =   3
            Left            =   550
            TabIndex        =   93
            Top             =   3240
            Width           =   375
         End
      End
      Begin VB.PictureBox downjpg 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   45
         Picture         =   "FormMainMode.frx":79783
         ScaleHeight     =   1455
         ScaleWidth      =   11415
         TabIndex        =   75
         Top             =   8160
         Visible         =   0   'False
         Width           =   11415
         Begin VB.Label cardusname 
            BackStyle       =   0  '�z��
            Caption         =   "�H��1"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   87
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label cardcomname 
            BackStyle       =   0  '�z��
            Caption         =   "�H��1"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   6840
            TabIndex        =   86
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label cardusspname 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "�ٸ�1"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   85
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label cardcomspname 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "�ٸ�1"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   1
            Left            =   7920
            TabIndex        =   84
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label cardusname 
            BackStyle       =   0  '�z��
            Caption         =   "�H��2"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   83
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label cardusname 
            BackStyle       =   0  '�z��
            Caption         =   "�H��3"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   82
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label cardusspname 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "�ٸ�2"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   81
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label cardusspname 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "�ٸ�3"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   80
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label cardcomname 
            BackStyle       =   0  '�z��
            Caption         =   "�H��2"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   6840
            TabIndex        =   79
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label cardcomname 
            BackStyle       =   0  '�z��
            Caption         =   "�H��3"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   6840
            TabIndex        =   78
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label cardcomspname 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "�ٸ�2"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   2
            Left            =   7920
            TabIndex        =   77
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label cardcomspname 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "�ٸ�3"
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   3
            Left            =   7920
            TabIndex        =   76
            Top             =   840
            Width           =   3135
         End
      End
      Begin VB.PictureBox upjpg 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   1900
         Left            =   0
         Picture         =   "FormMainMode.frx":820FF
         ScaleHeight     =   1905
         ScaleWidth      =   11415
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   11415
      End
      Begin VB.Timer stup 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   45
         Top             =   1800
      End
      Begin VB.Timer stdown 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   45
         Top             =   6600
      End
      Begin VB.Timer cardustr 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3405
         Top             =   7200
      End
      Begin VB.Timer cardcomtr 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7245
         Top             =   7320
      End
      Begin VB.Timer tr�j�H���ι�_�ϥΪ� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1800
         Top             =   7440
      End
      Begin VB.Timer tr�j�H���ι�_�q�� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   9720
         Top             =   7560
      End
      Begin UnlightVBE.uc��� PEASpersontalk 
         Height          =   1935
         Left            =   0
         TabIndex        =   110
         Top             =   -120
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3413
      End
      Begin UnlightVBE.�j�H���ι� �j�H���ι�_�q�� 
         Height          =   10005
         Left            =   20040
         TabIndex        =   104
         Top             =   -480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   17648
      End
      Begin UnlightVBE.�j�H���ι� �j�H���ι�_�ϥΪ� 
         Height          =   10005
         Left            =   -9960
         TabIndex        =   105
         Top             =   -480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   17648
      End
      Begin UnlightVBE.�j�H���ι� upjpg_2 
         Height          =   1935
         Left            =   0
         TabIndex        =   106
         Top             =   -480
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3413
      End
   End
   Begin VB.PictureBox PEAttackingEndingForm 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   9120
      Picture         =   "FormMainMode.frx":8D62B
      ScaleHeight     =   9915
      ScaleWidth      =   11340
      TabIndex        =   107
      Top             =   -1680
      Visible         =   0   'False
      Width           =   11340
      Begin VB.Timer PEAEtr1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5760
         Top             =   8400
      End
      Begin VB.Label bnt 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "�����C��"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9480
         TabIndex        =   109
         Top             =   8760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label bnreturnt 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "��^���"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   108
         Top             =   8760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image bn 
         Height          =   990
         Left            =   9480
         Picture         =   "FormMainMode.frx":B0596
         Top             =   8520
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Image bnreturn 
         Height          =   990
         Left            =   7680
         Picture         =   "FormMainMode.frx":B148B
         Top             =   8520
         Visible         =   0   'False
         Width           =   1470
      End
   End
End
Attribute VB_Name = "FormMainMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub atkingtrcom_Timer()
If �ثe��(29) = 1 Then
   �ثe��(31) = 0
   Formatkingcom.Left = FormMainMode.Left + (FormMainMode.Width - Formatkingcom.Width)
   Formatkingcom.Top = FormMainMode.Top + 380
   atkingtrcom.Enabled = False
   Formatkingcom.t1.Enabled = True
   Formatkingcom.Show 0, Me
Else
   �ثe��(29) = �ثe��(29) + 1
End If
End Sub

Private Sub atkingtrus_Timer()
If �ثe��(29) = 1 Then
   �ثe��(31) = 0
   Formatkingus.Left = FormMainMode.Left
   Formatkingus.Top = FormMainMode.Top + 380
   atkingtrus.Enabled = False
   Formatkingus.t1.Enabled = True
   Formatkingus.Show 0, Me
Else
   �ثe��(29) = �ثe��(29) + 1
End If
End Sub

Private Sub bloodnumus1_Change()
If Val(bloodnumus1.Caption) < 0 Then bloodnumus1.Caption = 0
End Sub

Private Sub bn_Click()
End
End Sub

Private Sub bnreturn_Click()
bnreturnt_Click
End Sub

Sub bnreturnt_Click()
����Ū�J���� = "PEGF"
�@��t����.�D���_PEStartForm���
FormMainMode.PEAttackingEndingForm.Visible = False
End Sub

Private Sub bnt_Click()
End
End Sub

Sub card_CardButtonClickin(Index As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(5)(CStr(Index))

Call tmpcard.Reverse
�@��t����.���ļ��� 3
End Sub

Sub card_CardButtonClickout(Index As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(6)(CStr(Index))

Call tmpcard.Reverse
FormMainMode.card(Index).CardRotationType = tmpcard.CardOnIn
�@��t����.���ļ��� 3
'===================================================================
If tmpcard.UpperType = a1a Then
   atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) + Val(tmpcard.UpperNum)
   If turnatk = 1 And movecp = 1 And �������m��l�`��(3) = 0 Then
       �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
   End If
   If turnatk = 1 And movecp = 1 Then
       �������m��l�`��(1) = �������m��l�`��(1) + Val(tmpcard.UpperNum)
       �������m��l�`��(3) = �������m��l�`��(3) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a5a Then
   atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) + Val(tmpcard.UpperNum)
   If turnatk = 1 And movecp > 1 And �������m��l�`��(3) = 0 Then
       �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
   End If
   If turnatk = 1 And movecp > 1 Then
       �������m��l�`��(1) = �������m��l�`��(1) + Val(tmpcard.UpperNum)
       �������m��l�`��(3) = �������m��l�`��(3) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a2a Then
   atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) + Val(tmpcard.UpperNum)
   If turnatk = 2 And �������m��l�`��(3) = 0 Then
       �������m��l�`��(3) = �������m��l�`��(3) + defus(����H����ԤH��(1, 2))
   End If
   If turnatk = 2 Then
      �������m��l�`��(1) = �������m��l�`��(1) + Val(tmpcard.UpperNum)
      �������m��l�`��(3) = �������m��l�`��(3) + Val(tmpcard.UpperNum)
   End If
End If
If tmpcard.UpperType = a3a Then
   atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) + Val(tmpcard.UpperNum)
End If
If tmpcard.UpperType = a4a Then
   atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) + Val(tmpcard.UpperNum)
End If
'======================================
If tmpcard.LowerType = a1a Then
   atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) - Val(tmpcard.LowerNum)
   If turnatk = 1 And movecp = 1 Then
       �������m��l�`��(1) = �������m��l�`��(1) - Val(tmpcard.LowerNum)
       �������m��l�`��(3) = �������m��l�`��(3) - Val(tmpcard.LowerNum)
   End If
   If �������m��l�`��(3) = atkus(����H����ԤH��(1, 2)) Then
       �������m��l�`��(3) = 0
   End If
End If
If tmpcard.LowerType = a5a Then
   atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) - Val(tmpcard.LowerNum)
   If turnatk = 1 And movecp > 1 Then
       �������m��l�`��(1) = �������m��l�`��(1) - Val(tmpcard.LowerNum)
       �������m��l�`��(3) = �������m��l�`��(3) - Val(tmpcard.LowerNum)
   End If
   If �������m��l�`��(3) = atkus(����H����ԤH��(1, 2)) Then
       �������m��l�`��(3) = 0
   End If
End If
If tmpcard.LowerType = a2a Then
   atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) - Val(tmpcard.LowerNum)
   If turnatk = 2 Then
       �������m��l�`��(1) = �������m��l�`��(1) - Val(tmpcard.LowerNum)
       �������m��l�`��(3) = �������m��l�`��(3) - Val(tmpcard.LowerNum)
   End If
End If
If tmpcard.LowerType = a3a Then
   atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) - Val(tmpcard.LowerNum)
End If
If tmpcard.LowerType = a4a Then
   atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) - Val(tmpcard.LowerNum)
End If
'==============================================
Select Case turnatk
    Case 1
        '===========================���涥�q���J�I(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 43, 4
        '============================
    Case 2
        '===========================���涥�q���J�I(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 43, 4
        '============================
    Case 3
        '===========================���涥�q���J�I(44)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 44
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 44, 3
        '============================
End Select
�԰��t����.��q��s���
FormMainMode.trgoi1.Enabled = True
End Sub


Sub card_CardClick(Index As Integer)
Dim tmpcard As clsActionCard
Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(Index)))(CStr(Index))
'======================�H�U���M�ݨƥ�d�ˬd
If tmpcard.UpperType = a7a And turnatk <> 1 And turnatk <> 2 Then
   '=========�H�϶A�G�N�ƥ�d�u�b�𨾶��q�ϥέ�h
   Exit Sub
End If
'====================================
If tmpcard.Location = 1 And (turnpageonin = 1 Or turnpageoninatking = 1) And tmpcard.Owner = 1 Then
   tmpcard.Location = 2
   If tmpcard.UpperType = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) + Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp = 1 And �������m��l�`��(3) = 0 Then
          �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
      End If
      If turnatk = 1 And movecp = 1 Then
          �������m��l�`��(1) = �������m��l�`��(1) + Val(tmpcard.UpperNum)
          �������m��l�`��(3) = �������m��l�`��(3) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) + Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp > 1 And �������m��l�`��(3) = 0 Then
          �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
      End If
      If turnatk = 1 And movecp > 1 Then
          �������m��l�`��(1) = �������m��l�`��(1) + Val(tmpcard.UpperNum)
          �������m��l�`��(3) = �������m��l�`��(3) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) + Val(tmpcard.UpperNum)
      If turnatk = 2 And �������m��l�`��(3) = 0 Then
          �������m��l�`��(3) = �������m��l�`��(3) + defus(����H����ԤH��(1, 2))
      End If
      If turnatk = 2 Then
         �������m��l�`��(1) = �������m��l�`��(1) + Val(tmpcard.UpperNum)
         �������m��l�`��(3) = �������m��l�`��(3) + Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) + Val(tmpcard.UpperNum)
   End If
   If tmpcard.UpperType = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) + Val(tmpcard.UpperNum)
   End If
   '=================
   turnpageonin = 0
   card(Index).LocationType = 2
   '===================
   �ثe��(5) = Utils.IndexOf(�԰��t����.CardDeckCollection(5), tmpcard)
   pageqlead(1) = Val(pageqlead(1)) + 1
   pageusglead = Val(pageusglead) - 1
   pageusleadmax(1) = Val(pageusleadmax(1)) + 1
   pageusqlead = Val(pageusqlead) + 1
   �ثe��(13) = 0
   '===================�H�U�O�X�P���
   �ثe��(3) = 0
   �ϥΪ̥X�P_�X�P���_�a��.Enabled = True
   '=============�H�U�O�P����(�X�P)(�ϥΪ�)
    �԰��t����.�y�Эp��_�ϥΪ̥X�P
    �P���ʼȮ��ܼ�(3) = Index
    tmpcard.XYLeft = card(Index).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 5, 6
    �ثe��(15) = 0
    �P����.Enabled = True
    �@��t����.���ļ��� 1
   '================�H�U�O��P���
   �ثe��(4) = 0
   �ثe��(21) = 1
   �ϥΪ̥X�P_��P���.Enabled = True
   '=================
   If tmpcard.UpperType = a6a Or tmpcard.UpperType = a7a Or tmpcard.UpperType = a8a Or tmpcard.UpperType = a9a Then
        '===================�H�U�O�ƥ�d�ˬd�αҰ�
        ��������ˬd.Enabled = False
        �ƥ�d�O���Ȯɼ�(1, 3) = 1
        Select Case tmpcard.UpperType
            Case a6a
                �ƥ�d.���|_�ϥΪ� Index, tmpcard.UpperNum
            Case a7a
                �ƥ�d.�A�G�N_�ϥΪ� Index, tmpcard.UpperNum
            Case a8a
                �ƥ�d.HP�^�__�ϥΪ� Index, tmpcard.UpperNum
            Case a9a
                �ƥ�d.�t��_�ϥΪ� Index, tmpcard.UpperNum
        End Select
        '===================
        Exit Sub
    Else
        ��������ˬd.Enabled = True
        GoTo vsssystemplay
    End If
End If
'=================================================================
If tmpcard.Location = 2 And (turnpageonin = 1 Or turnpageoninatking = 1) And tmpcard.Owner = 1 Then
   tmpcard.Location = 1
   '===================================
   If tmpcard.UpperType = a1a Then
      atkingpagetot(1, 1) = Val(atkingpagetot(1, 1)) - Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp = 1 Then
          �������m��l�`��(1) = �������m��l�`��(1) - Val(tmpcard.UpperNum)
          �������m��l�`��(3) = �������m��l�`��(3) - Val(tmpcard.UpperNum)
      End If
      If �������m��l�`��(3) = atkus(����H����ԤH��(1, 2)) Then
          �������m��l�`��(3) = 0
      End If
   End If
   If tmpcard.UpperType = a5a Then
      atkingpagetot(1, 5) = Val(atkingpagetot(1, 5)) - Val(tmpcard.UpperNum)
      If turnatk = 1 And movecp > 1 Then
          �������m��l�`��(1) = �������m��l�`��(1) - Val(tmpcard.UpperNum)
          �������m��l�`��(3) = �������m��l�`��(3) - Val(tmpcard.UpperNum)
      End If
      If �������m��l�`��(3) = atkus(����H����ԤH��(1, 2)) Then
          �������m��l�`��(3) = 0
      End If
   End If
   If tmpcard.UpperType = a2a Then
      atkingpagetot(1, 2) = Val(atkingpagetot(1, 2)) - Val(tmpcard.UpperNum)
      If turnatk = 2 Then
         �������m��l�`��(1) = �������m��l�`��(1) - Val(tmpcard.UpperNum)
         �������m��l�`��(3) = �������m��l�`��(3) - Val(tmpcard.UpperNum)
      End If
   End If
   If tmpcard.UpperType = a3a Then
      atkingpagetot(1, 3) = Val(atkingpagetot(1, 3)) - Val(tmpcard.UpperNum)
   End If
   If tmpcard.UpperType = a4a Then
      atkingpagetot(1, 4) = Val(atkingpagetot(1, 4)) - Val(tmpcard.UpperNum)
   End If
   '=============
   turnpageonin = 0
   card(Index).LocationType = 1
   '================
   �ثe��(5) = Utils.IndexOf(�԰��t����.CardDeckCollection(6), tmpcard)
   pageusleadmax(0) = Val(pageusleadmax(0)) + 1
   pageqlead(1) = Val(pageqlead(1)) - 1
   pageusglead = Val(pageusglead) + 1
   pageusqlead = Val(pageusqlead) - 1
   '=============�H�U�O�P����(�^�P)(�ϥΪ�)
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = Index
    tmpcard.XYLeft = card(Index).Left  '���w�ثeLeft(�y��)
    tmpcard.XYTop = card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.�d�P�P�ﶰ�X�� tmpcard, 6, 5
    �ثe��(15) = 0
    �P����.Enabled = True
    �@��t����.���ļ��� 1
   '================�H�U�O�X�P���
   �ثe��(3) = 0
   �ϥΪ̥X�P_�X�P���_�a�k.Enabled = True
   '=====================
   ��������ˬd.Enabled = True
   '=====================�H�U�O�ޯ��ˬd�αҰ�
    If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards") <> 0 Then
        vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards")) = 2 '(���q2)
    End If
    '====================
    GoTo vsssystemplay
End If
'==============================================
Exit Sub
vsssystemplay:
Select Case turnatk
    Case 1
        '===========================���涥�q���J�I(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 43, 4
        '============================
    Case 2
        '===========================���涥�q���J�I(ATK-42/DEF-43)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 42
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 42, 4
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 43
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 43, 4
        '============================
    Case 3
        '===========================���涥�q���J�I(44)
        ReDim VBEStageNum(0 To 1) As Integer
        VBEStageNum(0) = 44
        VBEStageNum(1) = -1 'Ĳ�o��(1.�ϥΪ�/2.�q��)
        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 44, 3
        '============================
End Select
�԰��t����.��q��s���
FormMainMode.trgoi1.Enabled = True
End Sub


Private Sub card_CardMouseMove(Index As Integer)

If �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(Index)) = 5 And turnpageonin = 1 Then
    card(Index).CardEventType = True
ElseIf �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(Index)) = 6 And turnpageonin = 1 Then
    card(Index).CardEventType = True
Else
    card(Index).CardEventType = False
End If
End Sub

Sub cnmove_Click()
Dim i As Integer, med As Integer
Dim tmpcard As clsActionCard
'======================
If �q����ƥ�d�O�_�X����ܼ� = True Then
    GoTo �q����ƥ�d���X���_���涥�q����
End If
'======================
If ����H����ԤH��(1, 1) > 1 Or ����H����ԤH��(2, 1) > 1 Then
   ��ܦC1.�H���԰��H�� = 3
Else
   ��ܦC1.�H���԰��H�� = 1
End If
'======================
movecom = 0
movecheckcom = 0
��ܦC1.���ʶ��q��ܭ� = 0
�q���貾�ʶ��q��ܼ� = 0
atkingtrn(1) = 0
atkingtrn(2) = 0
turnatk = 3
pageusqlead.Caption = 0
pagecomqlead.Caption = 0
�ثe��(6) = 0
�ثe��(17) = 1
�ثe��(21) = 1
�ثe��(25) = 0
���q���A�� = 3
'=============
If �t����ܬɭ������� = 1 Then
    draw2.Visible = False
    draw1.Visible = True
    move1.Visible = False
    move2.Visible = True
Else
    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\move1-2.gif"
End If
��ܦC1.��ܦC�Ϥ� = app_path & "gif\system\linemove.png"
cnmove.Visible = False
�԰��t����.cleanatkingpagetot
'======================�q����ƥ�d���X���
If �q����ƥ�d�O�_�X����ܼ� = False Then
    GoTo �q����ƥ�d���X���_���涥�q2
End If
'================================
�q����ƥ�d���X���_���涥�q����:
'----------�H�U���q���P�_�X�P�{���X�]���ʶ��q1�^
'====================���紼�z��AI�X�P�t��
    ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_��� 2, 3, namecom(����H����ԤH��(2, 2)), movecp, 0
    GoTo ���z��AI�X�P_���涥�q����
'======================
Dim movecomatk1, movecomatk2 As Integer
�԰��t����.moveatkin

For i = 1 To �԰��t����.CardDeckCollection(7).Count
    Set tmpcard = �԰��t����.CardDeckCollection(7)(i)
    If tmpcard.ComMark <> 1 Then
        If tmpcard.UpperType = a1a Then movecomatk1 = Val(movecomatk1) + Val(tmpcard.UpperNum)
        If tmpcard.UpperType = a5a Then movecomatk2 = Val(movecomatk2) + Val(tmpcard.UpperNum)
        If tmpcard.LowerType = a1a Then movecomatk1 = Val(movecomatk1) + Val(tmpcard.LowerNum)
        If tmpcard.LowerType = a5a Then movecomatk2 = Val(movecomatk2) + Val(tmpcard.LowerNum)
    End If
Next
'===========================================
�·�_�q��_���涥�q2: '���`���A-�·�-�q��-�{�����J�I(���涥�q2)
'===========================================
If movecomatk1 > movecomatk2 Then
      �q���貾�ʶ��q��ܼ� = 1
ElseIf movecomatk1 = movecomatk2 Then
      med = Int(Rnd() * 2) + 1
      If med = 1 Then
         �q���貾�ʶ��q��ܼ� = 1
      Else
         �q���貾�ʶ��q��ܼ� = 3
      End If
Else
      �q���貾�ʶ��q��ܼ� = 3
End If
'==============
���z��AI�X�P_���涥�q����:
�q����ƥ�d���X���_���涥�q2:
If �q����ƥ�d�O�_�X����ܼ� = False Then
    '==============
    �p�H���Y�����ʤ�V��(1) = 1
    �p�H���Y�����ʤ�V��(2) = 1
    �p�H���Y������_�ϥΪ�.Enabled = True
    �p�H���Y������_�q��.Enabled = True
    '==============
    ���q���A�� = 1
    �԰��t����.�ɶ��b_���]
    ��ܦC1.���ʶ��q����� = True
    �԰��t����.�ɶ��b_���
    �@��t����.���ļ��� 6
    '===========================���涥�q���J�I(94)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 94, 3
    '============================
End If
'======================�q����ƥ�d���X���_�����ᶥ�q2
If �q����ƥ�d�O�_�X����ܼ� = True Then
    �q���X�P.Enabled = True
End If
'===========================
End Sub

Private Sub cnmove2_Click()
turnpageonin = 0
�ثe��(31) = 0
OK���s�P���������ˬd.Enabled = True
cnmove2.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
�@��t����.���}�C������ Cancel, UnloadMode
End Sub

Private Sub cn1_Click()
turnatk = 4
�԰��t����.���q�R���ո`�]�w
'====================
�ثe��(2) = 1
�q����ƥ�d�O�_�X����ܼ� = False
'===========================���涥�q���J�I(0)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 0, 1
'============================
cn1.Visible = False
�ثe��(15) = 1
�o�P�ˬd.Enabled = True
End Sub

Private Sub cn2_Click()
If moveturn = 1 Then
  If �t����ܬɭ������� = 1 Then
        move1.Visible = True
        move2.Visible = False
        atkdef1.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\atk2.gif"
  End If
  ��ܦC1.goi1��� = True
  ��ܦC1.goi2��� = True
  ��ܦC1.���ʶ��q��ܭ� = 0
  ��ܦC1.���ʶ��q����� = False
Else
  If �t����ܬɭ������� = 1 Then
        atkdef1.Visible = False
        atkdef2.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\atk2.gif"
  End If
End If
'-------------
turnatk = 1
���q���A�� = 1
If movecp = 1 Then
    ��ܦC1.��ܦC�Ϥ� = app_path & "gif\system\lineusatk1.png"
Else
    ��ܦC1.��ܦC�Ϥ� = app_path & "gif\system\lineusatk2.png"
End If
cn2.Visible = False
FormMainMode.PEAFInterface.BnOKStartListen
'=============
�԰��t����.cleanatkingpagetot
'==============
��ܦC1.goi1 = 0
��ܦC1.goi2 = 0
�ثe��(6) = 0
�ثe��(17) = 1
�ثe��(21) = 1
�ثe��(15) = 0
�������m��l�`��(1) = 0
�������m��l�`��(2) = 0
�������m��l�`��(3) = 0
�������m��l�`��(4) = 0
��ƹs�ˬd��(1) = False
��ƹs�ˬd��(2) = False
�O�_�t�Τ��� = False
'==============
goicheck(1) = 0
goicheck(2) = 0
chkcomck = 0
atkingtrn(1) = 0
atkingtrn(2) = 0
'=====
If turnatk = 1 Then
 �԰��t����.chkdefcom
End If
'======================================
Erase Vss_EventPlayerAllActionOffNum
'===========================���涥�q���J�I(ATK-17/DEF-37)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 17, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 37, 2
'===========================���涥�q���J�I(ATK-92/DEF-93)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 92, 4
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 93, 4
'==============
�p�H���Y�����ʤ�V��(1) = 1
�p�H���Y�����ʤ�V��(2) = 2
�p�H���Y������_�ϥΪ�.Enabled = True
�p�H���Y������_�q��.Enabled = True
'==============
�@��t����.���ļ��� 6
�԰��t����.�ɶ��b_���]
trtimeline.Enabled = True
trgoi2.Enabled = True
'==============
�԰��t����.��q��s���
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
End Sub



Private Sub cn22_Click()
cn22.Visible = False
OK���s�P���������ˬd.Enabled = True
End Sub

Sub cn3_Click()
'======================
If �q����ƥ�d�O�_�X����ܼ� = True Then
    GoTo �q����ƥ�d���X���_���涥�q����
End If
'======================
If moveturn = 2 Then
  If �t����ܬɭ������� = 1 Then
        move1.Visible = True
        move2.Visible = False
        atkdef1.Visible = True
        atkdef2.Visible = False
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\def2.gif"
  End If
  ��ܦC1.goi1��� = True
  ��ܦC1.goi2��� = True
  ��ܦC1.���ʶ��q��ܭ� = 0
  ��ܦC1.���ʶ��q����� = False
Else
  If �t����ܬɭ������� = 1 Then
        atkdef1.Visible = False
        atkdef2.Visible = True
  Else
        FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\def2.gif"
  End If
End If
turnatk = 2
��ܦC1.��ܦC�Ϥ� = app_path & "gif\system\lineusdef.png"
�԰��t����.cleanatkingpagetot
'===============
��ܦC1.goi1 = 0
��ܦC1.goi2 = 0
�������m��l�`��(1) = 0
�������m��l�`��(2) = 0
�������m��l�`��(3) = 0
�������m��l�`��(4) = 0
��ƹs�ˬd��(1) = False
��ƹs�ˬd��(2) = False
�O�_�t�Τ��� = False
'=====
�ثe��(6) = 0
�ثe��(21) = 1
'===============
goicheck(1) = 0
goicheck(2) = 0
atkingtrn(1) = 0
atkingtrn(2) = 0
If turnatk = 2 Then
 �԰��t����.chkdef
End If
'==============
�԰��t����.��q��s���
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'============================
Erase Vss_EventPlayerAllActionOffNum
'===========================���涥�q���J�I(ATK-17/DEF-37)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 17, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 37, 2
'===========================���涥�q���J�I(ATK-92/DEF-93)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 92, 4
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 93, 4
'======================�q����ƥ�d���X���
If �q����ƥ�d�O�_�X����ܼ� = False Then
   GoTo �q����ƥ�d���X���_���涥�q2
End If
'================================
�q����ƥ�d���X���_���涥�q����:
'----------�H�U���q���P�_�X�P�{���X�]������^
'====================���紼�z��AI�X�P�t��
Dim wtyr As Integer '�Ȯ��ܼ�
If moveturn = 1 Then wtyr = 1 Else wtyr = 0
���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_��� 2, 1, namecom(����H����ԤH��(2, 2)), movecp, wtyr
GoTo ���z��AI�X�P_���涥�q����
 '==================
If turnatk = 2 And movecp = 1 Then
   �԰��t����.comatk1
ElseIf turnatk = 2 And movecp > 1 Then
   �԰��t����.comatk2
End If
'==============================
���z��AI�X�P_���涥�q����:
'==============================
�q����ƥ�d���X���_���涥�q2:
If �q����ƥ�d�O�_�X����ܼ� = False Then
    '==========
    cn3.Visible = False
    �ثe��(6) = 0
    �ثe��(17) = 1
    �ثe��(15) = 0
    '==============
    �p�H���Y�����ʤ�V��(1) = 2
    �p�H���Y�����ʤ�V��(2) = 1
    �p�H���Y������_�ϥΪ�.Enabled = True
    �p�H���Y������_�q��.Enabled = True
    '==============
    �԰��t����.�ɶ��b_���]
    trtimeline.Enabled = True
ElseIf �q����ƥ�d�O�_�X����ܼ� = True Then  '�q����ƥ�d���X���_�����ᶥ�q2
    �q���X�P.Enabled = True
End If
End Sub

Private Sub cn32_Click()
cn32.Visible = False
OK���s�P���������ˬd.Enabled = True
End Sub

Private Sub cn4_Click()
Dim uscomvsn As Integer
cn4.Visible = False
turnatk = 5
If moveturn = 1 Then uscomvsn = 2 Else uscomvsn = 1
'===========================���涥�q���J�I(50)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 50, 1
'============================
'===========================���涥�q���J�I(51)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 51, 1
'============================
'===========================���涥�q���J�I(52)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 52, 1
'============================
HP�ˬd���q�� = 4
�԰��t����.����HP�ˬd
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
For i = 1 To cMusicPlayer.UBound
    Unload cMusicPlayer(i)
Next
End Sub

Private Sub NextTurn_���q2_Timer()
Dim uscomvsn As Integer
Dim i As Integer, j As Integer, k As Integer
goidefus = 0
'======�H�U���~�P�{���X
If BattleCardNum < �P�`���q��(1) + �P�`���q��(2) Then
    �԰��t����.����ʧ@_�~�P
End If
'==========================
If moveturn = 1 Then uscomvsn = 2 Else uscomvsn = 1
'===========================���涥�q���J�I(53)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 53, 1
'============================
'===========================���涥�q���J�I(54)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 54, 1
'============================
'===========================���涥�q���J�I(55)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 55, 1
'============================
�԰��t����.�s���T�� BattleTurn & "�^�X�����C"
'=============
NextTurn_���q2.Enabled = False
'=============
If �԰��t����.����HP�ˬd_�����^�X�ˬd = True Then
    Exit Sub
End If
'==============�N�C�^�X���Ұʦ����k�s
For i = 1 To 2
   For j = 1 To 3
       For k = 1 To 8
           atkingck(i, j, k, 2) = 0
       Next
    Next
Next
'==============
BattleTurn = BattleTurn + 1
PEAFInterface.turn = BattleTurn
��ܦC1.goi1��� = False
��ܦC1.goi2��� = False
��ܦC1.goi1 = 0
��ܦC1.goi2 = 0
�������m��l�`��(1) = 0
�������m��l�`��(2) = 0
'====================
If �t����ܬɭ������� = 1 Then
    move1.Visible = True
    move2.Visible = False
    atkdef1.Visible = False
    atkdef2.Visible = False
    move3.Picture = LoadPicture(app_path & "gif\system\move3.gif")
    move4.Picture = LoadPicture(app_path & "gif\system\move4.gif")
Else
    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\stageblack.gif"
End If
��ܦC1.��ܦC�Ϥ� = app_path & "gif\system\DRAW.png"
'==============
�p�H���Y�����ʤ�V��(1) = 2
�p�H���Y�����ʤ�V��(2) = 2
�p�H���Y������_�ϥΪ�.Enabled = True
�p�H���Y������_�q��.Enabled = True
'==============
���ݮɶ���C(2).Add 1
���ݮɶ�_2.Enabled = True
End Sub

Private Sub OK���s�P���������ˬd_Timer()
If �ϥΪ̥X�P_�X�P���_�a��.Enabled = False And �ϥΪ̥X�P_�X�P���_�a�k.Enabled = False And �ϥΪ̥X�P_��P���.Enabled = False And ��������ˬd.Enabled = False Then
   OK���s�P���������ˬd.Enabled = False
   turnpageonin = 0
   Select Case turnatk
       Case 1
           �������q_���q��l.Enabled = True
       Case 2
           ���m���q_���q��l.Enabled = True
       Case 3
           ���ʶ��q_���q��l.Enabled = True
   End Select
End If
End Sub

Private Sub pagecomglead_Change()
pageglead(2) = Val(pagecomglead.Caption)
End Sub

Private Sub pageusglead_Change()
pageglead(1) = Val(pageusglead.Caption)
End Sub

Private Sub PEAEtr1_Timer()
Select Case PEAEtr1num
    Case 10
         If �԰��Ҧ��ӱѬ����� = 1 Then
             FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(app_path & "gif\system\gamewin.jpg")
         ElseIf �԰��Ҧ��ӱѬ����� = 2 Then
             FormMainMode.PEAttackingEndingForm.Picture = LoadPicture(app_path & "gif\system\gamelose.jpg")
         ElseIf �԰��Ҧ��ӱѬ����� = 3 Then
         
         End If
         FormMainMode.cMusicPlayer(0).MusicPlay
    Case 50
         PEAEtr1.Enabled = False
         '======================
         �԰��t����.�C����Ե����������
         '======================
         If Formsetting.chkautocontinuemode.Value = 1 Then
            bnreturnt_Click
         End If
         bnreturn.Visible = True
         bnreturnt.Visible = True
         bn.Visible = True
         bnt.Visible = True
End Select
PEAEtr1num = PEAEtr1num + 1
End Sub

Private Sub PEAFAnimateInterface_AnimateCheckPoint(ByVal uscom As Integer)
Vss_AtkingStartPlayNum(2) = 1 '�ޯ���椤�Ұ�
End Sub

Private Sub PEAFAnimateInterface_AnimateEnd(ByVal uscom As Integer)
Vss_AtkingStartPlayNum(3) = 1
End Sub

Private Sub PEAFInterface_ActiveMouseEnter(ByVal uscom As Integer, ByVal num As Integer)
Dim i As Integer
Dim tmpobj As clsPersonActiveSkill

Select Case uscom
 Case 1
    For i = 1 To �԰��t����.ActionCardTotNum
       card(i).CardEventType = False
    Next
 Case 2
    For i = 1 To 3
      cardcom(i).Visible = False
    Next
End Select
'============================
Set tmpobj = �԰��t����.ActiveSkillObj(uscom, num)
If Not tmpobj Is Nothing Then
    PEAFatkinghelpc.Stage = tmpobj.Stage
    PEAFatkinghelpc.Distance = tmpobj.Distance
    PEAFatkinghelpc.card = tmpobj.card
    PEAFatkinghelpc.Effect = tmpobj.Effect
    PEAFatkinghelpc.Left = atkinghelpxy(uscom, num, 1)
    PEAFatkinghelpc.Top = atkinghelpxy(uscom, num, 2)
    PEAFatkinghelpc.ZOrder
    PEAFatkinghelpc.Visible = True
End If
End Sub

Private Sub PEAFInterface_ActiveMouseExit(ByVal uscom As Integer, ByVal num As Integer)
PEAFatkinghelpc.Visible = False
End Sub

Sub PEAFInterface_BnOKClick()
Dim i As Integer
If turnpageonin = 1 Then
    turnpageonin = 0
    For i = 1 To �԰��t����.ActionCardTotNum
        FormMainMode.card(i).card_MouseExit
    Next
    FormMainMode.PEAFInterface.BnOKStopListen
    �԰��t����.�ɶ��b_����
    Select Case turnatk
        Case 1
            ���ݮɶ���C(1).Add 7
            ���ݮɶ�.Enabled = True
        Case 2
            ���ݮɶ���C(1).Add 8
            ���ݮɶ�.Enabled = True
        Case 3
            cnmove2_Click
    End Select
End If
End Sub

Private Sub PEAFInterface_BnOKMouseMove()
Dim i As Integer
For i = 1 To �԰��t����.ActionCardTotNum
   card(i).CardEventType = False
Next
End Sub

Private Sub PEAFInterface_InterfaceMouseMove()
Dim i As Integer
For i = 1 To �԰��t����.ActionCardTotNum
   card(i).CardEventType = False
Next
For i = 1 To 3
  cardcom(i).Visible = False
Next
For i = 1 To 3
  If i <> ����H����ԤH��(1, 2) Then
     cardus(i).Visible = False
  End If
Next
End Sub

Private Sub PEAFpersoncardcom_MouseEnter(Index As Integer)
cardcom(Index).Left = FormMainMode.PEAFpersoncardcom(Index).Left
cardcom(Index).Top = 480
cardcom(Index).Visible = True
cardcom(Index).ZOrder
Select Case Index
   Case 1
      cardcom(2).Visible = False
      cardcom(3).Visible = False
   Case 2
      cardcom(1).Visible = False
      cardcom(3).Visible = False
   Case 3
      cardcom(2).Visible = False
      cardcom(1).Visible = False
End Select
End Sub

Private Sub PEAFpersoncardus_MouseEnter(Index As Integer)
cardus(Index).Left = FormMainMode.PEAFpersoncardus(Index).Left
cardus(Index).Top = 5760
cardus(Index).Visible = True
cardus(Index).ZOrder
Select Case Index
   Case 1
      If ����H����ԤH��(1, 2) = 2 Then
          cardus(3).Visible = False
      Else
          cardus(2).Visible = False
      End If
   Case 2
      If ����H����ԤH��(1, 2) = 1 Then
          cardus(3).Visible = False
      Else
          cardus(1).Visible = False
      End If
   Case 3
      If ����H����ԤH��(1, 2) = 2 Then
          cardus(1).Visible = False
      Else
          cardus(2).Visible = False
      End If
End Select
End Sub

Private Sub PEASpke_Timer()
If swq = 35 Then
    PEASpke.Enabled = False
    If PEASpersontalk.��ܤ�r <> "" Then
        PEASpersontalk.��ܤ�r��� = True
    End If
ElseIf swq = 10 Then
    PEASpersontalk.Top = -120
    PEASpersontalk.��ܤ�r = �H���t����.�H����ܿ��
    If PEASpersontalk.��ܤ�r <> "" Then
        PEASpersontalk.Visible = True
        PEASpersontalk.��ܤ�r��� = False
        PEASpersontalk.ZOrder
    End If
    swq = Val(swq) + 1
Else
    swq = Val(swq) + 1
End If

End Sub

Private Sub PEAttackingForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 1 To �԰��t����.ActionCardTotNum
   card(i).CardEventType = False
Next
For i = 1 To 3
  cardcom(i).Visible = False
Next
For i = 1 To 3
  If i <> ����H����ԤH��(1, 2) Then
     cardus(i).Visible = False
  End If
Next
End Sub

Private Sub tr1_Timer()
Select Case tr1num
    Case 1
'        MsgBox "1-1"
        PEStext1.Visible = True
    Case 3
        If �Ĥ@���Ұ�Ū�J�{�ǼаO = False Then
'            �@��t����.�C����lŪ�J�{��
            �Ĥ@���Ұ�Ū�J�{�ǼаO = True
            ����Ū�J���� = "PEGF"   '====���ն��q-�����i�J�ۥѾ԰��Ҧ�
'            MsgBox "1-3"
        End If
    Case 5
        Select Case ����Ū�J����
            Case "PEGF"
'                MsgBox "1-5"
                �@��t����.�C����lŪ�J�{��
                �@��t����.�ۥѾ԰��Ҧ��]�w���Ū�J�{��
                �@��t����.�ۥѾ԰��Ҧ��]�w���򥻳]�w�{��
        End Select
    Case 7
        Select Case ����Ū�J����
            Case "PEGF"
'                MsgBox "1-7"
                �@��t����.�D���_PEGameFreeModeSettingForm���
        End Select
        tr1.Enabled = False
        PEStartForm.Visible = False
End Select
tr1num = tr1num + 1
End Sub

Private Sub trend_Timer()
If trend�Ȯ��ܼ� = 4 Then
   �@��t����.�D���_PEAttackingEndingForm���
   PEAttackingForm.Visible = False
   PEAEtr1num = 0
   PEAEtr1.Enabled = True
   trend.Enabled = False
ElseIf trend�Ȯ��ܼ� = 2 Then
   FormMainMode.cMusicPlayer(0).MusicStop
   FormMainMode.cMusicPlayer(0).IsLoop = False
   FormMainMode.cMusicPlayer(0).Filepath = app_path & "mp3\ulse15.mp3"
   trend�Ȯ��ܼ� = trend�Ȯ��ܼ� + 1
Else
   trend�Ȯ��ܼ� = trend�Ȯ��ܼ� + 1
End If
End Sub

Sub trgoi1_Timer()
'=========��s��l�`�ƶq���
If �������m��l�`��(1) < 0 Then
   ��ܦC1.goi1 = 0
Else
   ��ܦC1.goi1 = �������m��l�`��(1)
End If
FormMainMode.trgoi1.Enabled = False
'=====================
End Sub

Sub trgoi2_Timer()
'=========��s��l�`�ƶq���
If �������m��l�`��(2) < 0 Then
   ��ܦC1.goi2 = 0
Else
   ��ܦC1.goi2 = �������m��l�`��(2)
End If
trgoi2.Enabled = False

End Sub

Private Sub trnextend_Timer()
Select Case Val(�Y���淾�q�Ȯ��ܼ�(3))
   Case 1
      �ˮ`����_�ϥΪ� (Val(�Y���淾�q�Ȯ��ܼ�(2)))
   Case 2
      �ˮ`����_�q�� (Val(�Y���淾�q�Ȯ��ܼ�(2)))
End Select
'=============
���ݮɶ���C(2).Add 21
���ݮɶ�_2.Enabled = True
trnextend.Enabled = False
End Sub

Private Sub trtimeline_Timer()
Dim i As Integer

timelineout1.X1 = timelineout1.X1 + 2
timelineout2.X2 = timelineout2.X2 - 2
For i = 1 To 3
   �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, i) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, i) + 2
Next
Select Case timelineout1.X1
   Case Is <= 2624
       If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 1) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) + 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) + 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) + 1
       End If
       If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1
       End If
       If timelineout1.X1 >= 2624 Then
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 1) = 34
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) = 13
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 3) = 60
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) = 217
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = 217
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) = 50
            timelineout1.BorderColor = RGB(217, 217, 50)
            timelineout2.BorderColor = RGB(217, 217, 50)
        End If
   Case Is <= 3936
        If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 1) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) + 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) + 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) + 1
       End If
       If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1
       End If
       If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 3) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) - 1)
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) - 1)
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) - 1
       End If
       If timelineout1.X1 >= 3936 Then
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 1) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) = 11
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 3) = 47
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) = 255
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = 118
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) = 28
            timelineout1.BorderColor = RGB(255, 118, 28)
            timelineout2.BorderColor = RGB(255, 118, 28)
            '=========�ɶ��b(�~)
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 1) = 1
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 2) = 0
            �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = 0
            timelinein1.BorderColor = RGB(0, 0, 0)
            timelinein2.BorderColor = RGB(0, 0, 0)
        End If
    Case Is > 3936
       If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1, �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3))
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) - 1
       End If
       If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) >= �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 3) Then
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) = 0
           timelineout1.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) - 1)
           timelineout2.BorderColor = RGB(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2), �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) - 1)
           �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) = �ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) - 1
       End If
       '===================�ɶ��b(�~)
       Select Case �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 1)
           Case 1
                    If 255 - Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) < 9 Then
                       timelinein1.BorderColor = RGB(255, 0, 0)
                       timelinein2.BorderColor = RGB(255, 0, 0)
                       �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = 255
                    Else
                       timelinein1.BorderColor = RGB(Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) + 9, 0, 0)
                       timelinein2.BorderColor = RGB(Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) + 9, 0, 0)
                       �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) + 9
                    End If
                If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = 255 Then
                    �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 1) = 2
                End If
           Case 2
               If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) < 9 Then
                   timelinein1.BorderColor = RGB(0, 0, 0)
                   timelinein2.BorderColor = RGB(0, 0, 0)
                   �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = 0
                Else
                   timelinein1.BorderColor = RGB(Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) - 9, 0, 0)
                   timelinein2.BorderColor = RGB(Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) - 9, 0, 0)
                   �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = Val(�ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3)) - 9
                End If
                If �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 3) = 0 Then
                    �ɶ��b�C���ܤƬ����Ȯ��ܼ�(4, 1) = 1
                End If
       End Select
End Select
If timelineout1.X1 >= timelineout1.X2 Then
    �԰��t����.�ɶ��b_����
    turnpageonin = 0
    FormMainMode.PEAFInterface.BnOKStopListen
    ���ݮɶ���C(2).Add 4
    ���ݮɶ�_2.Enabled = True
End If
End Sub

Private Sub tr�ϥΪ�_��P_Timer()
�԰��t����.����ʧ@_�ϥΪ�_��P �ثe��(20)
tr�ϥΪ�_��P.Enabled = False
End Sub

Private Sub tr�ϥΪ̵P_���P_Timer()
�԰��t����.����ʧ@_�ϥΪ̵P_���P_�q�� �ثe��(20)
tr�ϥΪ̵P_���P.Enabled = False
End Sub

Private Sub tr�P��_�^�P_�ϥΪ�_Timer()
card(�ثe��(16)).Left = 240
card(�ثe��(16)).Top = 960
card(�ثe��(16)).Visible = True
�԰��t����.����ʧ@_�P��_�^�P_�ϥΪ� �ثe��(16)
tr�P��_�^�P_�ϥΪ�.Enabled = False
End Sub

Sub tr�P��_�^�P_�q��_Timer()
card(�ثe��(16)).Left = 240
card(�ثe��(16)).Top = 960
card(�ثe��(16)).Visible = True
�԰��t����.����ʧ@_�P��_�^�P_�q�� �ثe��(16)
tr�P��_�^�P_�q��.Enabled = False
End Sub


Private Sub tr�P��_��P_�ϥΪ�_Timer()
tr�P��_��P_�ϥΪ�.Enabled = False
If BattleCardNum > 0 Then
    �԰��t����.����ʧ@_��P_���εP 1
End If
End Sub

Private Sub tr�P��_��P_�q��_Timer()
tr�P��_��P_�q��.Enabled = False
If BattleCardNum > 0 Then
    �԰��t����.����ʧ@_��P_���εP 2
End If
End Sub

Private Sub tr�q���P_���P_Timer()
�԰��t����.����ʧ@_�q���P_���P_�ϥΪ� �ثe��(16)
tr�q���P_���P.Enabled = False
End Sub

Private Sub tr�q���P_��P_Timer()
�԰��t����.����ʧ@_�q��_��P �ثe��(16)
tr�q���P_��P.Enabled = False
End Sub

Private Sub tr�q���P_½�P_Timer()
    �԰��t����.����ʧ@_½�P �ثe��(16)
    tr�q���P_½�P.Enabled = False
    If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDestroyCards") <> 0 Then
        vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDestroyCards")) = 2 '(���q2)
    End If
    If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGiveCards") <> 0 Then
        vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGiveCards")) = 2 '(���q2)
    End If
    If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards") <> 0 Then
        vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards")) = 2 '(���q2)
    End If
   '=======================�H�U�O�ƥ�d�ˬd�αҰ�
   If �ƥ�d�O���Ȯɼ�(1, 5) = 2 And �ƥ�d�O���Ȯɼ�(1, 6) = 1 Then
        �ƥ�d�O���Ȯɼ�(1, 3) = 4
        �ƥ�d.�A�G�N_�ϥΪ� 0, 0 '==�ƥ�d����_�A�G�N_�ϥΪ�(���q4)
   End If
End Sub

Private Sub �H�������ˬd_Timer()
If �H�������ˬd�Ȯ��ܼ�(1) = 10 Then
    If �H�������ˬd�Ȯ��ܼ�(2) = 1 Then
        personusminijpg.�p�H������ = True
    End If
    If �H�������ˬd�Ȯ��ܼ�(3) = 1 Then
        personcomminijpg.�p�H������ = True
    End If
    �H�������ˬd�Ȯ��ܼ�(1) = Val(�H�������ˬd�Ȯ��ܼ�(1)) + 1
ElseIf Val(�H�������ˬd�Ȯ��ܼ�(1)) > 10 And personcomminijpg.�p�H������ = False And personusminijpg.�p�H������ = False Then
    �H�������ˬd.Enabled = False
    FormMainMode.���ݮɶ�.Enabled = True
Else
    �H�������ˬd�Ȯ��ܼ�(1) = Val(�H�������ˬd�Ȯ��ܼ�(1)) + 1
End If
End Sub

Private Sub �p�H���Y������_�ϥΪ�_Timer()
Dim pnm As Integer
If ��ܦC1.�ϥΪ̤�p�H���Ϥ�width > 1440 Then
    pnm = 0
Else
    pnm = 1440 - ��ܦC1.�ϥΪ̤�p�H���Ϥ�width
End If
Select Case �p�H���Y�����ʤ�V��(1)
    Case 1
        If ��ܦC1.�ϥΪ̤�p�H���Ϥ�left >= pnm Then
           ��ܦC1.�ϥΪ̤�p�H���Ϥ�left = pnm
           �԰��t����.�p�H���Y�����槹�P�__�ϥΪ�
           �p�H���Y������_�ϥΪ�.Enabled = False
           Exit Sub
        End If
           ��ܦC1.�ϥΪ̤�p�H���Ϥ�left = ��ܦC1.�ϥΪ̤�p�H���Ϥ�left + 100
        If ��ܦC1.�ϥΪ̤�p�H���Ϥ�left >= pnm Then
           ��ܦC1.�ϥΪ̤�p�H���Ϥ�left = pnm
           �p�H���Y������_�ϥΪ�.Enabled = False
           �԰��t����.�p�H���Y�����槹�P�__�ϥΪ�
        End If
    Case 2
        If ��ܦC1.�ϥΪ̤�p�H���Ϥ�left <= -��ܦC1.�ϥΪ̤�p�H���Ϥ�width Then
           ��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -��ܦC1.�ϥΪ̤�p�H���Ϥ�width
           �p�H���Y������_�ϥΪ�.Enabled = False
           Exit Sub
        End If
           ��ܦC1.�ϥΪ̤�p�H���Ϥ�left = ��ܦC1.�ϥΪ̤�p�H���Ϥ�left - 100
        If ��ܦC1.�ϥΪ̤�p�H���Ϥ�left <= -��ܦC1.�ϥΪ̤�p�H���Ϥ�width Then
           ��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -��ܦC1.�ϥΪ̤�p�H���Ϥ�width
           �p�H���Y������_�ϥΪ�.Enabled = False
        End If
End Select
End Sub

Private Sub �p�H���Y������_�q��_Timer()
Dim pnm As Integer
If ��ܦC1.�q����p�H���Ϥ�width > 1440 Then
    pnm = FormMainMode.ScaleWidth - ��ܦC1.�q����p�H���Ϥ�width
Else
    pnm = FormMainMode.ScaleWidth - 1440
End If
Select Case �p�H���Y�����ʤ�V��(2)
    Case 1
        If ��ܦC1.�q����p�H���Ϥ�left <= pnm Then
           ��ܦC1.�q����p�H���Ϥ�left = pnm
           �԰��t����.�p�H���Y�����槹�P�__�q��
           �p�H���Y������_�q��.Enabled = False
           Exit Sub
        End If
           ��ܦC1.�q����p�H���Ϥ�left = ��ܦC1.�q����p�H���Ϥ�left - 100
        If ��ܦC1.�q����p�H���Ϥ�left <= pnm Then
           ��ܦC1.�q����p�H���Ϥ�left = pnm
           �p�H���Y������_�q��.Enabled = False
           �԰��t����.�p�H���Y�����槹�P�__�q��
        End If
    Case 2
        If ��ܦC1.�q����p�H���Ϥ�left >= FormMainMode.ScaleWidth Then
           ��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
           �p�H���Y������_�q��.Enabled = False
           Exit Sub
        End If
           ��ܦC1.�q����p�H���Ϥ�left = ��ܦC1.�q����p�H���Ϥ�left + 100
        If ��ܦC1.�q����p�H���Ϥ�left >= FormMainMode.ScaleWidth Then
           ��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
           �p�H���Y������_�q��.Enabled = False
        End If
End Select
End Sub

Private Sub ���P���q_�p��_Timer()
Select Case �ثe��(10)
    Case 1
       �԰��t����.���P�p��Z�����_�ϥΪ�
       ���P���q_�p��.Enabled = False
       �ثe��(11) = 0
       �ثe��(12) = pageqlead(�ثe��(10)) - 1
       �P����_���P.Enabled = True
    Case 2
       �԰��t����.���P�p��Z�����_�q��
       ���P���q_�p��.Enabled = False
       �ثe��(11) = 0
       �ثe��(12) = pageqlead(�ثe��(10)) - 1
       �P����_���P.Enabled = True
    Case 3
       ���P���q_�p��.Enabled = False
       Select Case turnatk
          Case 1
             �԰��t����.����HP�ˬd
          Case 2
             �԰��t����.����HP�ˬd
          Case 3
             '===========================���涥�q���J�I(8)
              ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l moveturn, 8, 1
            '============================
             HP�ˬd���q�� = 1
             �԰��t����.����HP�ˬd
       End Select
End Select
End Sub

Private Sub ��q���J�ʵe_Timer()
If ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 2) = 0 Then
    If bloodlineout1.Width >= 5295 Then
        ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 2) = 1
    ElseIf 5295 - bloodlineout1.Width <= 106 Then
        ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 1) = 5295 - bloodlineout1.Width
        bloodlineout1.Width = bloodlineout1.Width + ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 1)
        ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 2) = 1
    Else
       bloodlineout1.Width = bloodlineout1.Width + ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 1)
    End If
End If
If ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 2) = 0 Then
    If bloodlineout2.Left <= 6060 Then
        ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 2) = 1
    ElseIf bloodlineout2.Left - 6060 <= 106 Then
        ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 1) = bloodlineout2.Left - 6060
        bloodlineout2.Left = bloodlineout2.Left - ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 1)
        ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 2) = 1
    Else
        bloodlineout2.Left = bloodlineout2.Left - ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 1)
    End If
End If
If ��q�p�ƾ��ʵe�Ȯ��ܼ�(1, 2) = 1 And ��q�p�ƾ��ʵe�Ȯ��ܼ�(2, 2) = 1 Then
   ��q���J�ʵe.Enabled = False
   ���ݮɶ���C(2).Add 1
   ���ݮɶ�_2.Enabled = True
End If
End Sub


Private Sub �������q_���q1_Timer()
'======================
If �q����ƥ�d�O�_�X����ܼ� = True Then
    GoTo �q����ƥ�d���X���_���涥�q����
End If
'======================�q����ƥ�d���X���
If �q����ƥ�d�O�_�X����ܼ� = False Then
    GoTo �q����ƥ�d���X���_���涥�q2
End If
'================================
�q����ƥ�d���X���_���涥�q����:
'====================���紼�z��AI�X�P�t��
Dim wtyr As Integer '�Ȯ��ܼ�
If moveturn = 2 Then wtyr = 1 Else wtyr = 0
���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_��� 2, 2, namecom(����H����ԤH��(2, 2)), movecp, wtyr
'================
�q����ƥ�d���X���_���涥�q2:
'================
�������q_���q1.Enabled = False
If �q����ƥ�d�O�_�X����ܼ� = False Then
    �ثe��(6) = 0
    �ثe��(17) = 1
    �ثe��(15) = 0
    �p�H���Y�����ʤ�V��(1) = 2
    �p�H���Y�����ʤ�V��(2) = 1
    �p�H���Y������_�ϥΪ�.Enabled = True
    �p�H���Y������_�q��.Enabled = True
End If
'======================�q����ƥ�d���X���_�����ᶥ�q2
If �q����ƥ�d�O�_�X����ܼ� = True Then
    �q���X�P.Enabled = True
End If
'===========================
End Sub

Private Sub �������q_���q2_Timer()
'----------�H�U�������Ҧ��{��
�Y���淾�q�Ȯ��ܼ�(2) = 0
�Y���淾�q�Ȯ��ܼ�(3) = 0
�Y���淾�q�Ȯ��ܼ�(5) = 0
�Y���淾�q�Ȯ��ܼ�(6) = 0
�Y���淾�q�Ȯ��ܼ�(7) = 0
�Y���淾�q�Ȯ��ܼ�(8) = 0
'==============================
HP�ˬd�ܼ� = False
'==============================
�԰��t����.��q��s���
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'===========================���涥�q���J�I(ATK-10/DEF-30)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 10, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 30, 2
'============================
'===========================���涥�q���J�I(ATK-11/DEF-31)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 11, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 31, 2
'============================
'=================
If �������m��l�`��(1) <= 0 Then
  �԰��t����.�s���T�� "�S�������C"
  �԰��t����.�s���T�� "�z�����F�����C"
  ��ƹs�ˬd��(1) = True
Else
  �԰��t����.�s���T�� "�M�w�����O" & �������m��l�`��(1) & "�I�C"
End If
If �������m��l�`��(2) <= 0 Then
   ��ƹs�ˬd��(2) = True
End If
'===========================���涥�q���J�I(ATK-12/DEF-32)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 12, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 32, 2
'============================
���q���A�� = 2
�������q_���q2.Enabled = False
'============================
HP�ˬd�ܼ� = True
HP�ˬd���q�� = 2
�ثe��(10) = 1
���P���q_�p��.Enabled = True
End Sub

Private Sub �������q_���q��l_Timer()
�԰��t����.�ɶ��b_���]
trtimeline.Enabled = True
�q����ƥ�d�O�_�X����ܼ� = False
'==============================
�԰��t����.��q��s���
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'==============================
�������q_���q��l.Enabled = False
�������q_���q1.Enabled = True
End Sub

Private Sub ���m���q_���q��l_Timer()
'----------�H�U�����m�Ҧ��{��
�Y���淾�q�Ȯ��ܼ�(2) = 0
�Y���淾�q�Ȯ��ܼ�(3) = 0
�Y���淾�q�Ȯ��ܼ�(5) = 0
�Y���淾�q�Ȯ��ܼ�(6) = 0
�Y���淾�q�Ȯ��ܼ�(7) = 0
�Y���淾�q�Ȯ��ܼ�(8) = 0
'====================
HP�ˬd�ܼ� = False
'==============================
�԰��t����.��q��s���
FormMainMode.trgoi1_Timer
FormMainMode.trgoi2_Timer
'===========================���涥�q���J�I(ATK-10/DEF-30)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 10, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 30, 2
'============================
'===========================���涥�q���J�I(ATK-11/DEF-31)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 11, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 31, 2
'============================
If �������m��l�`��(2) <= 0 Then
  �԰��t����.�s���T�� "�S�������C"
  �԰��t����.�s���T�� "�z���������F�����C"
  ��ƹs�ˬd��(2) = True
Else
  �԰��t����.�s���T�� "�M�w�����O" & �������m��l�`��(2) & "�I�C"
End If
If �������m��l�`��(1) <= 0 Then
   ��ƹs�ˬd��(1) = True
End If
'===========================���涥�q���J�I(ATK-12/DEF-32)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 12, 2
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 32, 2
'============================
���q���A�� = 4
���m���q_���q��l.Enabled = False
'============================
HP�ˬd�ܼ� = True
HP�ˬd���q�� = 2
�ثe��(10) = 1
���P���q_�p��.Enabled = True
End Sub

Sub �ϥΪ̥X�P_AI�X�P����_Timer()
Dim tmpcard As clsActionCard
If turnpageonin = 1 And �P����.Enabled = False Then
    If �԰��t����.CardDeckCollection(5).Count > 0 Then
        Set tmpcard = �԰��t����.CardDeckCollection(5)(�ثe��(32))
        If tmpcard.ComMark = 3 Then
            �ثe��(32) = �ثe��(32) - 1
            FormMainMode.card_CardClick tmpcard.CardNum
        End If
        �ثe��(32) = �ثe��(32) + 1
    End If
    If �ثe��(32) > �԰��t����.CardDeckCollection(5).Count Then
        �ϥΪ̥X�P_AI�X�P����.Enabled = False
        ���ݮɶ���C(1).Add 37
        ���ݮɶ�.Enabled = True
    End If
End If
End Sub

Sub �ϥΪ̥X�P_AI�X�P����_�ƥ�d_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

If turnpageonin = 1 And �P����.Enabled = False Then
    If �԰��t����.CardDeckCollection(5).Count > 0 Then
        For i = 1 To �԰��t����.CardDeckCollection(5).Count
            Set tmpcard = �԰��t����.CardDeckCollection(5)(i)
            If tmpcard.CardType = 2 Then
                If tmpcard.UpperType = a6a Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
                If tmpcard.UpperType = a7a And (turnatk = 1 Or turnatk = 2) Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
                If tmpcard.UpperType = a8a Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
                If tmpcard.UpperType = a9a Then
                    FormMainMode.card_CardClick tmpcard.CardNum
                    Exit Sub
                End If
            End If
        Next
    End If
    
    �ϥΪ̥X�P_AI�X�P����_�ƥ�d.Enabled = False
    ���ݮɶ���C(2).Add 46
    ���ݮɶ�_2.Enabled = True
End If
End Sub


Private Sub �ϥΪ̥X�P_��P���_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To �԰��t����.CardDeckCollection(5).Count
    If i >= �ثe��(5) Then
        Set tmpcard = �԰��t����.CardDeckCollection(5)(i)
        If �ثe��(13) = 0 Then
            If card(tmpcard.CardNum).Left = 2640 And card(tmpcard.CardNum).Top = 7980 Then  '���w��2�C��1�i�P
                �ثe��(13) = tmpcard.CardNum
                tmpcard.XYLeft = card(�ثe��(13)).Left  '���w�ثeLeft(�y��)
                tmpcard.XYTop = card(�ثe��(13)).Top  '���w�ثeTop(�y��)
                '==========�԰��t����.�p��P���ʶZ�����
                �Z�����_���P�Ȯɼ�(1, 1) = (9840 - tmpcard.XYLeft) \ 10 '�p��Left
                �Z�����_���P�Ȯɼ�(1, 2) = -((tmpcard.XYTop - 6700) \ 10)  '�p��Top
            End If
        End If
        If �ثe��(13) = tmpcard.CardNum Then
           card(�ثe��(13)).Left = card(�ثe��(13)).Left + �Z�����_���P�Ȯɼ�(1, 1)
           card(�ثe��(13)).Top = card(�ثe��(13)).Top + �Z�����_���P�Ȯɼ�(1, 2)
        Else
           card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (900 / 10)
        End If
  End If
Next
�ثe��(4) = �ثe��(4) + (900 / 10)
If �ثe��(4) >= 900 Then
    �ϥΪ̥X�P_��P���.Enabled = False
    Select Case �ثe��(21)
        Case 1
            '======�����ʧ@
        Case 2
             If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards")) = 2 '(���q2)
            End If
       Case 3
           '===========�ƥ�d����_�A�G�N_�q��(���q3)
            �ƥ�d�O���Ȯɼ�(2, 3) = 3
            �ƥ�d.�A�G�N_�q�� 0, 0
       Case 4
            If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDestroyCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDestroyCards")) = 2 '(���q2)
            End If
       Case 5
             If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGiveCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGiveCards")) = 2 '(���q2)
            End If
        Case 11
            ���ݮɶ���C(2).Add 38
            ���ݮɶ�_2.Enabled = True
    End Select
End If
End Sub

Private Sub �ϥΪ̥X�P_�X�P���_�a�k_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To �԰��t����.CardDeckCollection(6).Count
    Set tmpcard = �԰��t����.CardDeckCollection(6)(i)
    If i < �ثe��(5) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left + (480 / 10)
    End If
    If i >= �ثe��(5) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (500 / 10)
    End If
Next
�ثe��(3) = �ثe��(3) + (480 / 10)
If �ثe��(3) >= 480 Then
    �ϥΪ̥X�P_�X�P���_�a�k.Enabled = False
End If
End Sub

Private Sub �ϥΪ̥X�P_�X�P���_�a��_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To (�԰��t����.CardDeckCollection(6).Count - 1)
    Set tmpcard = �԰��t����.CardDeckCollection(6)(i)
    card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (480 / 10)
Next
�ثe��(3) = �ثe��(3) + (480 / 10)
If �ثe��(3) >= 480 Then
    �ϥΪ̥X�P_�X�P���_�a��.Enabled = False
End If
End Sub

Private Sub ���ʶ��q_���q��l_Timer()
If �ثe��(31) = 0 Then
    Dim movecpn As Integer, mfd As Integer
    movecpn = movecp
    '===============
    movecom = atkingpagetot(2, 3)
    moveus = atkingpagetot(1, 3)
    Erase Vss_PersonMoveActionChangeNum
    Erase Vss_PersonMoveControlNum
    Vss_PersonAttackFirstControlNum = 0
    '===========================���涥�q���J�I(2)
    �԰��t����.���ʶ��q���ʫe���涥�q�I�s 2
    '===========================���涥�q���J�I(3)
    �԰��t����.���ʶ��q���ʫe���涥�q�I�s 3
    '===========================���涥�q���J�I(4)
    �԰��t����.���ʶ��q���ʫe���涥�q�I�s 4
    '===========================���涥�q���J�I(70)
    �԰��t����.���ʶ��q���ʫe���涥�q�I�s 70
    '============================
    If Vss_PersonMoveControlNum(1, 2) = 0 Then
        moveus = moveus + Vss_PersonMoveControlNum(1, 1)
    Else
        moveus = Vss_PersonMoveControlNum(1, 1)
    End If
    If Vss_PersonMoveControlNum(2, 2) = 0 Then
        movecom = movecom + Vss_PersonMoveControlNum(2, 1)
    Else
        movecom = Vss_PersonMoveControlNum(2, 1)
    End If
    '==================================
    If moveus < 0 Then moveus = 0
    If movecom < 0 Then movecom = 0
    '==================================
    movecheckus = moveus
    movecheckcom = movecom
    ��ܦC1.�q���貾�ʭ� = movecheckcom
    '----------�H�U���q���P�_�X�P�{���X�]���ʶ��q2�^
    If movecheckcom <= 0 Then
       �q���貾�ʶ��q��ܼ� = 2
    End If
    '==================================
    If Vss_PersonMoveActionChangeNum(1, 1) = 1 Then
        ��ܦC1.���ʶ��q��ܭ� = Vss_PersonMoveActionChangeNum(1, 2)
    End If
    If Vss_PersonMoveActionChangeNum(2, 1) = 1 Then
        �q���貾�ʶ��q��ܼ� = Vss_PersonMoveActionChangeNum(2, 2)
    End If
    '===============
    If Vss_EventPlayerAllActionOffNum(1) = 1 Then ��ܦC1.���ʶ��q��ܭ� = 0
    If Vss_EventPlayerAllActionOffNum(2) = 1 Then �q���貾�ʶ��q��ܼ� = 0
    '==================================
    ReDim VBEStageNum(0 To 4) As Integer
    VBEStageNum(0) = 71
    VBEStageNum(1) = moveus '�ϥΪ̤��`���ʼ�
    VBEStageNum(2) = movecom '�q�����`���ʼ�
    VBEStageNum(3) = ��ܦC1.���ʶ��q��ܭ� '�ϥΪ̤�ثe���ʶ��q��ʿ��
    VBEStageNum(4) = �q���貾�ʶ��q��ܼ� '�q����ثe���ʶ��q��ʿ��
    '===========================���涥�q���J�I(71)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 71, 1
    '============================
    If ��ܦC1.���ʶ��q��ܭ� = 1 Or ��ܦC1.���ʶ��q��ܭ� = 3 Then
       If ��ܦC1.���ʶ��q��ܭ� = 3 Then
          moveus = -Val(moveus)
          ��ܦC1.�ϥΪ̤貾�ʤ��~ = 1
       ElseIf ��ܦC1.���ʶ��q��ܭ� = 1 Then
          ��ܦC1.�ϥΪ̤貾�ʤ��~ = 2
       End If
     ��ܦC1.�ϥΪ̤貾�ʭ� = movecheckus
    End If
    '========
    If �q���貾�ʶ��q��ܼ� = 1 Or �q���貾�ʶ��q��ܼ� = 3 Then
       If �q���貾�ʶ��q��ܼ� = 3 Then
          movecom = -Val(movecom)
          ��ܦC1.�q���貾�ʤ��~ = 1
       ElseIf �q���貾�ʶ��q��ܼ� = 1 Then
          ��ܦC1.�q���貾�ʤ��~ = 2
       End If
       ��ܦC1.�q���貾�ʭ� = movecheckcom
    ElseIf �q���貾�ʶ��q��ܼ� = 2 Then
        If livecom(����H����ԤH��(2, 2)) < livecommax(����H����ԤH��(2, 2)) Then
            �^�_����_�q�� 1, 1, 0, True, True
        End If
        ��ܦC1.�q���貾�ʭ� = 0
    ElseIf �q���貾�ʶ��q��ܼ� = 4 Then
        ��ܦC1.�q���貾�ʭ� = 0
        �洫��������Ȯ��ܼ�(2) = 1
    ElseIf �q���貾�ʶ��q��ܼ� = 0 Then
        ��ܦC1.�q���貾�ʭ� = 0
    End If
    '==============================
    If ��ܦC1.���ʶ��q��ܭ� = 2 Then
         �^�_����_�ϥΪ� 1, 1, 0, True, True
         ��ܦC1.�ϥΪ̤貾�ʭ� = 0
    ElseIf ��ܦC1.���ʶ��q��ܭ� = 0 Then
      ��ܦC1.�ϥΪ̤貾�ʭ� = 0
    ElseIf ��ܦC1.���ʶ��q��ܭ� = 4 Then
      ��ܦC1.�ϥΪ̤貾�ʭ� = 0
      �洫��������Ȯ��ܼ�(1) = 1
    End If
    '==============================
    If (��ܦC1.���ʶ��q��ܭ� = 1 Or ��ܦC1.���ʶ��q��ܭ� = 3) Then
        movecpn = Val(moveus) + Val(movecpn)
    End If
    If (�q���貾�ʶ��q��ܼ� = 1 Or �q���貾�ʶ��q��ܼ� = 3) Then
        movecpn = Val(movecom) + Val(movecpn)
    End If
    '==============================
    
    If movecpn < 1 Then
       movecpn = 1
    ElseIf movecpn > 3 Then
       movecpn = 3
    End If
    
    ����ʧ@_�Z���ܧ� movecpn, True, True
    
    If Vss_PersonAttackFirstControlNum = 1 Then
        �԰��t����.movetnus
    ElseIf Vss_PersonAttackFirstControlNum = 2 Then
        �԰��t����.movetncom
    Else
        If Val(movecheckus) > Val(movecheckcom) Then
            �԰��t����.movetnus
        ElseIf Val(movecheckus) < Val(movecheckcom) Then
            �԰��t����.movetncom
        Else
            Randomize
            mfd = Int(Rnd() * 2) + 1
            If mfd = 1 Then �԰��t����.movetnus
            If mfd = 2 Then �԰��t����.movetncom
        End If
    End If
    �Y���淾�q�Ȯ��ܼ�(4) = moveturn
    HP�ˬd�ܼ� = False
    ���ݮɶ���C(2).Add 23
    FormMainMode.���ݮɶ�_2.Enabled = True
Else
    '===========================���涥�q���J�I(5)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l moveturn, 5, 1
    '============================
    '===========================���涥�q���J�I(6)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l moveturn, 6, 1
    '============================
    '===========================���涥�q���J�I(7)
    ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l moveturn, 7, 1
    '============================
    �ثe��(6) = 0
    �ثe��(10) = 1
    ���q���A�� = 2
    �q���X�P_�G�P.Enabled = True
End If
���ʶ��q_���q��l.Enabled = False
End Sub

Private Sub ���ʹϤ������ˬd_Timer()
If ��ܦC1.���ʤ�V�Ϥ���� = False Then
   ���P���q_�p��.Enabled = True
   ���ʹϤ������ˬd.Enabled = False
   FormMainMode.PEAFInterface.BnOKVisable False
End If
End Sub

Private Sub �P����_Timer()
Dim i As Integer

card(�P���ʼȮ��ܼ�(3)).Left = card(�P���ʼȮ��ܼ�(3)).Left + �Z�����(2, 1, 1)
card(�P���ʼȮ��ܼ�(3)).Top = card(�P���ʼȮ��ܼ�(3)).Top + �Z�����(2, 1, 2)
If Abs(�P���ʼȮ��ܼ�(1) - card(�P���ʼȮ��ܼ�(3)).Left) <= 50 Or Abs(�P���ʼȮ��ܼ�(2) - card(�P���ʼȮ��ܼ�(3)).Top) <= 50 Then
   card(�P���ʼȮ��ܼ�(3)).Left = �P���ʼȮ��ܼ�(1)
   card(�P���ʼȮ��ܼ�(3)).Top = �P���ʼȮ��ܼ�(2)
   card(�P���ʼȮ��ܼ�(3)).ZOrder
   For i = 1 To 3
       FormMainMode.PEAFpersoncardcom(i).ZOrder
   Next
   FormMainMode.PEAFAnimateInterface.ZOrder
   �P����.Enabled = False
   Select Case �ثe��(15)
        Case 1
            �o�P�ˬd.Enabled = True
        Case 2
            �ثe��(8) = 0
            �q���X�P_��P���.Enabled = True
        Case 3
            'Nothing
        Case 4
            card(�ثe��(20)).Visible = False
            �ثe��(4) = 0
            �ثe��(13) = 0
            �ϥΪ̥X�P_��P���.Enabled = True
        Case 5
            card(�ثe��(16)).Visible = False
            �ثe��(8) = 0
            �q���X�P_��P���.Enabled = True
        Case 6
            '===========�ƥ�d����_���|_�ϥΪ�(���q2)
            card(�ƥ�d�O���Ȯɼ�(1, 4)).Visible = False
            ���ݮɶ���C(2).Add 6
            ���ݮɶ�_2.Enabled = True
        Case 7
             '===========�ƥ�d����_���|_�ϥΪ�(���q1)
            ���ݮɶ���C(2).Add 5
            ���ݮɶ�_2.Enabled = True
        Case 8
            '===========�ƥ�d����_���|_�ϥΪ�(���q3)
            �ƥ�d�O���Ȯɼ�(1, 3) = 3
            �ƥ�d.���|_�ϥΪ� 0, 0
        Case 9
             '===========�ƥ�d����_���|_�q��(���q1)
            ���ݮɶ���C(2).Add 7
            ���ݮɶ�_2.Enabled = True
        Case 10
            '===========�ƥ�d����_���|_�q��(���q3)
            card(�ƥ�d�O���Ȯɼ�(2, 4)).Visible = False
            ���ݮɶ���C(2).Add 8
            ���ݮɶ�_2.Enabled = True
        Case 11
            '===========�ƥ�d����_���|_�q��(���q4)
            �ƥ�d�O���Ȯɼ�(2, 3) = 4
            �ƥ�d.���|_�q�� 0, 0
        Case 12
            '===========�ƥ�d����_�A�G�N_�ϥΪ�(���q1)
            ���ݮɶ���C(2).Add 11
            ���ݮɶ�_2.Enabled = True
        Case 13
            '===========�ƥ�d����_�A�G�N_�ϥΪ�(���q6)
            card(�ƥ�d�O���Ȯɼ�(1, 4)).Visible = False
            �ƥ�d�O���Ȯɼ�(1, 3) = 6
            �ƥ�d.�A�G�N_�ϥΪ� 0, 0
        Case 14
            '===========�ƥ�d����_�A�G�N_�q��(���q1)
            ���ݮɶ���C(2).Add 13
            ���ݮɶ�_2.Enabled = True
        Case 15
            '===========�ƥ�d����_�A�G�N_�q��(���q5>6)
            card(�ƥ�d�O���Ȯɼ�(2, 4)).Visible = False
            �ƥ�d�O���Ȯɼ�(2, 3) = 6
            �ƥ�d.�A�G�N_�q�� 0, 0
        Case 16
            '===========�ƥ�d����_HP�^�__�ϥΪ�(���q1)
            ���ݮɶ���C(2).Add 16
            ���ݮɶ�_2.Enabled = True
            turnpageonin = 0
            FormMainMode.PEAFInterface.BnOKEnabled False
        Case 17
            '===========�ƥ�d����_HP�^�__�ϥΪ�(���q4)
            card(�ƥ�d�O���Ȯɼ�(1, 4)).Visible = False
            �ƥ�d�O���Ȯɼ�(1, 3) = 4
            �ƥ�d.HP�^�__�ϥΪ� 0, 0
        Case 18
            '===========�ƥ�d����_HP�^�__�q��(���q1)
            ���ݮɶ���C(2).Add 18
            ���ݮɶ�_2.Enabled = True
        Case 19
            '===========�ƥ�d����_HP�^�__�q��(���q4>5)
            card(�ƥ�d�O���Ȯɼ�(2, 4)).Visible = False
            �ƥ�d�O���Ȯɼ�(2, 3) = 5
            �ƥ�d.HP�^�__�q�� 0, 0
        Case 20
            �ثe��(4) = 0
            �ثe��(13) = 0
            �ϥΪ̥X�P_��P���.Enabled = True
        Case 21
            If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDrawCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDrawCards")) = 2 '(���q2)
            End If
        Case 22
           If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGetUsedCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGetUsedCards")) = 2 '(���q2)
            End If
        Case 23
            If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards")) = 3 '(���q3)
            End If
        Case 40
            ���ݮɶ���C(2).Add 37
            ���ݮɶ�_2.Enabled = True
        Case 41
            '===========�ƥ�d����_�t��_�ϥΪ�(���q1)
            ���ݮɶ���C(2).Add 39
            ���ݮɶ�_2.Enabled = True
            turnpageonin = 0
            FormMainMode.PEAFInterface.BnOKEnabled False
        Case 42
            '===========�ƥ�d����_�t��_�ϥΪ�(���q4>5)
            card(�ƥ�d�O���Ȯɼ�(1, 4)).Visible = False
            �ƥ�d�O���Ȯɼ�(1, 3) = 4
            �ƥ�d.�t��_�ϥΪ� 0, 0
        Case 43
            '===========�ƥ�d����_�t��_�q��(���q1)
            ���ݮɶ���C(2).Add 41
            ���ݮɶ�_2.Enabled = True
        Case 44
            '===========�ƥ�d����_�t��_�q��(���q4>5)
            card(�ƥ�d�O���Ȯɼ�(2, 4)).Visible = False
            �ƥ�d�O���Ȯɼ�(2, 3) = 5
            �ƥ�d.�t��_�q�� 0, 0
   End Select
End If
End Sub


Private Sub �P����_���P_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

If �ثe��(11) = pageqlead(�ثe��(10)) Then
    �԰��t����.checkpage
    �P����_���P.Enabled = False
    �ثe��(10) = �ثe��(10) + 1
    ���P���q_�p��.Enabled = True
    Exit Sub
End If
For i = 1 + �ثe��(11) To pageqlead(�ثe��(10)) - �ثe��(12)
    If Abs(240 - card(�Z�����_���P�Ȯɼ�(i, 3)).Left) <= 10 Or Abs(960 - card(�Z�����_���P�Ȯɼ�(i, 3)).Top) <= 10 Then
        card(�Z�����_���P�Ȯɼ�(i, 3)).Left = 240
        card(�Z�����_���P�Ȯɼ�(i, 3)).Top = 960
        card(�Z�����_���P�Ȯɼ�(i, 3)).Visible = False
        
        Set tmpcard = �԰��t����.CardDeckCollection(�԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(�Z�����_���P�Ȯɼ�(i, 3))))(CStr(�Z�����_���P�Ȯɼ�(i, 3)))
        tmpcard.Location = 3
        Select Case tmpcard.CardType
            Case 1 '���εP
                �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(�Z�����_���P�Ȯɼ�(i, 3))), 2
            Case 2 '�ƥ�d
                �԰��t����.�d�P�P�ﶰ�X�� tmpcard, �԰��t����.�d�P�P�ﶰ�X����_CollectionIndex(CStr(�Z�����_���P�Ȯɼ�(i, 3))), 9
        End Select
        
        �ثe��(11) = �ثe��(11) + 1
    End If
    card(�Z�����_���P�Ȯɼ�(i, 3)).Left = card(�Z�����_���P�Ȯɼ�(i, 3)).Left + �Z�����_���P�Ȯɼ�(i, 1)
    card(�Z�����_���P�Ȯɼ�(i, 3)).Top = card(�Z�����_���P�Ȯɼ�(i, 3)).Top + �Z�����_���P�Ȯɼ�(i, 2)
    If �ثe��(12) > 0 Then
        �ثe��(12) = �ثe��(12) - 1
    End If
Next

End Sub

Private Sub �o�P_�ϥΪ̶��q_Timer()
�o�P_�ϥΪ̶��q.Enabled = False
�ثe��(2) = 2

If Val(pageusglead) < �P�`���q��(1) Then
    �԰��t����.����ʧ@_��P_���εP 1
Else
    �o�P�ˬd.Enabled = True
End If
End Sub

Private Sub �o�P_�q�����q_Timer()
�o�P_�q�����q.Enabled = False
�ثe��(2) = 3

If Val(pagecomglead) < �P�`���q��(2) Then
    �԰��t����.����ʧ@_��P_���εP 2
Else
    �o�P�ˬd.Enabled = True
End If
End Sub

Private Sub �o�P�ˬd_Timer()
If (Val(pageusglead) >= �P�`���q��(1) And Val(pagecomglead) >= �P�`���q��(2)) Or BattleCardNum <= 0 Then
   �o�P�ˬd.Enabled = False
   �ثe��(15) = 0
   ���ݮɶ���C(1).Add 3
   ���ݮɶ�.Enabled = True
Else
   Select Case �ثe��(2)
       Case 1
           �o�P_�ϥΪ̶��q.Enabled = True
           �o�P�ˬd.Enabled = False
       Case 2
           �o�P_�q�����q.Enabled = True
           �o�P�ˬd.Enabled = False
        Case 3
           �ثe��(2) = 1
    End Select
End If

End Sub

Private Sub ���ݮɶ�_2_Timer()
Select Case �ثe��(14)
   Case 0
      �ثe��(14) = �ثe��(14) + 1
   Case 1
      If ���ݮɶ���C(2).Count <= 1 Then
          �ثe��(14) = 0
          ���ݮɶ�_2.Enabled = False
      End If
      If ���ݮɶ���C(2).Count = 0 Then Exit Sub
      Select Case ���ݮɶ���C(2).item(1)
          Case 1
              '========�}�l��l���q1
                ��ܦC1.Visible = True
                ��ܦC1.���ʶ��q����� = False
                ��ܦC1.���ʤ�V�Ϥ���� = False
                �@��t����.���ļ��� 6
                If �t����ܬɭ������� = 1 Then
                    draw1.Visible = False
                    draw2.Visible = True
                Else
                    FormMainMode.PEAFInterface.stagejpg app_path & "gif\system\draw2.gif"
                End If
                ���ݮɶ���C(1).Add 2
                ���ݮɶ�.Enabled = True
          Case 2
              cn22_Click
              FormMainMode.PEAFInterface.BnOKVisable False
           Case 3
              cn32_Click
              FormMainMode.PEAFInterface.BnOKVisable False
           Case 4
              Select Case turnatk
                    Case 1
                        ���ݮɶ���C(1).Add 7
                        ���ݮɶ�.Enabled = True
                    Case 2
                        ���ݮɶ���C(1).Add 8
                        ���ݮɶ�.Enabled = True
                    Case 3
                        cnmove2_Click
                End Select
           Case 5
                '===========�ƥ�d����_���|_�ϥΪ�(���q1>2)
                �ƥ�d�O���Ȯɼ�(1, 3) = 2
                �ƥ�d.���|_�ϥΪ� 0, 0
           Case 6
                '===========�ƥ�d����_���|_�ϥΪ�(���q2>3)
                �ƥ�d�O���Ȯɼ�(1, 3) = 3
                �ƥ�d.���|_�ϥΪ� 0, 0
           Case 7
                '===========�ƥ�d����_���|_�q��(���q1>2)
                �ƥ�d�O���Ȯɼ�(2, 3) = 2
                �ƥ�d.���|_�q�� 0, 0
           Case 8
                '===========�ƥ�d����_���|_�q��(���q3>4)
                �ƥ�d�O���Ȯɼ�(2, 3) = 4
                �ƥ�d.���|_�q�� 0, 0
            Case 9
                '===========�ƥ�d����_���|_�q��(���q2>3)
                �ƥ�d�O���Ȯɼ�(2, 3) = 3
                �ƥ�d.���|_�q�� 0, 0
            Case 10
                �q���X�P.Enabled = True
            Case 11
                '===========�ƥ�d����_�A�G�N_�ϥΪ�(���q1>2)
                �ƥ�d�O���Ȯɼ�(1, 3) = 2
                �ƥ�d.�A�G�N_�ϥΪ� 0, 0
            Case 12
                '===========�ƥ�d����_�A�G�N_�ϥΪ�(���q>5)
                �ƥ�d�O���Ȯɼ�(1, 3) = 5
                �ƥ�d.�A�G�N_�ϥΪ� 0, 0
            Case 13
                '===========�ƥ�d����_�A�G�N_�q��(���q1>2)
                �ƥ�d�O���Ȯɼ�(2, 3) = 2
                �ƥ�d.�A�G�N_�q�� 0, 0
            Case 14
                '===========�ƥ�d����_�A�G�N_�q��(���q>4)
                �ƥ�d�O���Ȯɼ�(2, 3) = 4
                �ƥ�d.�A�G�N_�q�� 0, 0
            Case 15
                '===========�ƥ�d����_�A�G�N_�q��(���q4>5)
                �ƥ�d�O���Ȯɼ�(2, 3) = 5
                �ƥ�d.�A�G�N_�q�� 0, 0
            Case 16
                '===========�ƥ�d����_HP�^�__�ϥΪ�(���q1>2)
                �ƥ�d�O���Ȯɼ�(1, 3) = 2
                �ƥ�d.HP�^�__�ϥΪ� 0, 0
            Case 17
                '===========�ƥ�d����_HP�^�__�ϥΪ�(���q2>3)
                �ƥ�d�O���Ȯɼ�(1, 3) = 3
                �ƥ�d.HP�^�__�ϥΪ� 0, 0
            Case 18
                '===========�ƥ�d����_HP�^�__�q��(���q1>2)
                �ƥ�d�O���Ȯɼ�(2, 3) = 2
                �ƥ�d.HP�^�__�q�� 0, 0
            Case 19
                '===========�ƥ�d����_HP�^�__�q��(���q2>3)
                �ƥ�d�O���Ȯɼ�(2, 3) = 3
                �ƥ�d.HP�^�__�q�� 0, 0
            Case 20
                '===========�ƥ�d����_HP�^�__�q��(���q3>4)
                �ƥ�d�O���Ȯɼ�(2, 3) = 4
                �ƥ�d.HP�^�__�q�� 0, 0
            Case 21
                Select Case turnatk
                   Case 1
                       �԰��t����.����ʧ@_�������q�����ɧޯ�Ұ�
                   Case 2
                       �԰��t����.����ʧ@_���m���q�����ɧޯ�Ұ�
               End Select
            Case 22
               FormMainMode.��l���槹�Ұ�.Enabled = True
            Case 23
                �ثe��(31) = 1
                FormMainMode.���ʶ��q_���q��l.Enabled = True
            Case 24
                If FormMainMode.PEAFDiceInterface.DiceStop = True Or ��ƹs�ˬd��(1) = True Or ��ƹs�ˬd��(2) = True Then
                    If ���涥�q�t��_�j�M���b���椧���涥�q("BattleStartDice") <> 0 Then
                        vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("BattleStartDice")) = 2 '(���q2)
                    End If
                Else
                    ���ݮɶ���C(2).Add 24
                    ���ݮɶ�_2.Enabled = True
                End If
            Case 25
                If FormMainMode.PEAFDiceInterface.DiceStop = True Or ��ƹs�ˬd��(1) = True Or ��ƹs�ˬd��(2) = True Then
                    �԰��t����.�Y�����P�_
                Else
                    ���ݮɶ���C(2).Add 25
                    ���ݮɶ�_2.Enabled = True
                End If
            Case 30
                If �q���X�P_�G�P.Enabled = False Then
                    ��ܦC1.���ʤ�V�Ϥ���� = True
                    ���ʹϤ������ˬd.Enabled = True
                Else
                    ���ݮɶ���C(2).Add 30
                    ���ݮɶ�_2.Enabled = True
                End If
            Case 39
                '===========�ƥ�d����_�t��_�ϥΪ�(���q1>2)
                �ƥ�d�O���Ȯɼ�(1, 3) = 2
                �ƥ�d.�t��_�ϥΪ� 0, 0
            Case 40
                '===========�ƥ�d����_�t��_�ϥΪ�(���q2>3)
                �ƥ�d�O���Ȯɼ�(1, 3) = 3
                �ƥ�d.�t��_�ϥΪ� 0, 0
            Case 41
                '===========�ƥ�d����_�t��_�q��(���q1>2)
                �ƥ�d�O���Ȯɼ�(2, 3) = 2
                �ƥ�d.�t��_�q�� 0, 0
            Case 42
                '===========�ƥ�d����_�t��_�q��(���q2>3)
                �ƥ�d�O���Ȯɼ�(2, 3) = 3
                �ƥ�d.�t��_�q�� 0, 0
            Case 43
                '===========�ƥ�d����_�t��_�q��(���q3>4)
                �ƥ�d�O���Ȯɼ�(2, 3) = 4
                �ƥ�d.�t��_�q�� 0, 0
            Case 45
                �ثe��(32) = 1
                FormMainMode.�ϥΪ̥X�P_AI�X�P����_�ƥ�d.Enabled = True
            Case 46
                '====================���紼�z��AI�X�P�t��
                Dim wtyr As Integer '�Ȯ��ܼ�
                If (moveturn = 1 And turnatk = 2) Or (moveturn = 2 And turnatk = 1) Then wtyr = 1 Else wtyr = 0
                ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_��� 1, turnatk, nameus(����H����ԤH��(1, 2)), movecp, wtyr
                ���z��AI�t����.���z��AI�t��_�ϥΪ̥X�P���q�P�_����
                �ثe��(32) = 1
                FormMainMode.�ϥΪ̥X�P_AI�X�P����.Enabled = True
            Case 48
                ����ʧ@_�q����U���q�X�P�������� turnatk
      End Select
      ���ݮɶ���C(2).Remove 1
End Select
End Sub

Private Sub ���ݮɶ�_Timer()
Select Case �ثe��(22)
    Case 0
       �ثe��(22) = �ثe��(22) + 1
    Case 1
        If ���ݮɶ���C(1).Count <= 1 Then
            �ثe��(22) = 0
            ���ݮɶ�.Enabled = False
        End If
        If ���ݮɶ���C(1).Count = 0 Then Exit Sub
        Select Case ���ݮɶ���C(1).item(1)
            Case 2   '========�}�l��l���q2
                ���ݮɶ���C(1).Add 5
                ���ݮɶ�.Enabled = True
            Case 3
                ���ݮɶ���C(1).Add 22
                ���ݮɶ�.Enabled = True
            Case 4
                �԰��t����.�s���T�� "�{�b���Z��" & movecp & "�C"
                �洫��������Ȯ��ܼ�(4) = 1
                �԰��t����.����ʧ@_���ʶ��q��ܰ���
            Case 5
                cn1_Click
            Case 6
                Erase Vss_EventPlayerAllActionOffNum
                '===========================���涥�q���J�I(1)
                ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 1, 1
                '============================
                cnmove_Click
            Case 7
                ���ݮɶ���C(2).Add 2
                ���ݮɶ�_2.Enabled = True
            Case 8
                ���ݮɶ���C(2).Add 3
                ���ݮɶ�_2.Enabled = True
            Case 9
                cn2_Click
                ��ܦC1.Visible = True
                �԰��t����.�ɶ��b_���
            Case 10
                �q����ƥ�d�O�_�X����ܼ� = False
                cn3_Click
                ��ܦC1.Visible = True
                �԰��t����.�ɶ��b_���
            Case 11
                �԰��t����.�ɶ��b_����
                ��ܦC1.Visible = False
                ���ݮɶ���C(1).Add 12
                ���ݮɶ�.Enabled = True
            Case 12
                If Val(�Y���淾�q�Ȯ��ܼ�(4)) = 1 Then
                   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
                    Case 1
                        '===========================���涥�q���J�I(ATK-13/DEF-33)
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 13, 2
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 33, 2
                        '============================
                    Case 2
                       '===========================���涥�q���J�I(ATK-13/DEF-33)
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 13, 2
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 33, 2
                        '============================
                    End Select
                Else
                   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
                    Case 1
                       '===========================���涥�q���J�I(ATK-13/DEF-33)
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 13, 2
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 33, 2
                        '============================
                    Case 2
                       '===========================���涥�q���J�I(ATK-13/DEF-33)
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 1, 13, 2
                        ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l 2, 33, 2
                        '============================
                    End Select
                End If
                �Y���淾�q�Ȯ��ܼ�(9) = �������m��l�`��(1)
                �Y���淾�q�Ȯ��ܼ�(10) = �������m��l�`��(2)
                �O�_�t�Τ��� = True
                �԰��t����.�Y�������
                ���ݮɶ���C(2).Add 25
                FormMainMode.���ݮɶ�_2.Enabled = True
            Case 13
                ���ݮɶ���C(1).Add 9
                ���ݮɶ�.Enabled = True
            Case 14
                ���ݮɶ���C(1).Add 10
                ���ݮɶ�.Enabled = True
            Case 15
                cn4_Click
            Case 17
                '===========================���涥�q���J�I(9)
                 ���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l moveturn, 9, 1
                '============================
                Select Case moveturn
                  Case 1
                     cn2_Click
                  Case 2
                     �q����ƥ�d�O�_�X����ܼ� = False
                     cn3_Click
                End Select
            Case 18
                 �԰��t����.����ʧ@_�洫�H������_�q��_��l
            Case 19
                 �԰��t����.����ʧ@_�洫�H������_�q��_�洫
            Case 20
                 �԰��t����.�ɶ��b_����
                 ��ܦC1.Visible = False
                 cn4_Click
            Case 21
                �洫��������Ȯ��ܼ�(4) = 2
                ����ʧ@_�H�����`�洫���q��ܰ���
            Case 22
                �԰��t����.�ƥ�d�B�z_����_�ϥΪ̤�
                �԰��t����.�ƥ�d�B�z_����_�q����
                ���ݮɶ���C(1).Add 6
                ���ݮɶ�.Enabled = True
            Case 30
                �q���X�P.Enabled = True
            Case 36
                FormMainMode.trend.Enabled = True
            Case 37
                Dim ckl As Integer
                '=============�ϥΪ̤��ܦ��
                If turnatk = 3 Then
                    ��ܦC1.���ʶ��q��ܭ� = �ثe��(33)
                End If
                '====================
                FormMainMode.PEAFInterface.BnOKStartListen
                FormMainMode.PEAFInterface_BnOKClick
                For ckl = 1 To �԰��t����.ActionCardTotNum
                    FormMainMode.card(ckl).CardEnabledType = True
                Next
        End Select
        ���ݮɶ���C(1).Remove 1
End Select
End Sub

Private Sub �q���X�P_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

�q���X�P.Enabled = False
If �q����ƥ�d�O�_�X����ܼ� = False Then
     '=========================�M�ݨƥ�d�X�P���q
    For i = 1 To �԰��t����.CardDeckCollection(7).Count
        Set tmpcard = �԰��t����.CardDeckCollection(7)(i)
        If tmpcard.CardType = 2 Then
            If tmpcard.UpperType = a6a Then
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a6a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            End If
            If tmpcard.UpperType = a7a And (turnatk = 1 Or turnatk = 2) Then
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a7a And (turnatk = 1 Or turnatk = 2) Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            End If
            If tmpcard.UpperType = a8a Then
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a8a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            End If
            If tmpcard.UpperType = a9a Then
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            ElseIf tmpcard.LowerType = a9a Then
                Call tmpcard.Reverse
                tmpcard.ComMark = 1
                �԰��t����.�q���P_�������P tmpcard.CardNum
                Exit Sub
            End If
        End If
    Next
    '==============================�ƥ�d���w�X�P����
    �q����ƥ�d�O�_�X����ܼ� = True
    Select Case turnatk
        Case 1
             �������q_���q1.Enabled = True
        Case 2
             cn3_Click
        Case 3
             cnmove_Click
    End Select
    Exit Sub
End If
'===========================================
If �q����ƥ�d�O�_�X����ܼ� = True Then
    Do
        �ثe��(6) = �ثe��(6) + 1
        If �ثe��(6) > �԰��t����.CardDeckCollection(7).Count Then
            �q����ƥ�d�O�_�X����ܼ� = False
            Select Case turnatk
               Case 1
                  �ثe��(6) = 0
                  �ثe��(10) = 1
                  �԰��t����.�ɶ��b_����
                  �q���X�P_�G�P.Enabled = True
                  trgoi2_Timer
               Case 2
                  �ثe��(6) = 0
                  �ثe��(10) = 1
                  �԰��t����.�ɶ��b_����
                  �q���X�P_�G�P.Enabled = True
                  trgoi2_Timer
                  trgoi1_Timer
               Case 3
                    ����ʧ@_�q����U���q�X�P�������� 3
            End Select
            Exit Sub
        End If
        Set tmpcard = �԰��t����.CardDeckCollection(7)(�ثe��(6))
        If tmpcard.ComMark = 1 Then
            �ثe��(6) = �ثe��(6) - 1
            �԰��t����.�q���P_�������P tmpcard.CardNum
            Exit Sub
        End If
    Loop
End If
End Sub

Private Sub �q���X�P_��P���_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

If �ثe��(8) < 240 Then
    For i = 1 To �԰��t����.CardDeckCollection(7).Count
        Set tmpcard = �԰��t����.CardDeckCollection(7)(i)
        If i >= �ثe��(9) Then
            card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left + (240 / 10)
        End If
    Next
    �ثe��(8) = �ثe��(8) + (240 / 10)
Else
    �q���X�P_��P���.Enabled = False
    Select Case �ثe��(17)
        Case 1
            If �P����.Enabled = False Then
                �q���X�P.Enabled = True
            Else
                �q���X�P_��P���.Enabled = True  '���ݵP���ʧ���
            End If
        Case 2
            '======�����ʧ@
        Case 3
            If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingSeizeEnemyCards")) = 3 '(���q3)
            End If
        Case 4
            If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDestroyCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingDestroyCards")) = 3 '(���q3)
            End If
        Case 5
            If ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGiveCards") <> 0 Then
                vbecommadnum(2, ���涥�q�t��_�j�M���b���椧���涥�q("AtkingGiveCards")) = 3 '(���q3)
            End If
        Case 6
           '===========�ƥ�d����_�A�G�N_�ϥΪ�(���q3)
            �ƥ�d�O���Ȯɼ�(1, 3) = 3
            �ƥ�d.�A�G�N_�ϥΪ� 0, 0
    End Select
    
End If
End Sub


Private Sub �q���X�P_�X�P���_�a�k_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To �԰��t����.CardDeckCollection(8).Count
    Set tmpcard = �԰��t����.CardDeckCollection(8)(i)
    If i < �ثe��(9) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left + (480 / 10)
    End If
    If i >= �ثe��(9) Then
       card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (500 / 10)
    End If
Next
�ثe��(7) = �ثe��(7) + (480 / 10)
If �ثe��(7) >= 480 Then
    �q���X�P_�X�P���_�a�k.Enabled = False
End If
End Sub

Private Sub �q���X�P_�X�P���_�a��_Timer()
Dim i As Integer
Dim tmpcard As clsActionCard

For i = 1 To (�԰��t����.CardDeckCollection(8).Count - 1)
    Set tmpcard = �԰��t����.CardDeckCollection(8)(i)
    card(tmpcard.CardNum).Left = card(tmpcard.CardNum).Left - (480 / 10)
Next
�ثe��(7) = �ثe��(7) + (480 / 10)
If �ثe��(7) >= 480 Then
    �q���X�P_�X�P���_�a��.Enabled = False
End If
End Sub


Private Sub �q���X�P_�G�P_Timer()
Dim tmpcard As clsActionCard

�ثe��(6) = �ثe��(6) + 1
If �ثe��(6) > �԰��t����.CardDeckCollection(8).Count Then
    �q���X�P_�G�P.Enabled = False
    Select Case turnatk
       Case 1, 2
            ����ʧ@_�q����U���q�X�P�������� turnatk
       Case 3
            ���ݮɶ���C(2).Add 30
            ���ݮɶ�_2.Enabled = True
    End Select
    Exit Sub
End If

Set tmpcard = �԰��t����.CardDeckCollection(8)(�ثe��(6))
�԰��t����.���εP�^�_���� tmpcard.CardNum
�@��t����.���ļ��� 4
End Sub

Private Sub ��������ˬd_Timer()
If �ϥΪ̥X�P_�X�P���_�a��.Enabled = False And �ϥΪ̥X�P_�X�P���_�a�k.Enabled = False And �ϥΪ̥X�P_��P���.Enabled = False And �P����.Enabled = False Then
   turnpageonin = 1
   ��������ˬd.Enabled = False
End If
End Sub

Private Sub ��l���槹�Ұ�_Timer()
Dim uscomvsn As Integer
��l���槹�Ұ�.Enabled = False
'===========================
If Val(�Y���淾�q�Ȯ��ܼ�(4)) = 1 Then
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
    Case 1
        uscomvsn = 1
    Case 2
        uscomvsn = 2
    End Select
Else
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
    Case 1
       uscomvsn = 2
    Case 2
       uscomvsn = 1
    End Select
End If
'===========================���涥�q���J�I(20)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 20, 1
'============================
'===========================���涥�q���J�I(21)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 21, 1
'============================
'===========================���涥�q���J�I(22)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 22, 1
'============================
'===========================���涥�q���J�I(23)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 23, 1
'============================
'===========================���涥�q���J�I(24)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 24, 1
'============================
'===========================���涥�q���J�I(25)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 25, 1
'============================
'===========================���涥�q���J�I(26)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 26, 1
'============================
'===========================���涥�q���J�I(27)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 27, 1
'============================
'===========================���涥�q���J�I(28)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 28, 1
'============================
'===========================���涥�q���J�I(29)
���涥�q�t����.���涥�q�t���`�D�n�{��_���涥�q�}�l uscomvsn, 29, 1
'============================
trnextend.Enabled = True
End Sub

Private Sub �v�l�]�w_Click()
FormDevSetting.smallleftus.Caption = personusminijpg.�p�H���v�lLeft
FormDevSetting.smalltopus.Caption = personusminijpg.�p�H���v�ltop�t
FormDevSetting.smallleftcom.Caption = personcomminijpg.�p�H���v�lLeft
FormDevSetting.smalltopcom.Caption = personcomminijpg.�p�H���v�ltop�t
FormDevSetting.smallpnleftus.Caption = personusminijpg.Left
FormDevSetting.smallpntopus.Caption = personusminijpg.Top
FormDevSetting.smallpnleftcom.Caption = personcomminijpg.Left
FormDevSetting.smallpntopcom.Caption = personcomminijpg.Top
FormDevSetting.personfus.Caption = ��ܦC1.�ϥΪ̤�p�H���Ϥ�left
FormDevSetting.personfcom.Caption = ��ܦC1.�q����p�H���Ϥ�left
If Formsetting.checktest.Value = 1 Then
    FormDevSetting.Height = 6825
ElseIf Formsetting.checktestpersondown.Value = 1 Then
    FormDevSetting.Height = 3075
End If
'=================
�԰��t����.�ɶ��b_����
'=================
FormDevSetting.Show 1
End Sub
Private Sub bnabout_Click()
FormAbout.Show 1
�@��t����.���ļ��� 11
End Sub

Private Sub bnconfig_Click()
Formsetting.Left = FormMainMode.Left + 915
Formsetting.Top = FormMainMode.Top + 300
�@��t����.�ۥѾ԰��Ҧ��]�w���U���]�wŪ�J�{��
�@��t����.���ļ��� 11
Formsetting.Show 1
End Sub

Sub PEGFbnstart_Click()
PEGameFreeModeSettingForm.Enabled = False
�@��t����.�}�l�C���i��{��
End Sub

Private Sub Form_Load()
'============
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'==============
Dim cmstr() As String
Dim i As Integer

cmstr = Split(Command$, "/")
If UBound(cmstr) > 0 Then
    For i = 0 To UBound(cmstr)
        If cmstr(i) = "wine" Then �@��t����.ProgramIsOnWine = True
    Next
End If
�@��t����.�P�_�r��_FormMainMode
�@��t����.�D���_PEStartForm���
End Sub
Private Sub personreadifus_Click()
cdgpersonus.ShowOpen
�H���t����.�d���H����TŪ�J_�춥�q cdgpersonus.filename
End Sub
Private Sub personlevelcom_Click(Index As Integer)
�H���t����.�M������H����T�ܼ� 2, Index
'�H���t����.�d���H����TŪ�J_�T���q_�q�� personnamecom(Index).Text, personlevelcom(Index).Text, Index, 2
�H���t����.�d���H����TŪ�J_�T���q personnamecom(Index).Text, personlevelcom(Index).Text, Index, 2
'�H���t����.�d���H����TŪ�J_�|���q_�q�� personnamecom(Index).Text, Index   '���Unlight�x��L�q�����ܭ�h
'�H���t����.�d���H����TŪ�J_�|���q personnamecom(Index).Text, Index, 2 '���Unlight�x��L�q�����ܭ�h
�H���t����.�d���H����T���_�q�� Index
End Sub

Private Sub personlevelus_Click(Index As Integer)
�H���t����.�M������H����T�ܼ� 1, Index
'�H���t����.�d���H����TŪ�J_�T���q_�ϥΪ� personnameus(Index).Text, personlevelus(Index).Text, Index, 1
�H���t����.�d���H����TŪ�J_�T���q personnameus(Index).Text, personlevelus(Index).Text, Index, 1
'�H���t����.�d���H����TŪ�J_�|���q_�ϥΪ� personnameus(Index).Text, Index
�H���t����.�d���H����TŪ�J_�|���q personnameus(Index).Text, Index, 1
�H���t����.�d���H����T���_�ϥΪ� Index
End Sub

Private Sub personnamecom_Click(Index As Integer)
If ���q���ƥ� = True Then
    ��s�H���M��_�q����_�ܧ� Index
    If personnamecom(Index).Text = "" Or personnamecom(Index).Text = "�m�H���n" Then
       personlevelcom(Index).Clear
        �H���t����.�����H��_�q�� Index
        �H���t����.�d���H����T���_�q�� Index
    Else
       �d���H����TŪ�J_�G���q_�q�� personnamecom(Index).Text, Index
    End If
    personlevelcom(Index).ListIndex = personlevelcom(Index).ListCount - 1
End If
End Sub

Private Sub personnameus_Click(Index As Integer)
If ���ϥΪ̨ƥ� = True Then
    ��s�H���M��_�ϥΪ̤�_�ܧ� Index
    If personnameus(Index).Text = "" Or personnameus(Index).Text = "�m�H���n" Then
        personlevelus(Index).Clear
        �H���t����.�����H��_�ϥΪ� Index
        �H���t����.�d���H����T���_�ϥΪ� Index
    Else
        �d���H����TŪ�J_�G���q_�ϥΪ� personnameus(Index).Text, Index
    End If
    personlevelus(Index).ListIndex = personlevelus(Index).ListCount - 1
End If
End Sub

Private Sub personresetcom_Click(Index As Integer)
personnamecom(Index).ListIndex = -1
personlevelcom(Index).Clear
End Sub

Private Sub personresetus_Click(Index As Integer)
personnameus(Index).ListIndex = -1
personlevelus(Index).Clear
End Sub
Private Sub start1_Timer()
Dim i As Integer
If st > 200 Then
   stup.Enabled = True
   stdown.Enabled = True
   start1.Enabled = False
   start2.Enabled = True
   For i = 1 To 3
      If PEASusbi1(i).Caption = "0" Then
         PEAScardus(i).Visible = False
         cardusname(i).Visible = False
         cardusspname(i).Visible = False
         Formchangeperson.card(i - 1).Visible = False
         Formchangeperson.bnok(i - 1).Visible = False
      Else
         PEAScardus(i).Visible = True
      End If
      If PEAScardcompi1(i).Caption = "0" Then
         PEAScardcom(i).Visible = False
         cardcomname(i).Visible = False
         cardcomspname(i).Visible = False
      Else
         PEAScardcom(i).Visible = True
      End If
   Next
   If Formsetting.chkpersonvsmode.Value = 1 Then
       For i = 2 To 3
           PEAScardcompi1(i).Caption = "?"
           PEAScardcompi2(i).Caption = "?"
           PEAScardcompi3(i).Caption = "?"
           PEAScardcom(i).Picture = LoadPicture(app_path & "gif\system\personunknown.jpg")
           cardcomname(i).Caption = "UnKnown"
           cardcomspname(i).Visible = False
        Next
    End If
    '==============
   downjpg.Visible = True
   upjpg_2.Visible = True
   �}�l�d�����ʰʵe������(1, 4) = ����H����ԤH��(1, 1)
   �}�l�d�����ʰʵe������(2, 4) = ����H����ԤH��(2, 1)
Else
  st = Val(st) + 1
End If
End Sub

Private Sub start2_Timer()
If sq = 401 Then
   tr�j�H���ι�_�ϥΪ�.Enabled = True
   tr�j�H���ι�_�q��.Enabled = True
   sq = Val(sq) + 1
ElseIf sq = 500 Then
   �@��t����.�D���_PEAttackingForm���
   PEAttackingStartForm.Visible = False
   start2.Enabled = False
   FormMainMode.��q���J�ʵe.Enabled = True
Else
   sq = Val(sq) + 1
End If
   
End Sub

Private Sub stdown_Timer()
If sq <= 400 Then
   If downjpg.Top <= 8400 Then
      downjpg.Top = 8400
      stdown.Enabled = False
      cardustr.Enabled = True
      cardcomtr.Enabled = True
   Else
      downjpg.Top = Val(downjpg.Top) - 60
   End If
Else
   If downjpg.Top >= Val(FormMainMode.Height) Then
      downjpg.Top = Val(FormMainMode.Height)
      stdown.Enabled = False
   Else
      downjpg.Top = Val(downjpg.Top) + 60
   End If
End If
End Sub

Private Sub stup_Timer()
If sq <= 400 Then
   If upjpg_2.Top >= 0 Then
      upjpg_2.Top = 0
      stup.Enabled = False
   Else
      upjpg_2.Top = Val(upjpg_2.Top) + 60
   End If
Else
   If upjpg_2.Top <= -Val(upjpg_2.Height) Then
      upjpg_2.Top = -Val(upjpg_2.Height)
      PEASpersontalk.Top = -Val(PEASpersontalk.Height)
      stup.Enabled = False
   Else
      upjpg_2.Top = Val(upjpg_2.Top) - 60
      PEASpersontalk.Top = Val(PEASpersontalk.Top) - 60
   End If
End If
End Sub

Private Sub tr�j�H���ι�_�ϥΪ�_Timer()
Dim bigall As Integer
Dim bigw As Integer
Dim kp As Integer

bigw = �j�H���ι�_�ϥΪ�.�j�H���Ϥ�width / 2
If 2580 - bigw < 0 Or Val(VBEPerson(1, 1, 2, 2, 5)) = 1 Then
    bigall = 0
Else
    bigall = 2580 - bigw
End If

kp = (�j�H���ι�_�ϥΪ�.�j�H���Ϥ�width + bigall) / 30
If sq <= 400 Then
   If �j�H���ι�_�ϥΪ�.Left >= bigall Then
       �j�H���ι�_�ϥΪ�.Left = bigall
       tr�j�H���ι�_�ϥΪ�.Enabled = False
       swq = 0
       PEASpke.Enabled = True
   Else
       If Abs(�j�H���ι�_�ϥΪ�.Left - bigall) < kp Then
          �j�H���ι�_�ϥΪ�.Left = �j�H���ι�_�ϥΪ�.Left + Abs(�j�H���ι�_�ϥΪ�.Left - bigall)
       Else
          �j�H���ι�_�ϥΪ�.Left = �j�H���ι�_�ϥΪ�.Left + kp
       End If
   End If
Else
   If �j�H���ι�_�ϥΪ�.Left <= -�j�H���ι�_�ϥΪ�.�j�H���Ϥ�width Then
       �j�H���ι�_�ϥΪ�.Left = -�j�H���ι�_�ϥΪ�.�j�H���Ϥ�width
       tr�j�H���ι�_�ϥΪ�.Enabled = False
       stup.Enabled = True
       stdown.Enabled = True
   Else
       �j�H���ι�_�ϥΪ�.Left = �j�H���ι�_�ϥΪ�.Left - kp
   End If
End If
End Sub

Private Sub tr�j�H���ι�_�q��_Timer()
Dim kr As Integer, kn As Integer

kn = �j�H���ι�_�q��.�j�H���Ϥ�width
Dim bigwn, bigall As Integer
bigwn = (�j�H���ι�_�q��.�j�H���Ϥ�width / 2)
If 8760 - bigwn > Val(FormMainMode.ScaleWidth) - Val(�j�H���ι�_�q��.�j�H���Ϥ�width) Or Val(VBEPerson(2, 1, 2, 2, 5)) = 1 Then
    bigall = Val(FormMainMode.ScaleWidth) - Val(�j�H���ι�_�q��.�j�H���Ϥ�width)
Else
    bigall = 8760 - bigwn
End If
kr = (Val(FormMainMode.ScaleWidth) - bigall) / 30
If sq <= 400 Then
   If �j�H���ι�_�q��.Left <= bigall Then
       �j�H���ι�_�q��.Left = bigall
       tr�j�H���ι�_�q��.Enabled = False
   Else
       If �j�H���ι�_�q��.Left - bigall < kr Then
           �j�H���ι�_�q��.Left = �j�H���ι�_�q��.Left - (�j�H���ι�_�q��.Left - bigall)
       Else
           �j�H���ι�_�q��.Left = �j�H���ι�_�q��.Left - kr
       End If
   End If
Else
   If �j�H���ι�_�q��.Left >= FormMainMode.ScaleWidth Then
       �j�H���ι�_�q��.Left = FormMainMode.ScaleWidth
       tr�j�H���ι�_�q��.Enabled = False
   Else
       �j�H���ι�_�q��.Left = �j�H���ι�_�q��.Left + kr
   End If
End If
End Sub

Private Sub cardcomtr_Timer()
Dim i As Integer
If sq <= 400 Then
  For i = 3 To 1 Step -1
     If PEAScardcom(i).Visible = True Then
        If i < 3 Then
           If PEAScardcom(i + 1).Visible = True And PEAScardcom(i + 1).Top - PEAScardcom(i).Top >= 4000 Then
               Select Case i
                  Case 2
                     If PEAScardcom(i).Top <= 4000 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                     End If
                     If PEAScardcom(i).Top >= 4080 Then
                         PEAScardcom(i).Top = 4080
                     End If
                  Case 1
                     If PEAScardcom(i).Top <= 3700 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                     End If
                End Select
           ElseIf PEAScardcom(i + 1).Visible = False Or PEAScardcom(i + 1).Top >= 3000 Then
               Select Case i
                  Case 2
                      If PEAScardcom(i).Top <= 4000 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                     End If
                     If PEAScardcom(i).Top >= 4080 Then
                         PEAScardcom(i).Top = 4080
                     End If
                  Case 1
                      If PEAScardcom(i).Top <= 3700 Then
                         PEAScardcom(i).Top = PEAScardcom(i).Top + 200
                      End If
                End Select
'               PEAScardcom(i).Top = PEAScardcom(i).Top + 200
           End If
        Else
           If PEAScardcom(i).Top <= 4400 Then
               PEAScardcom(i).Top = PEAScardcom(i).Top + 200
           End If
           If PEAScardcom(i).Top >= 4440 Then
                PEAScardcom(i).Top = 4440
           End If
        End If
    End If
  Next
  If PEAScardcom(1).Top >= 3720 Then
    PEAScardcom(1).Top = 3720
    cardcomtr.Enabled = False
    tr�j�H���ι�_�q��.Enabled = True
  End If
End If
End Sub

Private Sub cardustr_Timer()
Dim i As Integer
If sq <= 400 Then
  For i = 3 To 1 Step -1
     If PEAScardus(i).Visible = True Then
        If i < 3 Then
           If PEAScardus(i + 1).Visible = True And PEAScardus(i + 1).Top - PEAScardus(i).Top >= 4000 Then
               Select Case i
                  Case 2
                     If PEAScardus(i).Top <= 4000 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                     End If
                     If PEAScardus(i).Top >= 4080 Then
                         PEAScardus(i).Top = 4080
                     End If
                  Case 1
                     If PEAScardus(i).Top <= 3700 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                     End If
                End Select
           ElseIf PEAScardus(i + 1).Visible = False Or PEAScardus(i + 1).Top >= 3000 Then
               Select Case i
                  Case 2
                      If PEAScardus(i).Top <= 4000 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                     End If
                     If PEAScardus(i).Top >= 4080 Then
                         PEAScardus(i).Top = 4080
                     End If
                  Case 1
                      If PEAScardus(i).Top <= 3700 Then
                         PEAScardus(i).Top = PEAScardus(i).Top + 200
                      End If
                End Select
'               cardus(i).Top = cardus(i).Top + 200
           End If
        Else
           If PEAScardus(i).Top <= 4400 Then
               PEAScardus(i).Top = PEAScardus(i).Top + 200
           End If
           If PEAScardus(i).Top >= 4440 Then
                PEAScardus(i).Top = 4440
           End If
        End If
    End If
  Next
  If PEAScardus(1).Top >= 3720 Then
    PEAScardus(1).Top = 3720
    cardustr.Enabled = False
    tr�j�H���ι�_�ϥΪ�.Enabled = True
  End If
End If
End Sub
